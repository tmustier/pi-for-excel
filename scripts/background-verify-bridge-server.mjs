#!/usr/bin/env node
/**
 * Dev-only background verification bridge for real Excel taskpanes.
 *
 * The real pi-for-excel taskpane initiates outbound HTTPS long-poll requests to
 * this loopback server. Agents send commands to the server; the server queues
 * them for the taskpane and returns the taskpane's result. This avoids raw GUI
 * input and lets verification run while Excel stays in the background.
 */

import { randomBytes, randomUUID } from "node:crypto";
import fs from "node:fs";
import https from "node:https";
import os from "node:os";
import path from "node:path";
import { setTimeout as delay } from "node:timers/promises";

const DEFAULT_HOST = "localhost";
const DEFAULT_PORT = 3157;
const DEFAULT_CERT_PATH = "cert.pem";
const DEFAULT_KEY_PATH = "key.pem";
const CLIENT_TTL_MS = 60_000;
const POLL_TIMEOUT_MS = 25_000;
const COMMAND_DEFAULT_TIMEOUT_MS = 30_000;
const COMMAND_MAX_TIMEOUT_MS = 300_000;
const LOOPBACK_HOSTS = new Set(["127.0.0.1", "localhost"]);
const DEV_ORIGINS = new Set([
  "https://localhost:3141",
  "https://127.0.0.1:3141",
]);

function parseArgs(argv) {
  const out = { _: [] };
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg.startsWith("--")) {
      out._.push(arg);
      continue;
    }
    const key = arg.slice(2);
    const next = argv[i + 1];
    if (!next || next.startsWith("--")) {
      out[key] = true;
      continue;
    }
    out[key] = next;
    i += 1;
  }
  return out;
}

function usage() {
  return `Usage:
  # Start server. Use localhost when the taskpane URL is https://localhost:3157.
  PI_BACKGROUND_VERIFY_TOKEN=<token> PI_BACKGROUND_VERIFY_HOST=localhost node scripts/background-verify-bridge-server.mjs serve [--port 3157]

  # Send command to a connected taskpane. Keep PI_BACKGROUND_VERIFY_HOST consistent with the server.
  # If /health lists multiple clients, pass --clientId <id>. Use --timeout <ms> for long prompt runs.
  PI_BACKGROUND_VERIFY_TOKEN=<token> PI_BACKGROUND_VERIFY_HOST=localhost node scripts/background-verify-bridge-server.mjs command status
  PI_BACKGROUND_VERIFY_TOKEN=<token> PI_BACKGROUND_VERIFY_HOST=localhost node scripts/background-verify-bridge-server.mjs command officeProbe
  PI_BACKGROUND_VERIFY_TOKEN=<token> PI_BACKGROUND_VERIFY_HOST=localhost node scripts/background-verify-bridge-server.mjs command readRange '{"address":"Sheet1!A1:B2"}'
  PI_BACKGROUND_VERIFY_TOKEN=<token> PI_BACKGROUND_VERIFY_HOST=localhost node scripts/background-verify-bridge-server.mjs command writeRange '{"address":"Sheet1!A1:B2","values":[["marker",1],["sum","=SUM(B1)"]]}'
  PI_BACKGROUND_VERIFY_TOKEN=<token> PI_BACKGROUND_VERIFY_HOST=localhost node scripts/background-verify-bridge-server.mjs command clearRange '{"address":"Sheet1!A1:B2","applyTo":"contents"}'
  PI_BACKGROUND_VERIFY_TOKEN=<token> PI_BACKGROUND_VERIFY_HOST=localhost node scripts/background-verify-bridge-server.mjs command workbookWriteProbe '{"keepSheet":false}'
  PI_BACKGROUND_VERIFY_TOKEN=<token> PI_BACKGROUND_VERIFY_HOST=localhost node scripts/background-verify-bridge-server.mjs command configureProxy '{"enabled":true,"url":"https://localhost:3003"}'
  PI_BACKGROUND_VERIFY_TOKEN=<token> PI_BACKGROUND_VERIFY_HOST=localhost node scripts/background-verify-bridge-server.mjs command selectModel '{"provider":"openai-codex","modelId":"gpt-5.6-sol"}'
  PI_BACKGROUND_VERIFY_TOKEN=<token> PI_BACKGROUND_VERIFY_HOST=localhost node scripts/background-verify-bridge-server.mjs command submitPrompt '{"text":"Write SMOKE into A1, then tell me what changed","waitForIdle":true}'
  PI_BACKGROUND_VERIFY_TOKEN=<token> PI_BACKGROUND_VERIFY_HOST=localhost node scripts/background-verify-bridge-server.mjs command listCharts
`;
}

function readJsonBody(req, limitBytes = 1024 * 1024) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.setEncoding("utf8");
    req.on("data", (chunk) => {
      body += chunk;
      if (body.length > limitBytes) {
        reject(new Error("Request body too large"));
        req.destroy();
      }
    });
    req.on("end", () => {
      if (!body.trim()) {
        resolve({});
        return;
      }
      try {
        resolve(JSON.parse(body));
      } catch (error) {
        reject(error);
      }
    });
    req.on("error", reject);
  });
}

function sendJson(res, statusCode, value, origin) {
  res.statusCode = statusCode;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");
  if (origin && DEV_ORIGINS.has(origin)) {
    res.setHeader("Access-Control-Allow-Origin", origin);
    res.setHeader("Vary", "Origin");
  }
  res.end(JSON.stringify(value));
}

function assertToken(actual, expected) {
  if (!expected) throw new Error("Server token is not configured");
  if (actual !== expected) throw new Error("Invalid background verification token");
}

function recordString(value, key) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return "";
  const field = value[key];
  return typeof field === "string" ? field : "";
}

function sanitizeClientError(value) {
  const name = recordString(value, "name").slice(0, 120) || "Error";
  const message = (recordString(value, "message") || (typeof value === "string" ? value : "Taskpane command failed")).slice(0, 2_000);
  return { name, message };
}

class BridgeState {
  constructor() {
    this.clients = new Map();
    this.pendingCommands = new Set();
  }

  register(client) {
    const clientId = randomUUID();
    this.clients.set(clientId, {
      clientId,
      client,
      registeredAt: Date.now(),
      lastSeenAt: Date.now(),
      queue: [],
      waiters: [],
      results: new Map(),
    });
    return clientId;
  }

  getClient(clientId) {
    const record = this.clients.get(clientId);
    if (record) record.lastSeenAt = Date.now();
    return record ?? null;
  }

  isLiveClient(client, now = Date.now()) {
    return client.waiters.length > 0 || client.queue.length > 0 || now - client.lastSeenAt <= CLIENT_TTL_MS;
  }

  liveClients() {
    const now = Date.now();
    return Array.from(this.clients.values())
      .filter((client) => this.isLiveClient(client, now))
      .sort((left, right) => {
        const waitingDelta = Number(right.waiters.length > 0) - Number(left.waiters.length > 0);
        if (waitingDelta !== 0) return waitingDelta;
        return right.lastSeenAt - left.lastSeenAt;
      });
  }

  resolveClientForCommand(requestedClientId) {
    if (requestedClientId) {
      const client = this.clients.get(requestedClientId) ?? null;
      if (!client || !this.isLiveClient(client)) {
        return { error: `Unknown or inactive taskpane client: ${requestedClientId}` };
      }
      return { client };
    }

    const clients = this.liveClients();
    if (clients.length === 0) return { error: "No taskpane client connected" };
    if (clients.length > 1) {
      return {
        error: "Multiple taskpane clients connected; pass clientId from /health to target a workbook explicitly",
        clients: clients.map((client) => client.clientId),
      };
    }
    return { client: clients[0] };
  }

  enqueue(client, command) {
    this.pendingCommands.add(command.id);
    client.queue.push(command);
    const waiter = client.waiters.shift();
    if (waiter) waiter();
    return client.clientId;
  }

  closeCommand(commandId) {
    this.pendingCommands.delete(commandId);
    for (const client of this.clients.values()) {
      client.queue = client.queue.filter((command) => command.id !== commandId);
      client.results.delete(commandId);
    }
  }

  commandForClient(client) {
    return client.queue.shift() ?? null;
  }

  async waitForCommand(client, timeoutMs) {
    const existing = this.commandForClient(client);
    if (existing) return existing;

    let waiter = null;
    const waitForWake = new Promise((resolve) => {
      waiter = resolve;
      client.waiters.push(waiter);
    });

    try {
      await Promise.race([waitForWake, delay(timeoutMs)]);
    } finally {
      if (waiter) {
        const index = client.waiters.indexOf(waiter);
        if (index >= 0) client.waiters.splice(index, 1);
      }
    }

    return this.commandForClient(client);
  }

  storeResult(clientId, commandId, payload) {
    if (!this.pendingCommands.has(commandId)) return false;
    const client = this.clients.get(clientId);
    if (!client) return false;
    client.results.set(commandId, payload);
    return true;
  }

  takeResult(commandId) {
    for (const client of this.clients.values()) {
      if (!client.results.has(commandId)) continue;
      const value = client.results.get(commandId);
      client.results.delete(commandId);
      this.pendingCommands.delete(commandId);
      return value;
    }
    return null;
  }

  summary() {
    return {
      pendingCommands: this.pendingCommands.size,
      clients: Array.from(this.clients.values()).map((client) => ({
        clientId: client.clientId,
        client: client.client,
        registeredAt: client.registeredAt,
        lastSeenAt: client.lastSeenAt,
        queuedCommands: client.queue.length,
        activeLongPolls: client.waiters.length,
        pendingResults: client.results.size,
      })),
    };
  }
}

function createServer({ token, host, port, certPath, keyPath }) {
  const state = new BridgeState();
  const tls = {
    key: fs.readFileSync(keyPath),
    cert: fs.readFileSync(certPath),
  };

  const server = https.createServer(tls, async (req, res) => {
    const origin = req.headers.origin;
    if (req.method === "OPTIONS") {
      if (typeof origin === "string" && DEV_ORIGINS.has(origin)) {
        res.setHeader("Access-Control-Allow-Origin", origin);
        res.setHeader("Access-Control-Allow-Headers", "content-type");
        res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
        res.setHeader("Vary", "Origin");
      }
      res.statusCode = 204;
      res.end();
      return;
    }

    try {
      const url = new URL(req.url ?? "/", `https://${host}:${port}`);

      if (req.method === "GET" && url.pathname === "/health") {
        sendJson(res, 200, { ok: true, ...state.summary() }, origin);
        return;
      }

      if (req.method === "POST" && url.pathname === "/client/register") {
        const body = await readJsonBody(req);
        assertToken(body.token, token);
        const clientId = state.register(body.client ?? {});
        sendJson(res, 200, { ok: true, clientId }, origin);
        return;
      }

      if (req.method === "POST" && url.pathname === "/client/poll") {
        const body = await readJsonBody(req);
        assertToken(body.token, token);
        const clientId = typeof body.clientId === "string" ? body.clientId : "";
        const client = state.getClient(clientId);
        if (!client) {
          sendJson(res, 404, { ok: false, error: "Unknown client" }, origin);
          return;
        }
        const requestedPollTimeoutMs = Number(body.timeoutMs ?? POLL_TIMEOUT_MS);
        const pollTimeoutMs = Number.isFinite(requestedPollTimeoutMs)
          ? Math.max(1, Math.min(POLL_TIMEOUT_MS, Math.floor(requestedPollTimeoutMs)))
          : POLL_TIMEOUT_MS;
        const command = await state.waitForCommand(client, pollTimeoutMs);
        sendJson(res, 200, command ?? { type: "noop" }, origin);
        return;
      }

      if (req.method === "POST" && url.pathname === "/client/result") {
        const body = await readJsonBody(req);
        assertToken(body.token, token);
        const clientId = typeof body.clientId === "string" ? body.clientId : "";
        const commandId = typeof body.commandId === "string" ? body.commandId : "";
        if (!clientId || !commandId) throw new Error("clientId and commandId are required");
        const ok = body.ok === true;
        const stored = state.storeResult(clientId, commandId, {
          ok,
          result: body.result,
          error: ok ? undefined : sanitizeClientError(body.error),
        });
        sendJson(res, stored ? 200 : 404, { ok: stored }, origin);
        return;
      }

      if (req.method === "POST" && url.pathname === "/command") {
        const body = await readJsonBody(req);
        assertToken(body.token, token);
        const type = typeof body.type === "string" ? body.type : "";
        if (!type) throw new Error("Command type is required");
        const payload = body.payload ?? {};
        const requestedClientId = typeof body.clientId === "string" ? body.clientId.trim() : "";
        const target = state.resolveClientForCommand(requestedClientId);
        if (!target.client) {
          sendJson(res, 409, { ok: false, error: target.error, clients: target.clients ?? [] }, origin);
          return;
        }
        const commandId = randomUUID();
        const clientId = state.enqueue(target.client, {
          id: commandId,
          type,
          payload,
        });
        const timeoutMs = commandTimeoutMs(type, payload, body.timeoutMs);
        const started = Date.now();
        while (Date.now() - started < timeoutMs) {
          const result = state.takeResult(commandId);
          if (result) {
            sendJson(res, result.ok ? 200 : 500, result, origin);
            return;
          }
          await delay(100);
        }
        state.closeCommand(commandId);
        sendJson(res, 504, { ok: false, error: `Command ${type} timed out after ${timeoutMs}ms`, clientId }, origin);
        return;
      }

      sendJson(res, 404, { ok: false, error: "Not found" }, origin);
    } catch (error) {
      sendJson(res, 400, { ok: false, error: error instanceof Error ? error.message : String(error) }, origin);
    }
  });

  return { server, state };
}

function normalizePort(value) {
  const parsed = Number.parseInt(String(value ?? ""), 10);
  return Number.isInteger(parsed) && parsed > 0 && parsed < 65536 ? parsed : DEFAULT_PORT;
}

function normalizeHost(value) {
  const host = String(value ?? DEFAULT_HOST).trim() || DEFAULT_HOST;
  if (!LOOPBACK_HOSTS.has(host)) {
    throw new Error(`Background verification bridge must bind to loopback only; refused host ${host}`);
  }
  return host;
}

function finiteNumber(value) {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function clampTimeoutMs(value, fallback) {
  const timeoutMs = finiteNumber(value) ?? fallback;
  return Math.max(1_000, Math.min(COMMAND_MAX_TIMEOUT_MS, Math.floor(timeoutMs)));
}

function commandTimeoutMs(type, payload, explicitTimeoutMs) {
  if (finiteNumber(explicitTimeoutMs)) return clampTimeoutMs(explicitTimeoutMs, COMMAND_DEFAULT_TIMEOUT_MS);
  if (type === "submitPrompt") {
    const promptTimeout = payload && typeof payload === "object" ? finiteNumber(payload.timeoutMs) : undefined;
    return clampTimeoutMs((promptTimeout ?? 60_000) + 5_000, COMMAND_DEFAULT_TIMEOUT_MS);
  }
  return COMMAND_DEFAULT_TIMEOUT_MS;
}

function normalizeToken(args) {
  const token = typeof args.token === "string"
    ? args.token
    : (typeof process.env.PI_BACKGROUND_VERIFY_TOKEN === "string" ? process.env.PI_BACKGROUND_VERIFY_TOKEN : "");
  return token.trim();
}

function isUsableToken(token) {
  return token.length >= 16;
}

function existingCaPaths(...candidates) {
  return candidates.filter((candidate) => typeof candidate === "string" && candidate.length > 0 && fs.existsSync(candidate));
}

function defaultMkcertRootPath() {
  if (process.platform === "darwin") {
    return path.join(os.homedir(), "Library", "Application Support", "mkcert", "rootCA.pem");
  }
  return "";
}

function localAgent(caPaths) {
  const ca = caPaths.map((caPath) => fs.readFileSync(caPath));
  return new https.Agent({ ca });
}

function requestJson(url, payload, caPaths) {
  const body = JSON.stringify(payload);
  return new Promise((resolve, reject) => {
    const req = https.request(url, {
      method: "POST",
      agent: localAgent(caPaths),
      headers: {
        "Content-Type": "application/json; charset=utf-8",
        "Content-Length": Buffer.byteLength(body),
      },
    }, (res) => {
      let out = "";
      res.setEncoding("utf8");
      res.on("data", (chunk) => out += chunk);
      res.on("end", () => {
        try {
          const parsed = JSON.parse(out || "{}");
          if ((res.statusCode ?? 500) >= 400) reject(new Error(parsed.error ?? `HTTP ${res.statusCode}`));
          else resolve(parsed);
        } catch (error) {
          reject(error);
        }
      });
    });
    req.on("error", reject);
    req.end(body);
  });
}

async function runServe(args) {
  const token = normalizeToken(args);
  if (!token) {
    console.error("Refusing to start without PI_BACKGROUND_VERIFY_TOKEN or --token.");
    process.exit(2);
  }
  if (!isUsableToken(token)) {
    console.error("Refusing to start with a background verification token shorter than 16 characters.");
    process.exit(2);
  }
  const port = normalizePort(args.port ?? process.env.PI_BACKGROUND_VERIFY_PORT);
  const host = normalizeHost(args.host ?? process.env.PI_BACKGROUND_VERIFY_HOST ?? DEFAULT_HOST);
  const certPath = path.resolve(String(args.cert ?? DEFAULT_CERT_PATH));
  const keyPath = path.resolve(String(args.key ?? DEFAULT_KEY_PATH));
  const { server } = createServer({ token, host, port, certPath, keyPath });
  server.listen(port, host, () => {
    console.log(`[background-verify] listening on https://${host}:${port}`);
    console.log(`[background-verify] taskpane URL: https://${host}:${port}`);
  });
}

async function runCommand(args) {
  const token = normalizeToken(args);
  if (!token) throw new Error("PI_BACKGROUND_VERIFY_TOKEN or --token is required.");
  if (!isUsableToken(token)) throw new Error("Background verification token must be at least 16 characters.");
  const port = normalizePort(args.port ?? process.env.PI_BACKGROUND_VERIFY_PORT);
  const host = normalizeHost(args.host ?? process.env.PI_BACKGROUND_VERIFY_HOST ?? DEFAULT_HOST);
  const type = args._[1];
  if (!type) throw new Error("Command type is required.\n" + usage());
  const payloadRaw = args._[2];
  const payload = payloadRaw ? JSON.parse(payloadRaw) : {};
  const clientId = typeof args.clientId === "string" ? args.clientId : (typeof args["client-id"] === "string" ? args["client-id"] : undefined);
  const certPath = path.resolve(String(args.cert ?? DEFAULT_CERT_PATH));
  const caPaths = existingCaPaths(
    typeof args.ca === "string" ? path.resolve(args.ca) : "",
    typeof process.env.PI_BACKGROUND_VERIFY_CA_PATH === "string" ? path.resolve(process.env.PI_BACKGROUND_VERIFY_CA_PATH) : "",
    defaultMkcertRootPath(),
    certPath,
  );
  if (caPaths.length === 0) {
    throw new Error("No CA/cert file found for bridge TLS verification. Pass --ca or PI_BACKGROUND_VERIFY_CA_PATH.");
  }
  const response = await requestJson(`https://${host}:${port}/command`, {
    token,
    type,
    payload,
    clientId,
    timeoutMs: args.timeout ? Number(args.timeout) : undefined,
  }, caPaths);
  console.log(JSON.stringify(response, null, 2));
}

const args = parseArgs(process.argv.slice(2));
const mode = args._[0] ?? "help";
if (mode === "serve") {
  runServe(args).catch((error) => {
    console.error(error instanceof Error ? error.message : String(error));
    process.exit(1);
  });
} else if (mode === "command") {
  runCommand(args).catch((error) => {
    console.error(error instanceof Error ? error.message : String(error));
    process.exit(1);
  });
} else if (mode === "token") {
  console.log(randomBytes(24).toString("base64url"));
} else {
  console.log(usage());
}
