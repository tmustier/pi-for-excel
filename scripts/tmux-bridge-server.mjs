#!/usr/bin/env node

/**
 * Local tmux bridge for Pi for Excel.
 *
 * Modes:
 * - stub (default): in-memory session simulator for local development/testing.
 * - tmux: real tmux subprocess backend (guarded; no shell interpolation).
 *
 * Endpoints:
 * - GET  /health
 * - POST /v1/tmux
 */

import http from "node:http";
import https from "node:https";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { spawn, spawnSync } from "node:child_process";
import { randomUUID, timingSafeEqual } from "node:crypto";
import { setTimeout as delay } from "node:timers/promises";

const args = new Set(process.argv.slice(2));
const useHttps = args.has("--https") || process.env.HTTPS === "1" || process.env.HTTPS === "true";
const useHttp = args.has("--http");

if (useHttps && useHttp) {
  console.error("[pi-for-excel] Invalid args: can't use both --https and --http");
  process.exit(1);
}

const HOST = process.env.HOST || (useHttps ? "localhost" : "127.0.0.1");
const PORT = Number.parseInt(process.env.PORT || "3341", 10);

const MODE_RAW = (process.env.TMUX_BRIDGE_MODE || "stub").trim().toLowerCase();
const MODE = MODE_RAW === "tmux" ? "tmux" : MODE_RAW === "stub" ? "stub" : null;
if (!MODE) {
  console.error(`[pi-for-excel] Invalid TMUX_BRIDGE_MODE: ${MODE_RAW}. Use "stub" or "tmux".`);
  process.exit(1);
}

function resolveOptionalEnvPath(name) {
  const raw = process.env[name];
  if (typeof raw !== "string") return null;

  const trimmed = raw.trim();
  if (trimmed.length === 0) return null;

  return path.resolve(trimmed);
}

const certDir = resolveOptionalEnvPath("PI_FOR_EXCEL_CERT_DIR") ?? path.resolve(process.cwd());
const keyPath = resolveOptionalEnvPath("PI_FOR_EXCEL_KEY_PATH") ?? path.join(certDir, "key.pem");
const certPath = resolveOptionalEnvPath("PI_FOR_EXCEL_CERT_PATH") ?? path.join(certDir, "cert.pem");

const DEFAULT_ALLOWED_ORIGINS = new Set([
  "https://localhost:3000",
  "https://pi-for-excel.vercel.app",
]);

const MAX_JSON_BODY_BYTES = 256 * 1024;
const MAX_TEXT_LENGTH = 8000;
const MAX_WAIT_FOR_LENGTH = 256;
const MAX_KEY_TOKEN_LENGTH = 48;
const DEFAULT_CAPTURE_LINES = 200;
const DEFAULT_WAIT_TIMEOUT_MS = 5000;
const MIN_CAPTURE_LINES = 1;
const MAX_CAPTURE_LINES = 5000;
const MIN_WAIT_TIMEOUT_MS = 100;
const MAX_WAIT_TIMEOUT_MS = 120_000;
const MIN_CAPTURE_WAIT_MS = 0;
const MAX_CAPTURE_WAIT_MS = 120_000;
const STUB_WAIT_FOR_POLL_INTERVAL_MS = 100;
const REAL_WAIT_FOR_POLL_INTERVAL_MS = 120;

const SESSION_NAME_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,63}$/;
const KEY_TOKEN_PATTERN = /^[A-Za-z0-9._:+-]{1,48}$/;

const TMUX_ACTIONS = [
  "list_sessions",
  "create_session",
  "send_keys",
  "capture_pane",
  "send_and_capture",
  "kill_session",
];
const TMUX_ACTION_SET = new Set(TMUX_ACTIONS);

const allowedOrigins = (() => {
  const raw = process.env.ALLOWED_ORIGINS;
  if (!raw) return DEFAULT_ALLOWED_ORIGINS;

  const custom = new Set(
    raw
      .split(",")
      .map((value) => value.trim())
      .filter(Boolean),
  );

  return custom.size > 0 ? custom : DEFAULT_ALLOWED_ORIGINS;
})();

const authToken = (() => {
  const raw = process.env.TMUX_BRIDGE_TOKEN;
  if (typeof raw !== "string") return "";
  return raw.trim();
})();

const tmuxCommandTimeoutMs = (() => {
  const raw = Number.parseInt(process.env.TMUX_BRIDGE_COMMAND_TIMEOUT_MS || "10000", 10);
  if (!Number.isFinite(raw) || raw < 500) return 10_000;
  return Math.min(raw, 120_000);
})();

const socketDir = path.resolve(
  process.env.TMUX_BRIDGE_SOCKET_DIR || path.join(os.tmpdir(), "pi-for-excel-tmux-bridge"),
);
const socketPath = path.resolve(
  process.env.TMUX_BRIDGE_SOCKET_PATH || path.join(socketDir, "tmux.sock"),
);

class HttpError extends Error {
  /**
   * @param {number} status
   * @param {string} message
   */
  constructor(status, message) {
    super("HttpError");
    this.name = "HttpError";
    this.status = status;
    this.clientMessage = message;
  }
}

function isRecord(value) {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isAllowedOrigin(origin) {
  return typeof origin === "string" && allowedOrigins.has(origin);
}

function isLoopbackAddress(addr) {
  if (!addr) return false;
  if (addr === "::1" || addr === "0:0:0:0:0:0:0:1") return true;
  if (addr.startsWith("127.")) return true;
  if (addr.startsWith("::ffff:127.")) return true;
  return false;
}

function setCorsHeaders(req, res) {
  const origin = req.headers.origin;
  if (isAllowedOrigin(origin)) {
    res.setHeader("Access-Control-Allow-Origin", origin);
    res.setHeader("Vary", "Origin");
  }

  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader(
    "Access-Control-Allow-Headers",
    req.headers["access-control-request-headers"] || "content-type,authorization",
  );
  res.setHeader("Access-Control-Max-Age", "86400");
}

function respondJson(res, status, payload) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(JSON.stringify(payload));
}

function respondText(res, status, text) {
  res.statusCode = status;
  res.setHeader("Content-Type", "text/plain; charset=utf-8");
  res.end(text);
}

function extractBearerToken(headerValue) {
  if (typeof headerValue !== "string") return null;
  const prefix = "Bearer ";
  if (!headerValue.startsWith(prefix)) return null;
  const token = headerValue.slice(prefix.length).trim();
  return token.length > 0 ? token : null;
}

function secureEquals(left, right) {
  const leftBuffer = Buffer.from(left);
  const rightBuffer = Buffer.from(right);
  if (leftBuffer.length !== rightBuffer.length) return false;
  return timingSafeEqual(leftBuffer, rightBuffer);
}

function isAuthorized(req) {
  if (!authToken) return true;

  const candidate = extractBearerToken(req.headers.authorization);
  if (!candidate) return false;

  return secureEquals(candidate, authToken);
}

async function readJsonBody(req) {
  const chunks = [];
  let size = 0;

  for await (const chunk of req) {
    const part = typeof chunk === "string" ? Buffer.from(chunk) : chunk;
    size += part.length;

    if (size > MAX_JSON_BODY_BYTES) {
      throw new HttpError(413, `Request body too large (max ${MAX_JSON_BODY_BYTES} bytes).`);
    }

    chunks.push(part);
  }

  const text = Buffer.concat(chunks).toString("utf8").trim();
  if (text.length === 0) {
    throw new HttpError(400, "Missing JSON request body.");
  }

  try {
    return JSON.parse(text);
  } catch {
    throw new HttpError(400, "Invalid JSON body.");
  }
}

function normalizeOptionalString(value) {
  if (typeof value !== "string") return undefined;

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function parseBoundedInteger(value, options) {
  if (value === undefined) return options.defaultValue;
  if (typeof value !== "number" || !Number.isInteger(value)) {
    throw new HttpError(400, `${options.name} must be an integer.`);
  }
  if (value < options.min || value > options.max) {
    throw new HttpError(400, `${options.name} must be between ${options.min} and ${options.max}.`);
  }
  return value;
}

function parseOptionalBoundedInteger(value, options) {
  if (value === undefined) return undefined;
  if (typeof value !== "number" || !Number.isInteger(value)) {
    throw new HttpError(400, `${options.name} must be an integer.`);
  }
  if (value < options.min || value > options.max) {
    throw new HttpError(400, `${options.name} must be between ${options.min} and ${options.max}.`);
  }
  return value;
}

function normalizeSessionName(value, options = {}) {
  const session = normalizeOptionalString(value);
  if (!session) {
    if (options.required) {
      throw new HttpError(400, `session is required for ${options.action || "this action"}`);
    }
    return undefined;
  }

  if (!SESSION_NAME_PATTERN.test(session)) {
    throw new HttpError(
      400,
      "Invalid session name. Use 1-64 chars: letters, numbers, ., _, :, -",
    );
  }

  return session;
}

function normalizeCwd(value) {
  const cwd = normalizeOptionalString(value);
  if (!cwd) return undefined;

  if (!path.isAbsolute(cwd)) {
    throw new HttpError(400, "cwd must be an absolute path.");
  }

  let stats;
  try {
    stats = fs.statSync(cwd);
  } catch {
    throw new HttpError(400, `cwd does not exist: ${cwd}`);
  }

  if (!stats.isDirectory()) {
    throw new HttpError(400, `cwd is not a directory: ${cwd}`);
  }

  return cwd;
}

function normalizeKeys(value) {
  if (value === undefined) return undefined;
  if (!Array.isArray(value)) {
    throw new HttpError(400, "keys must be an array of strings.");
  }

  const out = [];

  for (const item of value) {
    if (typeof item !== "string") {
      throw new HttpError(400, "keys must contain only strings.");
    }

    const key = item.trim();
    if (key.length === 0) continue;

    if (key.length > MAX_KEY_TOKEN_LENGTH || !KEY_TOKEN_PATTERN.test(key)) {
      throw new HttpError(400, `Invalid key token: ${key}`);
    }

    out.push(key);
  }

  return out.length > 0 ? out : undefined;
}

function normalizeText(value) {
  const text = normalizeOptionalString(value);
  if (!text) return undefined;

  if (text.length > MAX_TEXT_LENGTH) {
    throw new HttpError(400, `text is too long (max ${MAX_TEXT_LENGTH} characters).`);
  }

  return text;
}

function normalizeWaitFor(value) {
  const waitFor = normalizeOptionalString(value);
  if (!waitFor) return undefined;

  if (waitFor.length > MAX_WAIT_FOR_LENGTH) {
    throw new HttpError(400, `wait_for is too long (max ${MAX_WAIT_FOR_LENGTH} characters).`);
  }

  try {
    void new RegExp(waitFor, "m");
  } catch {
    throw new HttpError(400, "wait_for must be a valid regular expression.");
  }

  return waitFor;
}

function parseTmuxRequest(payload) {
  if (!isRecord(payload)) {
    throw new HttpError(400, "Request body must be a JSON object.");
  }

  const action = normalizeOptionalString(payload.action);
  if (!action || !TMUX_ACTION_SET.has(action)) {
    throw new HttpError(400, "Invalid action.");
  }

  const request = {
    action,
    session: normalizeSessionName(payload.session),
    cwd: normalizeCwd(payload.cwd),
    text: normalizeText(payload.text),
    keys: normalizeKeys(payload.keys),
    enter: typeof payload.enter === "boolean" ? payload.enter : undefined,
    lines: parseBoundedInteger(payload.lines, {
      name: "lines",
      min: MIN_CAPTURE_LINES,
      max: MAX_CAPTURE_LINES,
      defaultValue: DEFAULT_CAPTURE_LINES,
    }),
    wait_for: normalizeWaitFor(payload.wait_for),
    timeout_ms: parseBoundedInteger(payload.timeout_ms, {
      name: "timeout_ms",
      min: MIN_WAIT_TIMEOUT_MS,
      max: MAX_WAIT_TIMEOUT_MS,
      defaultValue: DEFAULT_WAIT_TIMEOUT_MS,
    }),
    wait_ms: parseOptionalBoundedInteger(payload.wait_ms, {
      name: "wait_ms",
      min: MIN_CAPTURE_WAIT_MS,
      max: MAX_CAPTURE_WAIT_MS,
    }),
    join_wrapped: typeof payload.join_wrapped === "boolean" ? payload.join_wrapped : undefined,
  };

  switch (action) {
    case "list_sessions":
      return request;

    case "create_session":
      return request;

    case "capture_pane":
    case "kill_session": {
      request.session = normalizeSessionName(payload.session, {
        required: true,
        action,
      });
      return request;
    }

    case "send_keys":
    case "send_and_capture": {
      request.session = normalizeSessionName(payload.session, {
        required: true,
        action,
      });

      const hasInput = Boolean(request.text) || Boolean(request.keys && request.keys.length > 0) || request.enter === true;
      if (!hasInput) {
        throw new HttpError(400, `${action} requires at least one of: text, keys, or enter=true.`);
      }

      return request;
    }

    default:
      throw new HttpError(400, "Invalid action.");
  }
}

async function maybeDelayCapture(waitMs) {
  if (typeof waitMs === "number" && waitMs > 0) {
    await delay(waitMs);
  }
}

async function captureWithOptionalWaitFor(options) {
  const {
    waitFor,
    timeoutMs,
    pollIntervalMs,
    capture,
  } = options;

  if (!waitFor) {
    return capture();
  }

  const matcher = new RegExp(waitFor, "m");
  const startedAt = Date.now();

  while (Date.now() - startedAt <= timeoutMs) {
    const output = await capture();
    if (matcher.test(output)) {
      return output;
    }

    await delay(pollIntervalMs);
  }

  throw new HttpError(408, `wait_for regex did not match before timeout (${timeoutMs}ms).`);
}

function createStubBackend() {
  const sessions = new Map();

  function ensureSession(sessionName) {
    const session = sessions.get(sessionName);
    if (!session) {
      throw new HttpError(404, `Session not found: ${sessionName}`);
    }
    return session;
  }

  function appendLine(sessionName, line) {
    const session = ensureSession(sessionName);
    session.lines.push(line);

    if (session.lines.length > MAX_CAPTURE_LINES) {
      session.lines.splice(0, session.lines.length - MAX_CAPTURE_LINES);
    }
  }

  function captureLines(sessionName, lines) {
    const session = ensureSession(sessionName);
    return session.lines.slice(-lines).join("\n");
  }

  async function applySendInput(request) {
    const fragments = [];

    if (request.text) fragments.push(request.text);
    if (request.keys) fragments.push(...request.keys.map((token) => `<${token}>`));
    if (request.enter) fragments.push("<Enter>");

    if (fragments.length > 0) {
      appendLine(request.session, fragments.join(" "));
    }
  }

  async function handleSendAndCapture(request) {
    await applySendInput(request);

    const timeoutMs = request.timeout_ms ?? DEFAULT_WAIT_TIMEOUT_MS;
    const lines = request.lines ?? DEFAULT_CAPTURE_LINES;

    await maybeDelayCapture(request.wait_ms);

    const output = await captureWithOptionalWaitFor({
      waitFor: request.wait_for,
      timeoutMs,
      pollIntervalMs: STUB_WAIT_FOR_POLL_INTERVAL_MS,
      capture: async () => captureLines(request.session, lines),
    });

    return {
      ok: true,
      action: "send_and_capture",
      session: request.session,
      output,
    };
  }

  return {
    mode: "stub",
    async health() {
      return {
        backend: "stub",
        sessions: sessions.size,
      };
    },
    async handle(request) {
      switch (request.action) {
        case "list_sessions": {
          return {
            ok: true,
            action: "list_sessions",
            sessions: Array.from(sessions.keys()).sort(),
          };
        }

        case "create_session": {
          const sessionName = request.session || `pi-${randomUUID().slice(0, 8)}`;
          if (sessions.has(sessionName)) {
            throw new HttpError(409, `Session already exists: ${sessionName}`);
          }

          sessions.set(sessionName, {
            createdAt: Date.now(),
            cwd: request.cwd,
            lines: [],
          });

          return {
            ok: true,
            action: "create_session",
            session: sessionName,
          };
        }

        case "send_keys": {
          await applySendInput(request);
          return {
            ok: true,
            action: "send_keys",
            session: request.session,
          };
        }

        case "capture_pane": {
          await maybeDelayCapture(request.wait_ms);

          return {
            ok: true,
            action: "capture_pane",
            session: request.session,
            output: captureLines(request.session, request.lines ?? DEFAULT_CAPTURE_LINES),
          };
        }

        case "send_and_capture":
          return handleSendAndCapture(request);

        case "kill_session": {
          ensureSession(request.session);
          sessions.delete(request.session);
          return {
            ok: true,
            action: "kill_session",
            session: request.session,
          };
        }

        default:
          throw new HttpError(400, "Invalid action.");
      }
    },
  };
}

function ensureSocketDirectory() {
  fs.mkdirSync(path.dirname(socketPath), {
    recursive: true,
    mode: 0o700,
  });

  try {
    fs.chmodSync(path.dirname(socketPath), 0o700);
  } catch {
    // ignore chmod failures on platforms/filesystems that do not fully support it
  }
}

function probeTmuxBinary() {
  const probe = spawnSync("tmux", ["-V"], {
    encoding: "utf8",
  });

  if (probe.error) {
    throw new Error(`tmux binary not available: ${probe.error.message}`);
  }

  if (probe.status !== 0) {
    const stderr = typeof probe.stderr === "string" ? probe.stderr.trim() : "";
    throw new Error(stderr || `tmux -V failed with code ${String(probe.status)}`);
  }

  const stdout = typeof probe.stdout === "string" ? probe.stdout.trim() : "";
  return stdout.length > 0 ? stdout : "tmux";
}

async function runTmuxCommand(commandArgs, options = {}) {
  const argsWithSocket = [
    "-f",
    "/dev/null",
    "-S",
    socketPath,
    ...commandArgs,
  ];

  let timedOut = false;

  const result = await new Promise((resolve, reject) => {
    const child = spawn("tmux", argsWithSocket, {
      stdio: ["ignore", "pipe", "pipe"],
    });

    let stdout = "";
    let stderr = "";

    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");

    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });

    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });

    const timeoutId = setTimeout(() => {
      timedOut = true;
      child.kill("SIGKILL");
    }, tmuxCommandTimeoutMs);

    child.once("error", (error) => {
      clearTimeout(timeoutId);
      reject(error);
    });

    child.once("close", (code, signal) => {
      clearTimeout(timeoutId);
      resolve({
        code: code ?? -1,
        signal: signal ?? null,
        stdout,
        stderr,
      });
    });
  });

  if (timedOut) {
    throw new HttpError(504, `tmux command timed out after ${tmuxCommandTimeoutMs}ms.`);
  }

  const message = [result.stderr, result.stdout]
    .map((value) => value.trim())
    .find((value) => value.length > 0) || `tmux command failed with code ${result.code}`;

  if (result.code !== 0) {
    if (
      options.allowNoServer && (
        /no server running on/i.test(message)
        || /error connecting to .*\((No such file or directory|Connection refused)\)/i.test(message)
        || /failed to connect to server/i.test(message)
      )
    ) {
      return result;
    }

    if (/can't find session/i.test(message)) {
      throw new HttpError(404, message);
    }

    if (/duplicate session/i.test(message)) {
      throw new HttpError(409, message);
    }

    throw new HttpError(400, message);
  }

  if (result.signal) {
    throw new HttpError(500, `tmux process exited with signal ${result.signal}`);
  }

  return result;
}

function normalizeCaptureOutput(output) {
  return output.replace(/\r/g, "").trimEnd();
}

async function captureRealPane(sessionName, options) {
  const lines = options.lines ?? DEFAULT_CAPTURE_LINES;
  const joinWrapped = options.join_wrapped === true;

  const args = ["capture-pane", "-p"];
  if (joinWrapped) args.push("-J");
  args.push("-t", sessionName, "-S", `-${lines}`);

  const result = await runTmuxCommand(args);
  return normalizeCaptureOutput(result.stdout);
}

async function sendRealInput(request) {
  const target = request.session;

  if (request.text) {
    await runTmuxCommand(["send-keys", "-t", target, "-l", request.text]);
  }

  if (request.keys) {
    for (const key of request.keys) {
      await runTmuxCommand(["send-keys", "-t", target, key]);
    }
  }

  if (request.enter) {
    await runTmuxCommand(["send-keys", "-t", target, "Enter"]);
  }
}

function createRealTmuxBackend() {
  ensureSocketDirectory();
  const tmuxVersion = probeTmuxBinary();

  async function listSessions() {
    const result = await runTmuxCommand(["list-sessions", "-F", "#{session_name}"], {
      allowNoServer: true,
    });

    if (result.code !== 0) {
      return [];
    }

    return result.stdout
      .split("\n")
      .map((line) => line.trim())
      .filter((line) => line.length > 0)
      .sort();
  }

  async function createSession(request) {
    const sessionName = request.session || `pi-${randomUUID().slice(0, 8)}`;
    const args = ["new-session", "-d", "-s", sessionName];

    if (request.cwd) {
      args.push("-c", request.cwd);
    }

    await runTmuxCommand(args);

    return {
      ok: true,
      action: "create_session",
      session: sessionName,
    };
  }

  async function sendAndCapture(request) {
    await sendRealInput(request);

    const timeoutMs = request.timeout_ms ?? DEFAULT_WAIT_TIMEOUT_MS;

    await maybeDelayCapture(request.wait_ms);

    const output = await captureWithOptionalWaitFor({
      waitFor: request.wait_for,
      timeoutMs,
      pollIntervalMs: REAL_WAIT_FOR_POLL_INTERVAL_MS,
      capture: async () => captureRealPane(request.session, request),
    });

    return {
      ok: true,
      action: "send_and_capture",
      session: request.session,
      output,
    };
  }

  return {
    mode: "tmux",
    async health() {
      return {
        backend: "tmux",
        tmuxVersion,
        socketPath,
        sessions: (await listSessions()).length,
      };
    },
    async handle(request) {
      switch (request.action) {
        case "list_sessions": {
          return {
            ok: true,
            action: "list_sessions",
            sessions: await listSessions(),
          };
        }

        case "create_session":
          return createSession(request);

        case "send_keys": {
          await sendRealInput(request);
          return {
            ok: true,
            action: "send_keys",
            session: request.session,
          };
        }

        case "capture_pane": {
          await maybeDelayCapture(request.wait_ms);

          return {
            ok: true,
            action: "capture_pane",
            session: request.session,
            output: await captureRealPane(request.session, request),
          };
        }

        case "send_and_capture":
          return sendAndCapture(request);

        case "kill_session": {
          await runTmuxCommand(["kill-session", "-t", request.session]);
          return {
            ok: true,
            action: "kill_session",
            session: request.session,
          };
        }

        default:
          throw new HttpError(400, "Invalid action.");
      }
    },
  };
}

const backend = (() => {
  try {
    return MODE === "tmux" ? createRealTmuxBackend() : createStubBackend();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`[pi-for-excel] Failed to initialize tmux backend: ${message}`);
    console.error(
      "[pi-for-excel] Install tmux (for example: brew install tmux), " +
      "or run TMUX_BRIDGE_MODE=stub for simulated mode.",
    );
    process.exit(1);
  }
})();

const handler = async (req, res) => {
  try {
    const remote = req.socket?.remoteAddress;
    if (!isLoopbackAddress(remote)) {
      respondText(res, 403, "forbidden");
      return;
    }

    const origin = req.headers.origin;
    if (!isAllowedOrigin(origin)) {
      respondText(res, 403, "forbidden");
      return;
    }

    setCorsHeaders(req, res);

    if (req.method === "OPTIONS") {
      res.statusCode = 204;
      res.end();
      return;
    }

    const rawUrl = req.url || "/";
    const url = new URL(rawUrl, `http://${HOST}:${PORT}`);

    if (url.pathname === "/health") {
      if (req.method !== "GET") {
        throw new HttpError(405, "Method not allowed.");
      }

      respondJson(res, 200, {
        ok: true,
        mode: backend.mode,
        ...await backend.health(),
      });
      return;
    }

    if (url.pathname === "/v1/tmux") {
      if (req.method !== "POST") {
        throw new HttpError(405, "Method not allowed.");
      }

      if (!isAuthorized(req)) {
        throw new HttpError(401, "Unauthorized.");
      }

      const payload = await readJsonBody(req);
      const request = parseTmuxRequest(payload);
      const result = await backend.handle(request);

      respondJson(res, 200, {
        ok: true,
        ...result,
      });
      return;
    }

    throw new HttpError(404, "Not found.");
  } catch (error) {
    if (error instanceof HttpError) {
      respondJson(res, error.status, {
        ok: false,
        error: error.clientMessage,
      });
      return;
    }

    const detail = error instanceof Error
      ? (typeof error.stack === "string" && error.stack.length > 0 ? error.stack : error.message)
      : String(error);

    console.error(`[pi-for-excel] tmux bridge internal error: ${detail}`);

    respondJson(res, 500, {
      ok: false,
      error: "Internal server error.",
    });
  }
};

const server = (() => {
  if (!useHttps) {
    return http.createServer(handler);
  }

  if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
    console.error("[pi-for-excel] HTTPS requested but key.pem/cert.pem not found in repo root.");
    console.error("Generate them with mkcert (see README). Example: mkcert localhost");
    process.exit(1);
  }

  return https.createServer(
    {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath),
    },
    handler,
  );
})();

server.listen(PORT, HOST, () => {
  const scheme = useHttps ? "https" : "http";
  console.log(`[pi-for-excel] tmux bridge listening on ${scheme}://${HOST}:${PORT}`);
  console.log(`[pi-for-excel] mode: ${backend.mode}`);
  console.log(`[pi-for-excel] health: ${scheme}://${HOST}:${PORT}/health`);
  console.log(`[pi-for-excel] endpoint: ${scheme}://${HOST}:${PORT}/v1/tmux`);
  console.log(`[pi-for-excel] allowed origins: ${Array.from(allowedOrigins).join(", ")}`);

  if (authToken) {
    console.log("[pi-for-excel] auth: bearer token required for POST /v1/tmux");
  }

  if (backend.mode === "tmux") {
    console.log(`[pi-for-excel] tmux socket: ${socketPath}`);
  } else {
    console.log("[pi-for-excel] stub mode: commands are simulated and not executed in a real shell.");
    console.log("[pi-for-excel] use TMUX_BRIDGE_MODE=tmux for real command output.");
  }
});
