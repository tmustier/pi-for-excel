import assert from "node:assert/strict";
import { spawn, spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import https from "node:https";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { once } from "node:events";
import { setTimeout as delay } from "node:timers/promises";
import { test } from "node:test";

const TOKEN = "test-background-verify-token";
const SCRIPT_PATH = "scripts/background-verify-bridge-server.mjs";
const testCaByPort = new Map();

async function getFreePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.once("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();
      if (!address || typeof address === "string") {
        server.close();
        reject(new Error("Failed to allocate port"));
        return;
      }
      const { port } = address;
      server.close((error) => {
        if (error) reject(error);
        else resolve(port);
      });
    });
  });
}

async function createTempCert() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pi-bg-verify-cert-"));
  const keyPath = path.join(dir, "key.pem");
  const certPath = path.join(dir, "cert.pem");
  const result = spawnSync("openssl", [
    "req",
    "-x509",
    "-newkey",
    "rsa:2048",
    "-nodes",
    "-keyout",
    keyPath,
    "-out",
    certPath,
    "-subj",
    "/CN=localhost",
    "-addext",
    "subjectAltName=DNS:localhost,IP:127.0.0.1",
    "-days",
    "1",
  ], { encoding: "utf8" });

  if (result.error) {
    await rm(dir, { recursive: true, force: true });
    throw result.error;
  }
  if (result.status !== 0) {
    await rm(dir, { recursive: true, force: true });
    throw new Error(`openssl failed\n${result.stdout}\n${result.stderr}`);
  }

  return {
    certPath,
    keyPath,
    cleanup: () => rm(dir, { recursive: true, force: true }),
  };
}

function requestJson(port, requestPath, body, method = body === undefined ? "GET" : "POST") {
  const payload = body === undefined ? undefined : JSON.stringify(body);
  const ca = testCaByPort.get(port);
  assert.ok(ca, `missing test CA for port ${port}`);
  return new Promise((resolve, reject) => {
    const req = https.request(`https://localhost:${port}${requestPath}`, {
      method,
      agent: new https.Agent({ ca }),
      headers: payload
        ? {
            "content-type": "application/json; charset=utf-8",
            "content-length": Buffer.byteLength(payload),
          }
        : undefined,
    }, (res) => {
      let out = "";
      res.setEncoding("utf8");
      res.on("data", (chunk) => { out += chunk; });
      res.on("end", () => {
        try {
          resolve({ statusCode: res.statusCode, body: JSON.parse(out || "{}") });
        } catch (error) {
          reject(error);
        }
      });
    });
    req.on("error", reject);
    if (payload) req.end(payload);
    else req.end();
  });
}

async function startBridge() {
  const port = await getFreePort();
  const cert = await createTempCert();
  testCaByPort.set(port, readFileSync(cert.certPath));

  const child = spawn(process.execPath, [
    SCRIPT_PATH,
    "serve",
    "--port",
    String(port),
    "--token",
    TOKEN,
    "--cert",
    cert.certPath,
    "--key",
    cert.keyPath,
  ], {
    cwd: process.cwd(),
    stdio: ["ignore", "pipe", "pipe"],
  });

  let stdout = "";
  let stderr = "";
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");

  const ready = new Promise((resolve) => {
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
      if (stdout.includes("background-verify] listening")) resolve(undefined);
    });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
  });

  const exitBeforeReady = once(child, "exit").then(([code, signal]) => {
    throw new Error(`bridge exited before ready (code=${String(code)} signal=${String(signal)})\n${stdout}\n${stderr}`);
  });

  await Promise.race([
    ready,
    exitBeforeReady,
    delay(5000).then(() => {
      throw new Error(`bridge start timeout\n${stdout}\n${stderr}`);
    }),
  ]);

  const stop = async () => {
    if (!child.killed) child.kill("SIGTERM");
    await Promise.race([
      once(child, "exit"),
      delay(1000).then(() => {
        if (!child.killed) child.kill("SIGKILL");
      }),
    ]).catch(() => {});
    testCaByPort.delete(port);
    await cert.cleanup();
  };

  return { port, stop };
}

async function registerClient(port, label = "test-client") {
  const response = await requestJson(port, "/client/register", {
    token: TOKEN,
    client: { href: `https://localhost:3141/${label}`, userAgent: "node-test" },
  });
  assert.equal(response.statusCode, 200);
  assert.equal(response.body.ok, true);
  assert.equal(typeof response.body.clientId, "string");
  return response.body.clientId;
}

test("background verify bridge requires token and refuses commands without a taskpane client", async (t) => {
  const bridge = await startBridge();
  t.after(bridge.stop);

  const health = await requestJson(bridge.port, "/health");
  assert.equal(health.statusCode, 200);
  assert.equal(health.body.ok, true);

  const missingToken = await requestJson(bridge.port, "/command", { type: "status" });
  assert.equal(missingToken.statusCode, 400);
  assert.match(String(missingToken.body.error), /Invalid background verification token/u);

  const noClient = await requestJson(bridge.port, "/command", { token: TOKEN, type: "status" });
  assert.equal(noClient.statusCode, 409);
  assert.match(String(noClient.body.error), /No taskpane client connected/u);
});

test("background verify bridge rejects token-bearing GET poll requests", async (t) => {
  const bridge = await startBridge();
  t.after(bridge.stop);
  const clientId = await registerClient(bridge.port);

  const response = await requestJson(bridge.port, `/client/poll?token=${TOKEN}&clientId=${clientId}`);
  assert.equal(response.statusCode, 404);
  assert.match(String(response.body.error), /Not found/u);
});

test("background verify bridge delivers a targeted command through poll and result", async (t) => {
  const bridge = await startBridge();
  t.after(bridge.stop);
  const clientId = await registerClient(bridge.port);

  const commandPromise = requestJson(bridge.port, "/command", {
    token: TOKEN,
    type: "status",
    clientId,
    timeoutMs: 5000,
  });

  const poll = await requestJson(bridge.port, "/client/poll", { token: TOKEN, clientId });
  assert.equal(poll.statusCode, 200);
  assert.equal(poll.body.type, "status");
  assert.equal(typeof poll.body.id, "string");

  const stored = await requestJson(bridge.port, "/client/result", {
    token: TOKEN,
    clientId,
    commandId: poll.body.id,
    ok: true,
    result: { ready: true },
  });
  assert.equal(stored.statusCode, 200);
  assert.equal(stored.body.ok, true);

  const command = await commandPromise;
  assert.equal(command.statusCode, 200);
  assert.deepEqual(command.body, { ok: true, result: { ready: true } });
});

test("background verify bridge rejects untargeted commands when multiple clients are live", async (t) => {
  const bridge = await startBridge();
  t.after(bridge.stop);
  await registerClient(bridge.port, "one");
  await registerClient(bridge.port, "two");

  const response = await requestJson(bridge.port, "/command", { token: TOKEN, type: "status" });
  assert.equal(response.statusCode, 409);
  assert.match(String(response.body.error), /Multiple taskpane clients connected/u);
  assert.equal(response.body.clients.length, 2);
});

test("background verify bridge removes stale poll waiters after timeout", async (t) => {
  const bridge = await startBridge();
  t.after(bridge.stop);
  const clientId = await registerClient(bridge.port);

  const poll = await requestJson(bridge.port, "/client/poll", { token: TOKEN, clientId, timeoutMs: 10 });
  assert.equal(poll.statusCode, 200);
  assert.equal(poll.body.type, "noop");

  const health = await requestJson(bridge.port, "/health");
  assert.equal(health.statusCode, 200);
  assert.equal(health.body.clients[0].activeLongPolls, 0);
});

test("background verify bridge refuses valueless token flags", async () => {
  const port = await getFreePort();
  const child = spawn(process.execPath, [
    SCRIPT_PATH,
    "serve",
    "--port",
    String(port),
    "--token",
  ], {
    cwd: process.cwd(),
    env: { ...process.env, PI_BACKGROUND_VERIFY_TOKEN: "" },
    stdio: ["ignore", "pipe", "pipe"],
  });

  let output = "";
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk) => { output += chunk; });
  child.stderr.on("data", (chunk) => { output += chunk; });

  const [code] = await Promise.race([
    once(child, "exit"),
    delay(5000).then(() => {
      child.kill("SIGKILL");
      throw new Error("bridge did not exit after valueless token");
    }),
  ]);

  assert.equal(code, 2);
  assert.match(output, /without PI_BACKGROUND_VERIFY_TOKEN or --token/u);
});

test("background verify bridge refuses non-loopback bind hosts", async () => {
  const port = await getFreePort();
  const child = spawn(process.execPath, [
    SCRIPT_PATH,
    "serve",
    "--host",
    "0.0.0.0",
    "--port",
    String(port),
    "--token",
    TOKEN,
  ], {
    cwd: process.cwd(),
    stdio: ["ignore", "pipe", "pipe"],
  });

  let output = "";
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk) => { output += chunk; });
  child.stderr.on("data", (chunk) => { output += chunk; });

  const [code] = await Promise.race([
    once(child, "exit"),
    delay(5000).then(() => {
      child.kill("SIGKILL");
      throw new Error("bridge did not exit after invalid host");
    }),
  ]);

  assert.notEqual(code, 0);
  assert.match(output, /loopback only/u);
});
