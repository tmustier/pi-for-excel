import test from "node:test";
import assert from "node:assert/strict";
import { spawn, spawnSync } from "node:child_process";
import net from "node:net";
import { once } from "node:events";
import { setTimeout as delay } from "node:timers/promises";

const ORIGIN = "https://localhost:3000";
const BRIDGE_SCRIPT_PATH = new URL("../scripts/tmux-bridge-server.mjs", import.meta.url).pathname;

function hasTmuxBinary() {
  const result = spawnSync("tmux", ["-V"], {
    encoding: "utf8",
  });

  return !result.error && result.status === 0;
}

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
      server.close((err) => {
        if (err) reject(err);
        else resolve(port);
      });
    });
  });
}

async function startBridge(extraEnv = {}) {
  const port = await getFreePort();

  const child = spawn(process.execPath, [BRIDGE_SCRIPT_PATH], {
    env: {
      ...process.env,
      HOST: "127.0.0.1",
      PORT: String(port),
      ALLOWED_ORIGINS: ORIGIN,
      TMUX_BRIDGE_MODE: "stub",
      ...extraEnv,
    },
    stdio: ["ignore", "pipe", "pipe"],
  });

  let stdout = "";
  let stderr = "";

  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");

  const ready = new Promise((resolve, reject) => {
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
      if (stdout.includes("tmux bridge listening on")) {
        resolve(undefined);
      }
    });

    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });

    child.once("exit", (code, signal) => {
      reject(new Error(`bridge exited before ready (code=${String(code)} signal=${String(signal)})\n${stdout}\n${stderr}`));
    });
  });

  await Promise.race([
    ready,
    delay(5000).then(() => {
      throw new Error(`bridge start timeout\n${stdout}\n${stderr}`);
    }),
  ]);

  const stop = async () => {
    if (!child.killed) {
      child.kill("SIGTERM");
    }

    await Promise.race([
      once(child, "exit"),
      delay(2000).then(() => {
        if (!child.killed) {
          child.kill("SIGKILL");
        }
      }),
    ]).catch(() => {});
  };

  return {
    port,
    stop,
  };
}

function requestInit(method, body, token) {
  const headers = {
    Origin: ORIGIN,
  };

  if (body !== undefined) {
    headers["Content-Type"] = "application/json";
  }

  if (token) {
    headers.Authorization = `Bearer ${token}`;
  }

  return {
    method,
    headers,
    body: body !== undefined ? JSON.stringify(body) : undefined,
  };
}

test("tmux bridge health endpoint responds in stub mode", async (t) => {
  const bridge = await startBridge();
  t.after(async () => {
    await bridge.stop();
  });

  const response = await fetch(`http://127.0.0.1:${bridge.port}/health`, {
    headers: { Origin: ORIGIN },
  });

  assert.equal(response.status, 200);

  const payload = await response.json();
  assert.equal(payload.ok, true);
  assert.equal(payload.mode, "stub");
  assert.equal(payload.backend, "stub");
});

test("tmux mode health endpoint succeeds before any session exists", async (t) => {
  if (!hasTmuxBinary()) {
    t.skip("tmux binary is not available in this environment");
    return;
  }

  const socketPath = `/tmp/pi-tmux-bridge-test-${Date.now()}-${Math.random().toString(16).slice(2)}.sock`;

  const bridge = await startBridge({
    TMUX_BRIDGE_MODE: "tmux",
    TMUX_BRIDGE_SOCKET_PATH: socketPath,
  });

  t.after(async () => {
    await bridge.stop();
  });

  const response = await fetch(`http://127.0.0.1:${bridge.port}/health`, {
    headers: { Origin: ORIGIN },
  });

  assert.equal(response.status, 200);

  const payload = await response.json();
  assert.equal(payload.ok, true);
  assert.equal(payload.mode, "tmux");
  assert.equal(payload.backend, "tmux");
  assert.equal(payload.sessions, 0);
});

test("tmux bridge blocks disallowed origins", async (t) => {
  const bridge = await startBridge();
  t.after(async () => {
    await bridge.stop();
  });

  const response = await fetch(`http://127.0.0.1:${bridge.port}/health`, {
    headers: { Origin: "https://evil.example" },
  });

  assert.equal(response.status, 403);
  const text = await response.text();
  assert.equal(text, "forbidden");
});

test("tmux bridge enforces bearer token when configured", async (t) => {
  const bridge = await startBridge({
    TMUX_BRIDGE_TOKEN: "local-secret",
  });
  t.after(async () => {
    await bridge.stop();
  });

  const unauthorized = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "list_sessions",
  }));
  assert.equal(unauthorized.status, 401);

  const wrongToken = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "list_sessions",
  }, "wrong-token"));
  assert.equal(wrongToken.status, 401);

  const authorized = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "list_sessions",
  }, "local-secret"));
  assert.equal(authorized.status, 200);

  const payload = await authorized.json();
  assert.equal(payload.ok, true);
  assert.equal(payload.action, "list_sessions");
  assert.deepEqual(payload.sessions, []);
});

test("stub mode supports create/list/send/capture/kill lifecycle", async (t) => {
  const bridge = await startBridge();
  t.after(async () => {
    await bridge.stop();
  });

  const createSession = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "create_session",
    session: "demo",
  }));
  assert.equal(createSession.status, 200);

  const listAfterCreate = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "list_sessions",
  }));
  const listPayload = await listAfterCreate.json();
  assert.deepEqual(listPayload.sessions, ["demo"]);

  const send = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "send_keys",
    session: "demo",
    text: "echo hello",
    enter: true,
  }));
  assert.equal(send.status, 200);

  const capture = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "capture_pane",
    session: "demo",
    lines: 50,
  }));
  assert.equal(capture.status, 200);

  const capturePayload = await capture.json();
  assert.equal(capturePayload.action, "capture_pane");
  assert.match(capturePayload.output, /echo hello/);

  const kill = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "kill_session",
    session: "demo",
  }));
  assert.equal(kill.status, 200);

  const listAfterKill = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "list_sessions",
  }));
  const afterKillPayload = await listAfterKill.json();
  assert.deepEqual(afterKillPayload.sessions, []);
});

test("capture_pane wait_ms delays response in stub mode", async (t) => {
  const bridge = await startBridge();
  t.after(async () => {
    await bridge.stop();
  });

  const createSession = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "create_session",
    session: "slow-capture",
  }));
  assert.equal(createSession.status, 200);

  const send = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "send_keys",
    session: "slow-capture",
    text: "echo hello",
    enter: true,
  }));
  assert.equal(send.status, 200);

  const startedAt = Date.now();
  const capture = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "capture_pane",
    session: "slow-capture",
    lines: 20,
    wait_ms: 250,
  }));
  const elapsedMs = Date.now() - startedAt;

  assert.equal(capture.status, 200);
  assert.ok(elapsedMs >= 180, `Expected capture to wait at least ~180ms, got ${String(elapsedMs)}ms`);

  const payload = await capture.json();
  assert.match(payload.output, /echo hello/);
});

test("tmux bridge rejects invalid action payloads", async (t) => {
  const bridge = await startBridge();
  t.after(async () => {
    await bridge.stop();
  });

  const response = await fetch(`http://127.0.0.1:${bridge.port}/v1/tmux`, requestInit("POST", {
    action: "not_real",
  }));

  assert.equal(response.status, 400);
  const payload = await response.json();
  assert.equal(payload.ok, false);
  assert.match(payload.error, /Invalid action/);
});
