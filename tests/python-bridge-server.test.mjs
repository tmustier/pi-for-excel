import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import net from "node:net";
import { once } from "node:events";
import { setTimeout as delay } from "node:timers/promises";

const ORIGIN = "https://localhost:3000";
const BRIDGE_SCRIPT_PATH = new URL("../scripts/python-bridge-server.mjs", import.meta.url).pathname;

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
      PYTHON_BRIDGE_MODE: "stub",
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
      if (stdout.includes("python bridge listening on")) {
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

test("python bridge health endpoint responds in stub mode", async (t) => {
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
  assert.equal(payload.python.available, true);
  assert.equal(payload.libreoffice.available, true);
});

test("python bridge blocks disallowed origins", async (t) => {
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

test("python bridge enforces bearer token when configured", async (t) => {
  const bridge = await startBridge({
    PYTHON_BRIDGE_TOKEN: "local-secret",
  });
  t.after(async () => {
    await bridge.stop();
  });

  const unauthorized = await fetch(
    `http://127.0.0.1:${bridge.port}/v1/python-run`,
    requestInit("POST", { code: "print('hello')" }),
  );
  assert.equal(unauthorized.status, 401);

  const wrongToken = await fetch(
    `http://127.0.0.1:${bridge.port}/v1/python-run`,
    requestInit("POST", { code: "print('hello')" }, "wrong-token"),
  );
  assert.equal(wrongToken.status, 401);

  const authorized = await fetch(
    `http://127.0.0.1:${bridge.port}/v1/python-run`,
    requestInit("POST", { code: "print('hello')" }, "local-secret"),
  );
  assert.equal(authorized.status, 200);

  const payload = await authorized.json();
  assert.equal(payload.ok, true);
  assert.equal(payload.action, "run_python");
});

test("stub mode python endpoint returns deterministic payload", async (t) => {
  const bridge = await startBridge();
  t.after(async () => {
    await bridge.stop();
  });

  const response = await fetch(
    `http://127.0.0.1:${bridge.port}/v1/python-run`,
    requestInit("POST", {
      code: "result = {'value': 1}",
      input_json: "{\"hello\":\"world\"}",
    }),
  );

  assert.equal(response.status, 200);

  const payload = await response.json();
  assert.equal(payload.ok, true);
  assert.equal(payload.action, "run_python");
  assert.equal(payload.exit_code, 0);
  assert.match(payload.stdout, /simulated/i);
  assert.equal(payload.result_json, "{\"hello\":\"world\"}");
});

test("stub mode libreoffice endpoint returns derived output path", async (t) => {
  const bridge = await startBridge();
  t.after(async () => {
    await bridge.stop();
  });

  const response = await fetch(
    `http://127.0.0.1:${bridge.port}/v1/libreoffice-convert`,
    requestInit("POST", {
      input_path: "/tmp/source.xlsx",
      target_format: "csv",
    }),
  );

  assert.equal(response.status, 200);

  const payload = await response.json();
  assert.equal(payload.ok, true);
  assert.equal(payload.action, "convert");
  assert.equal(payload.output_path, "/tmp/source.csv");
  assert.equal(payload.converter, "stub");
});

test("python bridge rejects invalid python payloads", async (t) => {
  const bridge = await startBridge();
  t.after(async () => {
    await bridge.stop();
  });

  const response = await fetch(
    `http://127.0.0.1:${bridge.port}/v1/python-run`,
    requestInit("POST", {
      code: "",
    }),
  );

  assert.equal(response.status, 400);
  const payload = await response.json();
  assert.equal(payload.ok, false);
  assert.match(payload.error, /code is required/i);
});

test("python bridge rejects invalid libreoffice payloads", async (t) => {
  const bridge = await startBridge();
  t.after(async () => {
    await bridge.stop();
  });

  const response = await fetch(
    `http://127.0.0.1:${bridge.port}/v1/libreoffice-convert`,
    requestInit("POST", {
      input_path: "/tmp/source.xlsx",
      target_format: "docx",
    }),
  );

  assert.equal(response.status, 400);
  const payload = await response.json();
  assert.equal(payload.ok, false);
  assert.match(payload.error, /target_format must be one of/i);
});
