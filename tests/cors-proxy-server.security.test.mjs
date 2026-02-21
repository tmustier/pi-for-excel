import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import http from "node:http";
import net from "node:net";
import { setTimeout as delay } from "node:timers/promises";

const ORIGIN = "https://localhost:3000";
const PROXY_SCRIPT_PATH = new URL("../scripts/cors-proxy-server.mjs", import.meta.url).pathname;

async function getFreePort() {
  return new Promise((resolve, reject) => {
    const srv = net.createServer();
    srv.once("error", reject);
    srv.listen(0, "127.0.0.1", () => {
      const addr = srv.address();
      if (!addr || typeof addr === "string") {
        srv.close();
        reject(new Error("Failed to get free port"));
        return;
      }
      const { port } = addr;
      srv.close((err) => {
        if (err) reject(err);
        else resolve(port);
      });
    });
  });
}

async function startProxy(extraEnv = {}) {
  const port = await getFreePort();

  const child = spawn(process.execPath, [PROXY_SCRIPT_PATH], {
    env: {
      ...process.env,
      HOST: "127.0.0.1",
      PORT: String(port),
      ALLOWED_ORIGINS: ORIGIN,
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
      const hasListeningLog = stdout.includes("CORS proxy listening on");
      const hasTargetPolicyLog =
        stdout.includes("Allowed target hosts")
        || stdout.includes("target host allowlisting disabled");

      if (hasListeningLog && hasTargetPolicyLog) {
        resolve(undefined);
      }
    });

    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });

    child.once("exit", (code, signal) => {
      reject(new Error(`proxy exited before ready (code=${String(code)} signal=${String(signal)})\n${stdout}\n${stderr}`));
    });
  });

  await Promise.race([
    ready,
    delay(5000).then(() => {
      throw new Error(`proxy start timeout\n${stdout}\n${stderr}`);
    }),
  ]);

  const stop = async () => {
    if (!child.killed) {
      child.kill("SIGTERM");
    }

    const exited = Promise.race([
      once(child, "exit"),
      delay(2000).then(() => {
        if (!child.killed) {
          child.kill("SIGKILL");
        }
      }),
    ]);

    await exited.catch(() => {});
  };

  return {
    port,
    stop,
    getLogs: () => ({ stdout, stderr }),
  };
}

async function startMockTarget(responseText = "ok") {
  const port = await getFreePort();
  const server = http.createServer((_req, res) => {
    res.statusCode = 200;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end(responseText);
  });

  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", () => resolve(undefined));
  });

  const stop = async () => {
    await new Promise((resolve) => {
      server.close(() => resolve(undefined));
    });
  };

  return { port, stop };
}

test("proxy blocks loopback targets by default", async (t) => {
  const proxy = await startProxy();
  t.after(async () => {
    await proxy.stop();
  });

  const target = encodeURIComponent("http://127.0.0.1:59999/");
  const response = await fetch(`http://127.0.0.1:${proxy.port}/?url=${target}`, {
    headers: { Origin: ORIGIN },
  });

  assert.equal(response.status, 403);
  const text = await response.text();
  assert.match(text, /blocked_target_loopback/);
});

test("proxy blocks non-allowlisted hosts by default", async (t) => {
  const proxy = await startProxy();
  t.after(async () => {
    await proxy.stop();
  });

  const target = encodeURIComponent("https://example.com/");
  const response = await fetch(`http://127.0.0.1:${proxy.port}/?url=${target}`, {
    headers: { Origin: ORIGIN },
  });

  assert.equal(response.status, 403);
  const text = await response.text();
  assert.match(text, /blocked_target_not_allowlisted/);
});

test("proxy default allowlist includes supported web search providers", async (t) => {
  const proxy = await startProxy();
  t.after(async () => {
    await proxy.stop();
  });

  const { stdout } = proxy.getLogs();
  const allowlistLine = stdout
    .split("\n")
    .find((line) => line.includes("Allowed target hosts"));

  assert.ok(allowlistLine, "Expected startup logs to include default allowed target hosts");

  const loggedHosts = new Set(
    allowlistLine
      .split(":")
      .slice(1)
      .join(":")
      .split(",")
      .map((value) => value.trim())
      .filter(Boolean),
  );

  const requiredHosts = [
    "s.jina.ai",
    "api.firecrawl.dev",
    "google.serper.dev",
    "api.tavily.com",
    "api.search.brave.com",
  ];

  for (const host of requiredHosts) {
    assert.ok(loggedHosts.has(host), `Expected ${host} in default target allowlist log`);
  }
});

test("proxy allows GitHub enterprise OAuth-style endpoints on custom domains by default", async (t) => {
  const proxy = await startProxy();
  t.after(async () => {
    await proxy.stop();
  });

  const target = encodeURIComponent("https://ghe.example.invalid/login/device/code");
  const response = await fetch(`http://127.0.0.1:${proxy.port}/?url=${target}`, {
    headers: { Origin: ORIGIN },
  });

  assert.equal(response.status, 502);
  const text = await response.text();
  assert.match(text, /Proxy error/);
  assert.doesNotMatch(text, /blocked_target_not_allowlisted/);
});

test("explicit ALLOWED_TARGET_HOSTS keeps enterprise-path host checks strict", async (t) => {
  const proxy = await startProxy({
    ALLOWED_TARGET_HOSTS: "api.openai.com",
  });
  t.after(async () => {
    await proxy.stop();
  });

  const target = encodeURIComponent("https://ghe.example.invalid/login/device/code");
  const response = await fetch(`http://127.0.0.1:${proxy.port}/?url=${target}`, {
    headers: { Origin: ORIGIN },
  });

  assert.equal(response.status, 403);
  const text = await response.text();
  assert.match(text, /blocked_target_not_allowlisted/);
});

test("proxy can allow local targets with explicit overrides", async (t) => {
  const target = await startMockTarget("hello-from-local");
  t.after(async () => {
    await target.stop();
  });

  const proxy = await startProxy({
    ALLOW_LOOPBACK_TARGETS: "1",
    ALLOW_PRIVATE_TARGETS: "1",
  });
  t.after(async () => {
    await proxy.stop();
  });

  const url = encodeURIComponent(`http://127.0.0.1:${target.port}/health`);
  const response = await fetch(`http://127.0.0.1:${proxy.port}/?url=${url}`, {
    headers: { Origin: ORIGIN },
  });

  assert.equal(response.status, 200);
  const text = await response.text();
  assert.equal(text, "hello-from-local");
});

test("proxy enforces ALLOWED_TARGET_HOSTS when configured", async (t) => {
  const proxy = await startProxy({
    ALLOWED_TARGET_HOSTS: "api.openai.com",
  });
  t.after(async () => {
    await proxy.stop();
  });

  const blocked = encodeURIComponent("https://example.com/");
  const response = await fetch(`http://127.0.0.1:${proxy.port}/?url=${blocked}`, {
    headers: { Origin: ORIGIN },
  });

  assert.equal(response.status, 403);
  const text = await response.text();
  assert.match(text, /blocked_target_not_allowlisted/);
});
