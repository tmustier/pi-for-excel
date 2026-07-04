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

async function startMockTarget(responseText = "ok", extraHeaders = {}) {
  const port = await getFreePort();
  const server = http.createServer((_req, res) => {
    res.statusCode = 200;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    for (const [key, value] of Object.entries(extraHeaders)) {
      res.setHeader(key, value);
    }
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

test("proxy default allowlist includes supported OAuth and web search providers", async (t) => {
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
    "platform.claude.com",
    "auth.openai.com",
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

test("proxy keeps its own CORS headers when upstream sends conflicting ones", async (t) => {
  const target = await startMockTarget("upstream-cors", {
    // llama.cpp-style upstream: empty/incorrect CORS headers that would break
    // the browser integration if forwarded verbatim.
    "Access-Control-Allow-Origin": "",
    "Access-Control-Allow-Methods": "GET",
    "Access-Control-Allow-Headers": "x-upstream-only",
    "Access-Control-Max-Age": "0",
    "Vary": "Accept-Encoding",
    "X-Upstream-Custom": "passthrough",
  });
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

  const url = encodeURIComponent(`http://127.0.0.1:${target.port}/v1/models`);
  const response = await fetch(`http://127.0.0.1:${proxy.port}/?url=${url}`, {
    headers: { Origin: ORIGIN },
  });

  assert.equal(response.status, 200);
  assert.equal(await response.text(), "upstream-cors");

  // Our CORS policy must win over upstream values.
  assert.equal(response.headers.get("access-control-allow-origin"), ORIGIN);
  assert.equal(
    response.headers.get("access-control-allow-methods"),
    "GET,POST,PUT,PATCH,DELETE,OPTIONS",
  );
  assert.equal(response.headers.get("access-control-max-age"), "86400");
  assert.equal(response.headers.get("vary"), "Origin");

  // Unrelated upstream headers still pass through.
  assert.equal(response.headers.get("x-upstream-custom"), "passthrough");
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

test("proxy blocks requests without an allowed Origin header", async (t) => {
  const proxy = await startProxy();
  t.after(async () => {
    await proxy.stop();
  });

  const target = encodeURIComponent("https://api.openai.com/");

  const noOrigin = await fetch(`http://127.0.0.1:${proxy.port}/?url=${target}`);
  assert.equal(noOrigin.status, 403);

  const badOrigin = await fetch(`http://127.0.0.1:${proxy.port}/?url=${target}`, {
    headers: { Origin: "https://evil.example.com" },
  });
  assert.equal(badOrigin.status, 403);
});

test("proxy /healthz responds without an Origin header and never proxies", async (t) => {
  const proxy = await startProxy();
  t.after(async () => {
    await proxy.stop();
  });

  const health = await fetch(`http://127.0.0.1:${proxy.port}/healthz`);
  assert.equal(health.status, 200);
  assert.equal(await health.text(), "ok");

  // Query strings (including ?url=) must not turn /healthz into a proxy path.
  const withUrl = await fetch(
    `http://127.0.0.1:${proxy.port}/healthz?url=${encodeURIComponent("https://api.openai.com/")}`,
  );
  assert.equal(withUrl.status, 200);
  assert.equal(await withUrl.text(), "ok");
});

test("proxy boots with ALLOWED_CLIENT_CIDRS, warns, and still serves loopback", async (t) => {
  const proxy = await startProxy({
    ALLOWED_CLIENT_CIDRS: "10.96.0.0/13, 192.168.1.5",
  });
  t.after(async () => {
    await proxy.stop();
  });

  const health = await fetch(`http://127.0.0.1:${proxy.port}/healthz`);
  assert.equal(health.status, 200);

  const { stdout } = proxy.getLogs();
  assert.match(stdout, /WARNING: accepting non-loopback clients from: 10\.96\.0\.0\/13, 192\.168\.1\.5/);
});

test("proxy exits on invalid ALLOWED_CLIENT_CIDRS (fail closed)", async () => {
  const port = await getFreePort();
  const child = spawn(process.execPath, [PROXY_SCRIPT_PATH], {
    env: {
      ...process.env,
      HOST: "127.0.0.1",
      PORT: String(port),
      ALLOWED_ORIGINS: ORIGIN,
      ALLOWED_CLIENT_CIDRS: "0.0.0.0/0",
    },
    stdio: ["ignore", "pipe", "pipe"],
  });

  let stderr = "";
  child.stderr.setEncoding("utf8");
  child.stderr.on("data", (chunk) => {
    stderr += chunk;
  });

  const [code] = await Promise.race([
    once(child, "exit"),
    delay(5000).then(() => {
      child.kill("SIGKILL");
      throw new Error("proxy did not exit on invalid ALLOWED_CLIENT_CIDRS");
    }),
  ]);

  assert.equal(code, 1);
  assert.match(stderr, /Invalid ALLOWED_CLIENT_CIDRS entries: 0\.0\.0\.0\/0/);
});

test("proxy exits when TLS_KEY_PATH/TLS_CERT_PATH are missing files", async () => {
  const port = await getFreePort();
  const child = spawn(process.execPath, [PROXY_SCRIPT_PATH, "--https"], {
    env: {
      ...process.env,
      HOST: "127.0.0.1",
      PORT: String(port),
      ALLOWED_ORIGINS: ORIGIN,
      TLS_KEY_PATH: "/nonexistent/org.key",
      TLS_CERT_PATH: "/nonexistent/org.pem",
    },
    stdio: ["ignore", "pipe", "pipe"],
  });

  let stderr = "";
  child.stderr.setEncoding("utf8");
  child.stderr.on("data", (chunk) => {
    stderr += chunk;
  });

  const [code] = await Promise.race([
    once(child, "exit"),
    delay(5000).then(() => {
      child.kill("SIGKILL");
      throw new Error("proxy did not exit on missing TLS files");
    }),
  ]);

  assert.equal(code, 1);
  assert.match(stderr, /TLS key\/cert not found/);
  assert.match(stderr, /\/nonexistent\/org\.key/);
});

test("configured ALLOWED_TARGET_HOSTS is enforced even with loopback/private overrides", async (t) => {
  const target = await startMockTarget("hello");
  const proxy = await startProxy({
    ALLOWED_TARGET_HOSTS: "api.openai.com",
    ALLOW_LOOPBACK_TARGETS: "1",
    ALLOW_PRIVATE_TARGETS: "1",
  });
  t.after(async () => {
    await proxy.stop();
    await target.stop();
  });

  // Loopback target not in the configured allowlist → blocked despite override flags.
  const blocked = await fetch(
    `http://127.0.0.1:${proxy.port}/?url=${encodeURIComponent(`http://127.0.0.1:${target.port}/`)}`,
    { headers: { Origin: ORIGIN } },
  );
  assert.equal(blocked.status, 403);
  assert.match(await blocked.text(), /blocked_target_not_allowlisted/);
});

test("allowlisted loopback target works with overrides on a configured allowlist", async (t) => {
  const target = await startMockTarget("hello");
  const proxy = await startProxy({
    ALLOWED_TARGET_HOSTS: "127.0.0.1",
    ALLOW_LOOPBACK_TARGETS: "1",
  });
  t.after(async () => {
    await proxy.stop();
    await target.stop();
  });

  const ok = await fetch(
    `http://127.0.0.1:${proxy.port}/?url=${encodeURIComponent(`http://127.0.0.1:${target.port}/`)}`,
    { headers: { Origin: ORIGIN } },
  );
  assert.equal(ok.status, 200);
  assert.equal(await ok.text(), "hello");
});

test("proxy exits on explicit ALLOWED_TARGET_HOSTS with no valid entries (fail closed)", async () => {
  const port = await getFreePort();
  const child = spawn(process.execPath, [PROXY_SCRIPT_PATH], {
    env: {
      ...process.env,
      HOST: "127.0.0.1",
      PORT: String(port),
      ALLOWED_ORIGINS: ORIGIN,
      ALLOWED_TARGET_HOSTS: " ,, ://bad ,",
    },
    stdio: ["ignore", "pipe", "pipe"],
  });

  let stderr = "";
  child.stderr.setEncoding("utf8");
  child.stderr.on("data", (chunk) => {
    stderr += chunk;
  });

  const [code] = await Promise.race([
    once(child, "exit"),
    delay(5000).then(() => {
      child.kill("SIGKILL");
      throw new Error("proxy did not exit on invalid ALLOWED_TARGET_HOSTS");
    }),
  ]);

  assert.equal(code, 1);
  assert.match(stderr, /ALLOWED_TARGET_HOSTS was set but contained no valid host entries/);
});
