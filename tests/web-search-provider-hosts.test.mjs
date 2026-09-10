import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import net from "node:net";
import { setTimeout as delay } from "node:timers/promises";
import test from "node:test";

import { readTaskpaneConnectSrcTokens } from "./helpers/taskpane-csp.mjs";
import { PROXY_REACHABILITY_TARGET_URL } from "../src/auth/proxy-validation.ts";
import { WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS } from "../src/tools/web-search-config.ts";

const ORIGIN = "https://localhost:3141";
const PROXY_SCRIPT_PATH = new URL("../scripts/cors-proxy-server.mjs", import.meta.url).pathname;

function uniqueHosts(hosts) {
  return Array.from(new Set(hosts));
}

async function getFreePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.once("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();
      if (!address || typeof address === "string") {
        server.close();
        reject(new Error("Failed to allocate a proxy test port"));
        return;
      }
      server.close((error) => error ? reject(error) : resolve(address.port));
    });
  });
}

async function stopChild(child) {
  if (!child.killed) child.kill("SIGTERM");
  await Promise.race([
    once(child, "exit"),
    delay(2_000).then(() => {
      if (!child.killed) child.kill("SIGKILL");
    }),
  ]).catch(() => {});
}

async function startProxy() {
  const port = await getFreePort();
  const child = spawn(process.execPath, [PROXY_SCRIPT_PATH], {
    env: {
      ...process.env,
      HOST: "127.0.0.1",
      PORT: String(port),
      ALLOWED_ORIGINS: ORIGIN,
      OAUTH_CALLBACK_SERVER: "0",
    },
    stdio: ["ignore", "pipe", "pipe"],
  });
  let stdout = "";
  let stderr = "";
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk) => { stdout += chunk; });
  child.stderr.on("data", (chunk) => { stderr += chunk; });

  await Promise.race([
    new Promise((resolve, reject) => {
      const checkReady = () => {
        if (stdout.includes("CORS proxy listening on") && stdout.includes("Allowed target hosts")) {
          resolve(undefined);
        }
      };
      child.stdout.on("data", checkReady);
      child.once("exit", (code, signal) => {
        reject(new Error(`proxy exited before ready (code=${String(code)} signal=${String(signal)})\n${stdout}\n${stderr}`));
      });
      checkReady();
    }),
    delay(5_000).then(() => {
      throw new Error(`proxy start timeout\n${stdout}\n${stderr}`);
    }),
  ]);

  return { port, stop: () => stopChild(child) };
}

async function requestThroughProxy(port, targetUrl) {
  return fetch(`http://127.0.0.1:${port}/?url=${encodeURIComponent(targetUrl)}`, {
    headers: { Origin: ORIGIN },
    signal: AbortSignal.timeout(15_000),
  });
}

test("proxy behavior permits every web-search provider host and rejects an unlisted host", async (t) => {
  const proxy = await startProxy();
  t.after(async () => proxy.stop());

  for (const host of uniqueHosts(WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS)) {
    const response = await requestThroughProxy(proxy.port, `https://${host}/`);
    const text = await response.text();
    assert.doesNotMatch(
      text,
      /blocked_target_(?:not_allowlisted|loopback|private_ip|invalid_host)/,
      `Expected ${host} to pass target policy (upstream status ${response.status})`,
    );
  }

  const disallowed = await requestThroughProxy(proxy.port, "https://example.com/");
  assert.equal(disallowed.status, 403);
  assert.match(await disallowed.text(), /blocked_target_not_allowlisted/);
});

test("proxy behavior permits its reachability probe target", async (t) => {
  const proxy = await startProxy();
  t.after(async () => proxy.stop());

  const response = await requestThroughProxy(proxy.port, PROXY_REACHABILITY_TARGET_URL);
  const text = await response.text();
  assert.doesNotMatch(
    text,
    /blocked_target_(?:not_allowlisted|loopback|private_ip|invalid_host)/,
    `Expected reachability target to pass proxy policy (upstream status ${response.status})`,
  );
});

// Static contract: taskpane CSP is a shipped policy string; HTTP behavior cannot prove
// that every provider origin remains declared in connect-src.
test("static contract: taskpane CSP connect-src allows all web-search provider hosts", async () => {
  const connectTokens = await readTaskpaneConnectSrcTokens();

  for (const host of uniqueHosts(WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS)) {
    const origin = `https://${host}`;
    assert.ok(connectTokens.has(origin), `Missing ${origin} in /src/taskpane.html CSP connect-src`);
  }
});
