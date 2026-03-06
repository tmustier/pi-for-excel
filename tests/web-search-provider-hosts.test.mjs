import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { readTaskpaneConnectSrcTokens } from "./helpers/taskpane-csp.mjs";
import { PROXY_REACHABILITY_TARGET_URL } from "../src/auth/proxy-validation.ts";
import { WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS } from "../src/tools/web-search-config.ts";

function uniqueHosts(hosts) {
  return Array.from(new Set(hosts));
}

async function readProxyAllowlistHosts() {
  const source = await readFile(new URL("../scripts/cors-proxy-server.mjs", import.meta.url), "utf8");
  const allowlistMatch = source.match(/const DEFAULT_ALLOWED_TARGET_HOSTS = new Set\(\[(?<hosts>[\s\S]*?)\]\);/);
  const hostBlock = allowlistMatch?.groups?.hosts;

  if (typeof hostBlock !== "string") {
    throw new Error("Could not locate DEFAULT_ALLOWED_TARGET_HOSTS in proxy script");
  }

  const hosts = Array.from(
    hostBlock.matchAll(/"([^"\n]+)"/g),
    (match) => match[1],
  );

  return new Set(hosts);
}

test("proxy default host allowlist includes all web-search provider hosts", async () => {
  const proxyHosts = await readProxyAllowlistHosts();

  for (const host of uniqueHosts(WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS)) {
    assert.ok(proxyHosts.has(host), `Missing ${host} in proxy DEFAULT_ALLOWED_TARGET_HOSTS`);
  }
});

test("proxy reachability probe target host is allowlisted by default", async () => {
  const proxyHosts = await readProxyAllowlistHosts();
  const probeHost = new URL(PROXY_REACHABILITY_TARGET_URL).hostname.toLowerCase();

  assert.ok(
    proxyHosts.has(probeHost),
    `Proxy reachability target host ${probeHost} must stay in DEFAULT_ALLOWED_TARGET_HOSTS`,
  );
});

test("taskpane CSP connect-src allows all web-search provider hosts", async () => {
  const connectTokens = await readTaskpaneConnectSrcTokens();

  for (const host of uniqueHosts(WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS)) {
    const origin = `https://${host}`;
    assert.ok(connectTokens.has(origin), `Missing ${origin} in /src/taskpane.html CSP connect-src`);
  }
});

