import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS } from "../src/tools/web-search-config.ts";

function isRecord(value) {
  return typeof value === "object" && value !== null;
}

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

async function readTaskpaneConnectSrcTokens() {
  const raw = await readFile(new URL("../vercel.json", import.meta.url), "utf8");
  const parsed = JSON.parse(raw);
  if (!isRecord(parsed)) {
    throw new Error("Invalid vercel.json root structure");
  }

  const headersRaw = parsed.headers;
  if (!Array.isArray(headersRaw)) {
    throw new Error("vercel.json is missing top-level headers array");
  }

  const taskpaneEntry = headersRaw.find((entry) => isRecord(entry) && entry.source === "/src/taskpane.html");
  if (!isRecord(taskpaneEntry)) {
    throw new Error("vercel.json is missing /src/taskpane.html header configuration");
  }

  const headerListRaw = taskpaneEntry.headers;
  if (!Array.isArray(headerListRaw)) {
    throw new Error("/src/taskpane.html entry has no headers array");
  }

  const cspEntry = headerListRaw.find((entry) => isRecord(entry) && entry.key === "Content-Security-Policy");
  if (!isRecord(cspEntry) || typeof cspEntry.value !== "string") {
    throw new Error("Missing Content-Security-Policy value for /src/taskpane.html");
  }

  const connectSrcDirective = cspEntry.value
    .split(";")
    .map((directive) => directive.trim())
    .find((directive) => directive.startsWith("connect-src "));

  if (!connectSrcDirective) {
    throw new Error("CSP missing connect-src directive");
  }

  const tokens = connectSrcDirective
    .split(/\s+/)
    .slice(1)
    .filter((token) => token.length > 0);

  return new Set(tokens);
}

test("proxy default host allowlist includes all web-search provider hosts", async () => {
  const proxyHosts = await readProxyAllowlistHosts();

  for (const host of uniqueHosts(WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS)) {
    assert.ok(proxyHosts.has(host), `Missing ${host} in proxy DEFAULT_ALLOWED_TARGET_HOSTS`);
  }
});

test("taskpane CSP connect-src allows all web-search provider hosts", async () => {
  const connectTokens = await readTaskpaneConnectSrcTokens();

  for (const host of uniqueHosts(WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS)) {
    const origin = `https://${host}`;
    assert.ok(connectTokens.has(origin), `Missing ${origin} in /src/taskpane.html CSP connect-src`);
  }
});
