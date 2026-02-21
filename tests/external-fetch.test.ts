import assert from "node:assert/strict";
import { test } from "node:test";

import {
  buildProxyDownErrorMessage,
  getEnabledProxyBaseUrl,
  isLikelyProxyConnectionError,
  resolveOutboundRequestUrl,
  type ProxyAwareSettingsStore,
} from "../src/tools/external-fetch.ts";

class MemorySettingsStore implements ProxyAwareSettingsStore {
  private readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    const value = this.values.has(key) ? this.values.get(key) ?? null : null;
    return Promise.resolve(value);
  }

  set(key: string, value: unknown): void {
    this.values.set(key, value);
  }
}

void test("getEnabledProxyBaseUrl falls back to localhost:3003 when enabled and URL missing", async () => {
  const settings = new MemorySettingsStore();
  settings.set("proxy.enabled", true);

  const proxyUrl = await getEnabledProxyBaseUrl(settings);
  assert.equal(proxyUrl, "https://localhost:3003");
});

void test("getEnabledProxyBaseUrl ignores proxy URL when disabled", async () => {
  const settings = new MemorySettingsStore();
  settings.set("proxy.enabled", false);
  settings.set("proxy.url", "https://localhost:3004");

  const proxyUrl = await getEnabledProxyBaseUrl(settings);
  assert.equal(proxyUrl, undefined);
});

void test("resolveOutboundRequestUrl wraps target URL when proxy is enabled", () => {
  const resolved = resolveOutboundRequestUrl({
    targetUrl: "https://example.com/resource?q=1",
    proxyBaseUrl: "https://localhost:3003",
  });

  assert.equal(resolved.proxied, true);
  assert.equal(
    resolved.requestUrl,
    "https://localhost:3003/?url=https%3A%2F%2Fexample.com%2Fresource%3Fq%3D1",
  );
});

/* ── Proxy-down error detection ─────────────────────────────── */

void test("isLikelyProxyConnectionError returns true for WebKit 'Load failed' when proxy is set", () => {
  assert.equal(isLikelyProxyConnectionError("Load failed", "https://localhost:3003"), true);
});

void test("isLikelyProxyConnectionError returns true for Chrome 'Failed to fetch' when proxy is set", () => {
  assert.equal(isLikelyProxyConnectionError("Failed to fetch", "https://localhost:3003"), true);
});

void test("isLikelyProxyConnectionError returns true for Node 'fetch failed' when proxy is set", () => {
  assert.equal(isLikelyProxyConnectionError("fetch failed", "https://localhost:3003"), true);
});

void test("isLikelyProxyConnectionError returns true for ECONNREFUSED when proxy is set", () => {
  assert.equal(
    isLikelyProxyConnectionError("connect ECONNREFUSED 127.0.0.1:3003", "https://localhost:3003"),
    true,
  );
});

void test("isLikelyProxyConnectionError returns false when no proxy is configured", () => {
  assert.equal(isLikelyProxyConnectionError("Load failed", undefined), false);
});

void test("isLikelyProxyConnectionError returns false for non-network errors with proxy", () => {
  assert.equal(
    isLikelyProxyConnectionError("Invalid JSON in response body", "https://localhost:3003"),
    false,
  );
});

void test("isLikelyProxyConnectionError returns false when proxy answered with upstream fetch failure", () => {
  assert.equal(
    isLikelyProxyConnectionError(
      "fetch_page request failed (502): Proxy error: fetch failed",
      "https://localhost:3003",
    ),
    false,
  );
});

void test("buildProxyDownErrorMessage includes tool label, fix command, and original error", () => {
  const message = buildProxyDownErrorMessage("Web search", "Load failed");
  assert.ok(message.startsWith("Error: Web search failed"));
  assert.ok(message.includes("npx pi-for-excel-proxy"));
  assert.ok(message.includes("Do not retry"));
  assert.ok(message.includes("Load failed"));
});
