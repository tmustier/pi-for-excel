import assert from "node:assert/strict";
import { test } from "node:test";

import { isLikelyProxyConnectionError } from "../src/tools/external-fetch.ts";

void test("fetch_page proxy connection errors classify transport failures", () => {
  const proxyUrl = "https://localhost:3003";
  const rows = [
    "Load failed",
    "Failed to fetch",
    "fetch failed",
    "connect ECONNREFUSED 127.0.0.1:3003",
  ];

  for (const message of rows) {
    assert.equal(isLikelyProxyConnectionError(message, proxyUrl), true, message);
  }
});

void test("fetch_page proxy connection errors reject upstream and unrelated failures", () => {
  const proxyUrl = "https://localhost:3003";
  const rows = [
    "Invalid JSON in response body",
    "fetch_page request failed (502): Proxy error: fetch failed",
    "JSON-RPC error: upstream fetch failed while calling backend service",
  ];

  for (const message of rows) {
    assert.equal(isLikelyProxyConnectionError(message, proxyUrl), false, message);
  }
});
