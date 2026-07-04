/**
 * Tests for org/central-deployment build-time configuration:
 * - VITE_PI_DEFAULT_PROXY_URL → resolveDefaultProxyUrl()
 * - VITE_PI_ALLOWED_PROVIDERS → provider allowlist filtering
 */

import test from "node:test";
import assert from "node:assert/strict";

import { DEFAULT_LOCAL_PROXY_URL, resolveDefaultProxyUrl } from "../src/auth/proxy-validation.ts";
import { filterProvidersByAllowlist, resolveAllowedProviderIds } from "../src/ui/provider-allowlist.ts";

void test("resolveDefaultProxyUrl falls back to local default when unset", () => {
  assert.equal(resolveDefaultProxyUrl(undefined), DEFAULT_LOCAL_PROXY_URL);
  assert.equal(resolveDefaultProxyUrl(""), DEFAULT_LOCAL_PROXY_URL);
  assert.equal(resolveDefaultProxyUrl("   "), DEFAULT_LOCAL_PROXY_URL);
  assert.equal(resolveDefaultProxyUrl(42), DEFAULT_LOCAL_PROXY_URL);
});

void test("resolveDefaultProxyUrl accepts https URLs and strips trailing slashes", () => {
  assert.equal(
    resolveDefaultProxyUrl("https://pi-proxy.example.com:3003"),
    "https://pi-proxy.example.com:3003",
  );
  assert.equal(
    resolveDefaultProxyUrl("https://pi-proxy.example.com:3003/"),
    "https://pi-proxy.example.com:3003",
  );
});

void test("resolveDefaultProxyUrl refuses http (mixed content) and garbage", () => {
  assert.equal(resolveDefaultProxyUrl("http://pi-proxy.example.com:3003"), DEFAULT_LOCAL_PROXY_URL);
  assert.equal(resolveDefaultProxyUrl("pi-proxy.example.com"), DEFAULT_LOCAL_PROXY_URL);
  assert.equal(resolveDefaultProxyUrl("https://"), DEFAULT_LOCAL_PROXY_URL);
});

const PROVIDERS = [
  { id: "anthropic", label: "Anthropic" },
  { id: "openai", label: "OpenAI (API)" },
  { id: "deepseek", label: "DeepSeek" },
];

void test("resolveAllowedProviderIds returns null when unset", () => {
  assert.equal(resolveAllowedProviderIds(undefined), null);
  assert.equal(resolveAllowedProviderIds(""), null);
  assert.equal(resolveAllowedProviderIds(" , "), null);
  assert.equal(resolveAllowedProviderIds(7), null);
});

void test("resolveAllowedProviderIds parses and lowercases ids", () => {
  const ids = resolveAllowedProviderIds(" OpenAI, deepseek ,");
  assert.notEqual(ids, null);
  assert.deepEqual([...(ids as Set<string>)].sort(), ["deepseek", "openai"]);
});

void test("filterProvidersByAllowlist passes through with no restriction", () => {
  assert.deepEqual(filterProvidersByAllowlist(PROVIDERS, null), PROVIDERS);
});

void test("filterProvidersByAllowlist keeps only allowlisted providers in order", () => {
  const allowed = resolveAllowedProviderIds("deepseek,openai");
  const filtered = filterProvidersByAllowlist(PROVIDERS, allowed);
  assert.deepEqual(filtered.map((p) => p.id), ["openai", "deepseek"]);
});

void test("filterProvidersByAllowlist fails open on fully mismatched allowlist", () => {
  const allowed = resolveAllowedProviderIds("no-such-provider");
  const filtered = filterProvidersByAllowlist(PROVIDERS, allowed);
  assert.deepEqual(filtered.map((p) => p.id), ["anthropic", "openai", "deepseek"]);
});
