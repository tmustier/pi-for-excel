import assert from "node:assert/strict";
import { test } from "node:test";

import {
  checkApiKeyFormat,
  clearWebSearchApiKey,
  getApiKeyForProvider,
  getWebSearchEndpoint,
  isApiKeyRequired,
  loadWebSearchProviderConfig,
  saveWebSearchApiKey,
  saveWebSearchProvider,
  WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS,
  WEB_SEARCH_PROVIDERS,
  type WebSearchConfigStore,
} from "../src/tools/web-search-config.ts";

class MemorySettingsStore implements WebSearchConfigStore {
  private readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.has(key) ? this.values.get(key) ?? null : null);
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }

  delete(key: string): Promise<void> {
    this.values.delete(key);
    return Promise.resolve();
  }
}

void test("web search config defaults to jina provider", async () => {
  const settings = new MemorySettingsStore();
  const config = await loadWebSearchProviderConfig(settings);

  assert.equal(config.provider, "jina");
  assert.equal(getApiKeyForProvider(config), undefined);
});

void test("web search config infers serper when only serper key is present", async () => {
  const settings = new MemorySettingsStore();
  await saveWebSearchApiKey(settings, "serper", "sp-legacy");

  const config = await loadWebSearchProviderConfig(settings);

  assert.equal(config.provider, "serper");
  assert.equal(getApiKeyForProvider(config), "sp-legacy");
});

void test("web search config infers brave when only brave key is present", async () => {
  const settings = new MemorySettingsStore();
  await saveWebSearchApiKey(settings, "brave", "br-legacy");

  const config = await loadWebSearchProviderConfig(settings);

  assert.equal(config.provider, "brave");
  assert.equal(getApiKeyForProvider(config), "br-legacy");
});

void test("web search config stores provider-specific api keys", async () => {
  const settings = new MemorySettingsStore();

  await saveWebSearchProvider(settings, "tavily");
  await saveWebSearchApiKey(settings, "tavily", "tv-123");
  await saveWebSearchApiKey(settings, "brave", "br-123");

  const config = await loadWebSearchProviderConfig(settings);

  assert.equal(config.provider, "tavily");
  assert.equal(getApiKeyForProvider(config), "tv-123");
  assert.equal(getApiKeyForProvider(config, "brave"), "br-123");
});

void test("all providers require an API key", () => {
  assert.equal(isApiKeyRequired("jina"), true);
  assert.equal(isApiKeyRequired("firecrawl"), true);
  assert.equal(isApiKeyRequired("serper"), true);
  assert.equal(isApiKeyRequired("tavily"), true);
  assert.equal(isApiKeyRequired("brave"), true);
});

void test("web search config infers firecrawl when only firecrawl key is present", async () => {
  const settings = new MemorySettingsStore();
  await saveWebSearchApiKey(settings, "firecrawl", "fc-legacy");

  const config = await loadWebSearchProviderConfig(settings);

  assert.equal(config.provider, "firecrawl");
  assert.equal(getApiKeyForProvider(config), "fc-legacy");
});

void test("web search config stores firecrawl api key", async () => {
  const settings = new MemorySettingsStore();

  await saveWebSearchProvider(settings, "firecrawl");
  await saveWebSearchApiKey(settings, "firecrawl", "fc_test_key");

  const config = await loadWebSearchProviderConfig(settings);

  assert.equal(config.provider, "firecrawl");
  assert.equal(getApiKeyForProvider(config, "firecrawl"), "fc_test_key");
});

void test("web search config stores jina api key", async () => {
  const settings = new MemorySettingsStore();

  await saveWebSearchProvider(settings, "jina");
  await saveWebSearchApiKey(settings, "jina", "jina_test_key");

  const config = await loadWebSearchProviderConfig(settings);

  assert.equal(config.provider, "jina");
  assert.equal(getApiKeyForProvider(config, "jina"), "jina_test_key");
});

void test("web search config clears only the selected provider key", async () => {
  const settings = new MemorySettingsStore();

  await saveWebSearchApiKey(settings, "serper", "sp-123");
  await saveWebSearchApiKey(settings, "brave", "br-123");
  await clearWebSearchApiKey(settings, "serper");

  const config = await loadWebSearchProviderConfig(settings);

  assert.equal(getApiKeyForProvider(config, "serper"), undefined);
  assert.equal(getApiKeyForProvider(config, "brave"), "br-123");
});

// ── checkApiKeyFormat ────────────────────────────────────

void test("checkApiKeyFormat returns null for valid keys", () => {
  assert.equal(checkApiKeyFormat("jina", "jina_abc123defXYZ456"), null);
  assert.equal(checkApiKeyFormat("firecrawl", "fc-abc123def456"), null);
  assert.equal(checkApiKeyFormat("tavily", "tvly-abc123def456"), null);
  assert.equal(checkApiKeyFormat("serper", "abc123def456ghi789xyz0"), null);
  assert.equal(checkApiKeyFormat("brave", "BSAabc123def456"), null);
});

void test("checkApiKeyFormat catches empty and whitespace keys", () => {
  assert.ok(checkApiKeyFormat("jina", ""));
  assert.ok(checkApiKeyFormat("jina", "   "));
  assert.match(checkApiKeyFormat("jina", "jina_abc def") ?? "", /spaces/i);
  assert.match(checkApiKeyFormat("jina", "jina_abc\ndef") ?? "", /spaces/i);
});

void test("checkApiKeyFormat catches too-short keys", () => {
  assert.match(checkApiKeyFormat("serper", "abc") ?? "", /short/i);
});

void test("checkApiKeyFormat catches repeated long segments", () => {
  const key = "jina_abcdef1234567890";
  assert.match(checkApiKeyFormat("jina", `${key}${key}`) ?? "", /repeated long segment/i);
  assert.match(checkApiKeyFormat("jina", `${key}${key}Z`) ?? "", /repeated long segment/i);
});

void test("checkApiKeyFormat warns on wrong prefix", () => {
  assert.match(checkApiKeyFormat("jina", "sk-abc123def456abc123") ?? "", /jina_/);
  assert.match(checkApiKeyFormat("firecrawl", "jina_abc123def456abc1") ?? "", /fc-/);
  assert.match(checkApiKeyFormat("tavily", "sk-abc123def456abc123") ?? "", /tvly-/);
});

void test("checkApiKeyFormat does not warn on unknown prefix for serper/brave", () => {
  assert.equal(checkApiKeyFormat("serper", "anyformat1234567890"), null);
  assert.equal(checkApiKeyFormat("brave", "anyformat1234567890"), null);
});

void test("web search provider endpoints expose stable host list", () => {
  const expectedHosts = WEB_SEARCH_PROVIDERS.map((provider) => new URL(getWebSearchEndpoint(provider)).hostname.toLowerCase());

  assert.deepEqual(WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS, expectedHosts);
  assert.equal(new Set(WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS).size, WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS.length);
});
