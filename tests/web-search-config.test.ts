import assert from "node:assert/strict";
import { test } from "node:test";

import {
  clearWebSearchApiKey,
  getApiKeyForProvider,
  isApiKeyRequired,
  loadWebSearchProviderConfig,
  saveWebSearchApiKey,
  saveWebSearchProvider,
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

void test("jina does not require an API key", () => {
  assert.equal(isApiKeyRequired("jina"), false);
});

void test("serper/tavily/brave require an API key", () => {
  assert.equal(isApiKeyRequired("serper"), true);
  assert.equal(isApiKeyRequired("tavily"), true);
  assert.equal(isApiKeyRequired("brave"), true);
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
