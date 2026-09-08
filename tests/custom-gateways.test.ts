import assert from "node:assert/strict";
import { test } from "node:test";

import type { Api, Model } from "@earendil-works/pi-ai";
import type { CustomProvider } from "../src/storage/local/custom-providers-store.js";

import {
  DEFAULT_OPENAI_GATEWAY_CONTEXT_WINDOW,
  collectCustomProviderRuntimeInfo,
  deleteOpenAiGatewayConfig,
  listOpenAiGatewayConfigs,
  resolveCustomProviderModel,
  saveOpenAiGatewayConfig,
  type CustomProvidersStoreLike,
} from "../src/auth/custom-gateways.ts";

const knownRegistryModel: Model<Api> = {
  id: "openai/gpt-5.6-sol",
  name: "OpenAI: GPT-5.6 Sol",
  api: "openai-completions",
  provider: "openrouter",
  baseUrl: "https://openrouter.ai/api/v1",
  reasoning: true,
  thinkingLevelMap: {
    off: "none",
    low: "low",
    high: "high",
  },
  input: ["text", "image"],
  cost: {
    input: 2,
    output: 10,
    cacheRead: 0.2,
    cacheWrite: 2.5,
  },
  contextWindow: 1_050_000,
  maxTokens: 128_000,
  headers: { "X-Registry-Only": "not-copied" },
  compat: { thinkingFormat: "openrouter" },
};

class MemoryCustomProvidersStore implements CustomProvidersStoreLike {
  private readonly providers = new Map<string, CustomProvider>();

  get(id: string): Promise<CustomProvider | null> {
    return Promise.resolve(this.providers.get(id) ?? null);
  }

  set(provider: CustomProvider): Promise<void> {
    this.providers.set(provider.id, provider);
    return Promise.resolve();
  }

  delete(id: string): Promise<void> {
    this.providers.delete(id);
    return Promise.resolve();
  }

  getAll(): Promise<CustomProvider[]> {
    return Promise.resolve(Array.from(this.providers.values()));
  }
}

void test("saveOpenAiGatewayConfig stores normalized endpoint/model/provider", async () => {
  const store = new MemoryCustomProvidersStore();

  const saved = await saveOpenAiGatewayConfig(store, {
    endpointUrl: "https://gateway.example.com/v1/",
    modelId: "gpt-4o-mini",
    apiKey: " sk-test ",
  });

  assert.equal(saved.displayName, "gateway.example.com");
  assert.equal(saved.endpointUrl, "https://gateway.example.com/v1");
  assert.equal(saved.modelId, "gpt-4o-mini");
  assert.equal(saved.apiKey, "sk-test");
  assert.match(saved.providerName, /^Gateway · gateway\.example\.com/);
  assert.equal(saved.contextWindow, DEFAULT_OPENAI_GATEWAY_CONTEXT_WINDOW);

  const listed = await listOpenAiGatewayConfigs(store);
  assert.equal(listed.length, 1);
  assert.equal(listed[0]?.providerName, saved.providerName);
  assert.equal(listed[0]?.contextWindow, DEFAULT_OPENAI_GATEWAY_CONTEXT_WINDOW);
});

void test("saveOpenAiGatewayConfig uses matched registry metadata by default", async () => {
  const store = new MemoryCustomProvidersStore();

  const saved = await saveOpenAiGatewayConfig(store, {
    endpointUrl: "https://openrouter.ai/api/v1/",
    modelId: knownRegistryModel.id,
  }, {
    getModels: () => [knownRegistryModel],
  });

  assert.equal(saved.contextWindow, 1_050_000);

  const storedModel = (await store.get(saved.id))?.models?.[0];
  assert.equal(storedModel?.contextWindow, 1_050_000);
  assert.equal(storedModel?.maxTokens, 128_000);
  assert.equal(storedModel?.reasoning, true);
  assert.deepEqual(storedModel?.thinkingLevelMap, knownRegistryModel.thinkingLevelMap);
  assert.deepEqual(storedModel?.input, ["text", "image"]);
  assert.deepEqual(storedModel?.cost, knownRegistryModel.cost);
  assert.equal(storedModel?.headers, undefined);
  assert.equal(storedModel?.compat, undefined);
});

void test("saveOpenAiGatewayConfig lets explicit context override matched registry context", async () => {
  const store = new MemoryCustomProvidersStore();

  const saved = await saveOpenAiGatewayConfig(store, {
    endpointUrl: knownRegistryModel.baseUrl,
    modelId: knownRegistryModel.id,
    contextWindow: 64_000,
  }, {
    getModels: () => [knownRegistryModel],
  });

  const storedModel = (await store.get(saved.id))?.models?.[0];
  assert.equal(saved.contextWindow, 64_000);
  assert.equal(storedModel?.contextWindow, 64_000);
  assert.equal(storedModel?.maxTokens, 64_000);
  assert.equal(storedModel?.reasoning, true);
});

void test("saveOpenAiGatewayConfig falls back when registry endpoint, id, or API differs", async () => {
  const mismatchedModels: Model<Api>[] = [
    { ...knownRegistryModel, baseUrl: "https://different.example.com/v1" },
    { ...knownRegistryModel, id: "openai/different-model" },
    { ...knownRegistryModel, api: "openai-responses" },
  ];

  for (const registryModel of mismatchedModels) {
    const store = new MemoryCustomProvidersStore();
    const saved = await saveOpenAiGatewayConfig(store, {
      endpointUrl: knownRegistryModel.baseUrl,
      modelId: knownRegistryModel.id,
    }, {
      getModels: () => [registryModel],
    });

    const storedModel = (await store.get(saved.id))?.models?.[0];
    assert.equal(saved.contextWindow, DEFAULT_OPENAI_GATEWAY_CONTEXT_WINDOW);
    assert.equal(storedModel?.reasoning, false);
    assert.deepEqual(storedModel?.input, ["text"]);
    assert.deepEqual(storedModel?.cost, {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
    });
    assert.equal(storedModel?.maxTokens, 4_096);
  }
});

void test("saveOpenAiGatewayConfig stores custom context window metadata", async () => {
  const store = new MemoryCustomProvidersStore();

  const saved = await saveOpenAiGatewayConfig(store, {
    displayName: "Big context",
    endpointUrl: "https://gateway.example.com/v1",
    modelId: "big-model",
    contextWindow: 131_072,
  });

  assert.equal(saved.contextWindow, 131_072);

  const listed = await listOpenAiGatewayConfigs(store);
  assert.equal(listed[0]?.contextWindow, 131_072);
});

void test("saveOpenAiGatewayConfig clamps maxTokens to the configured context window", async () => {
  const store = new MemoryCustomProvidersStore();

  const saved = await saveOpenAiGatewayConfig(store, {
    displayName: "Tight budget",
    endpointUrl: "https://gateway.example.com/v1",
    modelId: "small-model",
    contextWindow: 2_048,
  });

  const storedModel = (await store.get(saved.id))?.models?.[0];
  assert.equal(storedModel?.contextWindow, 2_048);
  assert.equal(storedModel?.maxTokens, 2_048);
});

void test("saveOpenAiGatewayConfig rejects invalid context window values", async () => {
  const store = new MemoryCustomProvidersStore();

  await assert.rejects(
    saveOpenAiGatewayConfig(store, {
      endpointUrl: "https://gateway.example.com/v1",
      modelId: "too-small",
      contextWindow: 512,
    }),
    /at least 1024/i,
  );
});

void test("gateway provider names stay unique when display names collide", async () => {
  const store = new MemoryCustomProvidersStore();

  const first = await saveOpenAiGatewayConfig(store, {
    displayName: "ACME",
    endpointUrl: "https://acme.example.com/v1",
    modelId: "model-a",
  });

  const second = await saveOpenAiGatewayConfig(store, {
    displayName: "ACME",
    endpointUrl: "https://acme-2.example.com/v1",
    modelId: "model-b",
  });

  assert.notEqual(first.providerName, second.providerName);
  assert.match(second.providerName, /\(2\)$/);
});

void test("resolveCustomProviderModel refreshes renamed gateway models by base URL and model id", async () => {
  const store = new MemoryCustomProvidersStore();

  const firstSave = await saveOpenAiGatewayConfig(store, {
    displayName: "Warehouse API",
    endpointUrl: "https://warehouse.example.com/v1",
    modelId: "supply-chain",
    contextWindow: 16_384,
  });

  const persistedModel = (await store.get(firstSave.id))?.models?.[0];
  assert.ok(persistedModel);
  if (!persistedModel) {
    throw new Error("Persisted model missing");
  }

  await saveOpenAiGatewayConfig(store, {
    id: firstSave.id,
    displayName: "Warehouse API EU",
    endpointUrl: "https://warehouse.example.com/v1",
    modelId: "supply-chain",
    contextWindow: 262_144,
  });

  const refreshed = resolveCustomProviderModel(await store.getAll(), persistedModel);
  assert.ok(refreshed);
  assert.equal(refreshed?.contextWindow, 262_144);
  assert.match(refreshed?.provider ?? "", /^Gateway · Warehouse API EU/);
});

void test("resolveCustomProviderModel refuses ambiguous base-url fallback matches", async () => {
  const store = new MemoryCustomProvidersStore();

  await saveOpenAiGatewayConfig(store, {
    displayName: "Warehouse API US",
    endpointUrl: "https://warehouse.example.com/v1",
    modelId: "supply-chain",
  });
  await saveOpenAiGatewayConfig(store, {
    displayName: "Warehouse API EU",
    endpointUrl: "https://warehouse.example.com/v1",
    modelId: "supply-chain",
  });

  const resolved = resolveCustomProviderModel(await store.getAll(), {
    api: "openai-completions",
    id: "supply-chain",
    provider: "Gateway · Warehouse API",
    baseUrl: "https://warehouse.example.com/v1",
  });

  assert.equal(resolved, null);
});

void test("deleteOpenAiGatewayConfig only removes managed gateway entries", async () => {
  const store = new MemoryCustomProvidersStore();

  const saved = await saveOpenAiGatewayConfig(store, {
    displayName: "Delete me",
    endpointUrl: "https://delete.example.com/v1",
    modelId: "model-delete",
  });

  await store.set({
    id: "manual-openai-provider",
    name: "Manual provider",
    type: "openai-completions",
    baseUrl: "https://manual.example.com/v1",
    models: [{
      id: "manual-model",
      name: "manual-model",
      api: "openai-completions",
      provider: "manual-provider",
      baseUrl: "https://manual.example.com/v1",
      reasoning: false,
      input: ["text"],
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
      contextWindow: 8192,
      maxTokens: 1024,
    }],
  });

  await deleteOpenAiGatewayConfig(store, saved.id);
  await deleteOpenAiGatewayConfig(store, "manual-openai-provider");

  const all = await store.getAll();
  assert.equal(all.length, 1);
  assert.equal(all[0]?.id, "manual-openai-provider");
});

void test("collectCustomProviderRuntimeInfo includes custom provider names and api keys", async () => {
  const store = new MemoryCustomProvidersStore();

  const gateway = await saveOpenAiGatewayConfig(store, {
    displayName: "Runtime Gateway",
    endpointUrl: "https://runtime.example.com/v1",
    modelId: "runtime-model",
    apiKey: "runtime-key",
  });

  await store.set({
    id: "custom-ollama-1",
    name: "Local Ollama",
    type: "ollama",
    baseUrl: "http://localhost:11434",
  });

  const runtimeInfo = collectCustomProviderRuntimeInfo(await store.getAll());

  assert.ok(runtimeInfo.providerNames.has(gateway.providerName));
  assert.ok(runtimeInfo.providerNames.has("Local Ollama"));
  assert.equal(runtimeInfo.apiKeys.get(gateway.providerName), "runtime-key");
  assert.equal(runtimeInfo.apiKeys.get("Local Ollama"), undefined);
  assert.equal(runtimeInfo.defaultModel?.id, "runtime-model");
});
