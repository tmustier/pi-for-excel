import assert from "node:assert/strict";
import { test } from "node:test";

import type { ModelsStore, ModelsStoreEntry } from "@earendil-works/pi-ai";

import { BrowserModelRuntime } from "../src/models/browser-model-runtime.ts";
import type { CustomProvider } from "../src/storage/local/custom-providers-store.js";

import {
  DEFAULT_OPENAI_GATEWAY_CONTEXT_WINDOW,
  deleteOpenAiGatewayConfig,
  listOpenAiGatewayConfigs,
  resolveCustomProviderModel,
  saveOpenAiGatewayConfig,
  type CustomProvidersStoreLike,
} from "../src/auth/custom-gateways.ts";

class MemoryCatalogs implements ModelsStore {
  private readonly entries = new Map<string, ModelsStoreEntry>();

  read(providerId: string): Promise<ModelsStoreEntry | undefined> {
    return Promise.resolve(this.entries.get(providerId));
  }

  write(providerId: string, entry: ModelsStoreEntry): Promise<void> {
    this.entries.set(providerId, entry);
    return Promise.resolve();
  }

  delete(providerId: string): Promise<void> {
    this.entries.delete(providerId);
    return Promise.resolve();
  }
}

class MemoryProviderKeys {
  get(): Promise<null> { return Promise.resolve(null); }
  set(): Promise<void> { return Promise.resolve(); }
  delete(): Promise<void> { return Promise.resolve(); }
  list(): Promise<string[]> { return Promise.resolve([]); }
}

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

async function restartGatewayRuntime(store: MemoryCustomProvidersStore): Promise<BrowserModelRuntime> {
  const runtime = new BrowserModelRuntime({
    providerKeys: new MemoryProviderKeys(),
    modelCatalogs: new MemoryCatalogs(),
    getProxyUrl: () => Promise.resolve(undefined),
  });
  await runtime.syncCustomProviders(await store.getAll());
  return runtime;
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

  const restarted = await restartGatewayRuntime(store);
  const model = restarted.models.getModel(saved.providerName, "small-model");
  assert.equal(model?.contextWindow, 2_048);
  assert.equal(model?.maxTokens, 2_048);
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

  const firstRuntime = await restartGatewayRuntime(store);
  const persistedModel = firstRuntime.models.getModel(firstSave.providerName, "supply-chain");
  assert.ok(persistedModel);
  if (!persistedModel) throw new Error("Gateway model missing");

  await saveOpenAiGatewayConfig(store, {
    id: firstSave.id,
    displayName: "Warehouse API EU",
    endpointUrl: "https://warehouse.example.com/v1",
    modelId: "supply-chain",
    contextWindow: 262_144,
  });

  const restarted = await restartGatewayRuntime(store);
  const refreshed = restarted.models.getModel("Gateway · Warehouse API EU", "supply-chain");
  assert.equal(refreshed?.contextWindow, 262_144);
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

  const restarted = await restartGatewayRuntime(store);
  assert.equal(restarted.models.getProvider(saved.providerName), undefined);
  assert.ok(restarted.models.getProvider("manual-provider"));
});

void test("a restarted runtime resolves custom providers and their credentials", async () => {
  const store = new MemoryCustomProvidersStore();

  const gateway = await saveOpenAiGatewayConfig(store, {
    displayName: "Runtime Gateway",
    endpointUrl: "https://runtime.example.com/v1",
    modelId: "runtime-model",
    apiKey: "runtime-key",
  });

  await store.set({
    id: "custom-keyless-1",
    name: "Keyless provider",
    type: "openai-responses",
    baseUrl: "https://keyless.example.com/v1",
    models: [{
      id: "keyless-model",
      name: "keyless-model",
      api: "openai-responses",
      provider: "Keyless provider",
      baseUrl: "https://keyless.example.com/v1",
      reasoning: false,
      input: ["text"],
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
      contextWindow: 8_192,
      maxTokens: 1_024,
    }],
  });

  const restarted = await restartGatewayRuntime(store);

  assert.ok(restarted.models.getProvider(gateway.providerName));
  assert.ok(restarted.models.getProvider("Keyless provider"));
  assert.equal((await restarted.models.getAuth(gateway.providerName))?.auth.apiKey, "runtime-key");
  assert.ok(restarted.models.getModel(gateway.providerName, "runtime-model"));
});
