import assert from "node:assert/strict";
import { test } from "node:test";

import type { ModelsStore, ModelsStoreEntry } from "@earendil-works/pi-ai";
import { getBuiltinModel } from "@earendil-works/pi-ai/providers/all";

import { BrowserModelRuntime } from "../src/models/browser-model-runtime.ts";
import { ModelRefreshOwner, type ModelRefreshRuntime } from "../src/models/model-refresh-owner.ts";
import type { CustomProvider } from "../src/storage/local/custom-providers-store.ts";
import type { ProviderKeysStoreLike } from "../src/storage/local/provider-credentials-store.ts";

class MemoryProviderKeys implements ProviderKeysStoreLike {
  private readonly keys = new Map<string, string>();

  get(provider: string): Promise<string | null> {
    return Promise.resolve(this.keys.get(provider) ?? null);
  }

  set(provider: string, key: string): Promise<void> {
    this.keys.set(provider, key);
    return Promise.resolve();
  }

  delete(provider: string): Promise<void> {
    this.keys.delete(provider);
    return Promise.resolve();
  }

  list(): Promise<string[]> {
    return Promise.resolve([...this.keys.keys()]);
  }
}

class MemoryCatalogs implements ModelsStore {
  read(_providerId: string): Promise<ModelsStoreEntry | undefined> {
    return Promise.resolve(undefined);
  }

  write(_providerId: string, _entry: ModelsStoreEntry): Promise<void> {
    return Promise.resolve();
  }

  delete(_providerId: string): Promise<void> {
    return Promise.resolve();
  }
}

function customProvider(): CustomProvider {
  return {
    id: "local-provider",
    name: "Local provider",
    type: "openai-completions",
    baseUrl: "https://models.example.test/v1",
    apiKey: "test-key",
    models: [{
      id: "plain-model",
      name: "Plain model",
      api: "openai-completions",
      provider: "Gateway · Local provider",
      baseUrl: "https://models.example.test/v1",
      reasoning: false,
      input: ["text"],
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
      contextWindow: 32_768,
      maxTokens: 4_096,
    }],
  };
}

function createOwner(
  keys: MemoryProviderKeys,
  runtimes: ModelRefreshRuntime[],
  providers: readonly CustomProvider[] = [],
): ModelRefreshOwner {
  const modelRuntime = new BrowserModelRuntime({
    providerKeys: keys,
    modelCatalogs: new MemoryCatalogs(),
    getProxyUrl: () => Promise.resolve(undefined),
    fetchFn: () => {
      throw new Error("cache-only reconciliation must not use the network");
    },
  });
  return new ModelRefreshOwner({
    modelRuntime,
    loadCustomProviders: () => Promise.resolve(providers),
    getRuntimes: () => runtimes,
  });
}

void test("provider refresh leaves a session on its configured provider", async () => {
  const keys = new MemoryProviderKeys();
  await keys.set("openai", "test-key");
  const original = getBuiltinModel("openai", "gpt-5.6-sol");
  let selected = original;
  const runtimes: ModelRefreshRuntime[] = [{
    runtimeId: "session-a",
    model: selected,
    isBusy: false,
    applyModel: (model) => {
      selected = model;
      runtimes[0].model = model;
    },
  }];

  await createOwner(keys, runtimes).restoreCached();

  assert.equal(selected.provider, "openai");
  assert.equal(selected.id, "gpt-5.6-sol");
});

void test("provider refresh switches an unavailable session to an available non-reasoning model", async () => {
  const keys = new MemoryProviderKeys();
  let selected = getBuiltinModel("anthropic", "claude-opus-4-8");
  let thinkingLevel = "high";
  const runtimes: ModelRefreshRuntime[] = [{
    runtimeId: "session-a",
    model: selected,
    isBusy: false,
    applyModel: (model, nextThinkingLevel) => {
      selected = model;
      thinkingLevel = nextThinkingLevel ?? thinkingLevel;
      runtimes[0].model = model;
    },
  }];

  await createOwner(keys, runtimes, [customProvider()]).restoreCached();

  assert.equal(selected.provider, "Gateway · Local provider");
  assert.equal(selected.id, "plain-model");
  assert.equal(thinkingLevel, "off");
});

void test("provider refresh applies metadata-only changes to a session on the same model", async () => {
  const keys = new MemoryProviderKeys();
  const provider = customProvider();
  const catalogModel = provider.models[0];
  assert.ok(catalogModel);
  // Same identity and token limits as the catalogue entry; everything else is
  // what an older gateway save carried before the registry match succeeded.
  const staleModel = {
    ...catalogModel,
    reasoning: false,
    input: ["text" as const],
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
  };
  provider.models[0] = {
    ...catalogModel,
    reasoning: true,
    thinkingLevelMap: { low: "minimal", high: "xhigh" },
    input: ["text", "image"],
    cost: { input: 0.6, output: 2.5, cacheRead: 0.1, cacheWrite: 0 },
  };
  let selected = staleModel;
  const runtimes: ModelRefreshRuntime[] = [{
    runtimeId: "session-a",
    model: staleModel,
    isBusy: false,
    applyModel: (model) => {
      selected = model;
      runtimes[0].model = model;
    },
  }];

  await createOwner(keys, runtimes, [provider]).restoreCached();

  assert.deepEqual(selected, provider.models[0]);
});
