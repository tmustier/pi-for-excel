import assert from "node:assert/strict";
import { test } from "node:test";

import type { CustomProvider } from "@mariozechner/pi-web-ui/dist/storage/stores/custom-providers-store.js";

import {
  collectCustomProviderRuntimeInfo,
  deleteOpenAiGatewayConfig,
  listOpenAiGatewayConfigs,
  saveOpenAiGatewayConfig,
  type CustomProvidersStoreLike,
} from "../src/auth/custom-gateways.ts";

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

  const listed = await listOpenAiGatewayConfigs(store);
  assert.equal(listed.length, 1);
  assert.equal(listed[0]?.providerName, saved.providerName);
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
