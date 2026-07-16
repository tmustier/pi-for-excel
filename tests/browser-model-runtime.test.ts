import assert from "node:assert/strict";
import { test } from "node:test";

import type {
  Credential,
  ModelsStore,
  ModelsStoreEntry,
} from "@earendil-works/pi-ai";

import {
  BrowserModelRuntime,
  type BrowserProviderRegistration,
} from "../src/models/browser-model-runtime.ts";
import type { CustomProvider } from "../src/storage/local/custom-providers-store.ts";
import {
  ProviderCredentialsStore,
  type ProviderKeysStoreLike,
} from "../src/storage/local/provider-credentials-store.ts";

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
    return Promise.resolve(Array.from(this.keys.keys()));
  }
}

class MemoryCatalogs implements ModelsStore {
  readonly entries = new Map<string, ModelsStoreEntry>();

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

function createRuntime(args?: {
  providerKeys?: MemoryProviderKeys;
  catalogs?: MemoryCatalogs;
  fetchFn?: typeof globalThis.fetch;
}): BrowserModelRuntime {
  return new BrowserModelRuntime({
    providerKeys: args?.providerKeys ?? new MemoryProviderKeys(),
    modelCatalogs: args?.catalogs ?? new MemoryCatalogs(),
    getProxyUrl: () => Promise.resolve(undefined),
    ...(args?.fetchFn !== undefined ? { fetchFn: args.fetchFn } : {}),
  });
}

function gatewayProvider(): CustomProvider {
  return {
    id: "stored-gateway",
    name: "Acme gateway",
    type: "openai-completions",
    baseUrl: "https://gateway.example.com/v1",
    apiKey: "gateway-secret",
    models: [{
      id: "configured-model",
      name: "Configured model",
      api: "openai-completions",
      provider: "Gateway · Acme",
      baseUrl: "https://gateway.example.com/v1",
      reasoning: false,
      input: ["text"],
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
      contextWindow: 65_536,
      maxTokens: 8_192,
    }],
  };
}

void test("browser runtime exposes built-in models through IndexedDB-backed credentials", async () => {
  const providerKeys = new MemoryProviderKeys();
  await providerKeys.set("openai", "test-key");
  const runtime = createRuntime({ providerKeys });

  assert.equal(runtime.models.getModel("openai", "gpt-5.6-sol")?.provider, "openai");
  const available = await runtime.models.getAvailable("openai");
  assert.ok(available.some((model) => model.id === "gpt-5.6-sol"));

  const auth = await runtime.models.getAuth("openai");
  assert.equal(auth?.auth.apiKey, "test-key");
});

void test("custom gateway discovery merges remote model ids and persists the catalogue", async () => {
  const catalogs = new MemoryCatalogs();
  let requestedUrl = "";
  let authorization = "";
  const fetchFn: typeof globalThis.fetch = (input, init) => {
    requestedUrl = typeof input === "string"
      ? input
      : input instanceof URL
        ? input.toString()
        : input.url;
    authorization = new Headers(init?.headers).get("authorization") ?? "";
    return Promise.resolve(new Response(JSON.stringify({
      data: [{ id: "remote-b" }, { id: "remote-a" }, { id: "remote-a" }],
    }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }));
  };

  const runtime = createRuntime({ catalogs, fetchFn });
  await runtime.syncCustomProviders([gatewayProvider()]);
  assert.equal(runtime.shouldProxyProvider("Gateway · Acme"), true);
  const result = await runtime.refresh({ allowNetwork: true });

  assert.equal(result.errors.size, 0);
  assert.equal(requestedUrl, "https://gateway.example.com/v1/models");
  assert.equal(authorization, "Bearer gateway-secret");
  assert.deepEqual(
    runtime.models.getModels("Gateway · Acme").map((model) => model.id),
    ["configured-model", "remote-a", "remote-b"],
  );
  assert.deepEqual(
    catalogs.entries.get("Gateway · Acme")?.models.map((model) => model.id),
    ["remote-a", "remote-b"],
  );
});

void test("a fresh runtime restores discovered models without network access", async () => {
  const catalogs = new MemoryCatalogs();
  let networkCalls = 0;
  const fetchFn: typeof globalThis.fetch = () => {
    networkCalls += 1;
    return Promise.resolve(new Response(JSON.stringify({ data: [{ id: "cached-model" }] }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }));
  };

  const first = createRuntime({ catalogs, fetchFn });
  await first.syncCustomProviders([gatewayProvider()]);
  await first.refresh({ allowNetwork: true });
  assert.equal(networkCalls, 1);

  const second = createRuntime({ catalogs, fetchFn });
  await second.syncCustomProviders([gatewayProvider()]);
  await second.refresh({ allowNetwork: false });

  assert.equal(networkCalls, 1, "cache-only startup must not call the gateway");
  assert.ok(second.models.getModel("Gateway · Acme", "cached-model"));
});

void test("cached discovery is rebound to the current provider transport", async () => {
  const catalogs = new MemoryCatalogs();
  const first = createRuntime({
    catalogs,
    fetchFn: () => Promise.resolve(new Response(JSON.stringify({
      data: [{ id: "cached-model" }],
    }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    })),
  });
  await first.syncCustomProviders([gatewayProvider()]);
  await first.refresh({ allowNetwork: true });

  const providerId = "Gateway · Acme";
  const cached = catalogs.entries.get(providerId);
  assert.ok(cached);
  catalogs.entries.set(providerId, {
    ...cached,
    models: cached.models.map((model) => ({
      ...model,
      headers: { "x-stale-endpoint-header": "must-not-survive" },
    })),
  });

  const updated = gatewayProvider();
  updated.type = "openai-responses";
  updated.baseUrl = "https://new-gateway.example.com/v2";
  updated.apiKey = "new-gateway-secret";
  updated.models = updated.models?.map((model) => ({
    ...model,
    api: "openai-responses",
    baseUrl: updated.baseUrl,
  }));

  const second = createRuntime({
    catalogs,
    fetchFn: () => {
      throw new Error("cache-only restore must not use the network");
    },
  });
  await second.syncCustomProviders([updated]);
  await second.refresh({ allowNetwork: false });

  const restored = second.models.getModel(providerId, "cached-model");
  assert.equal(restored?.api, "openai-responses");
  assert.equal(restored?.baseUrl, "https://new-gateway.example.com/v2");
  assert.equal(restored?.headers, undefined);
  assert.equal((await second.models.getAuth(providerId))?.auth.apiKey, "new-gateway-secret");
});

void test("failed remote discovery retains the last cached overlay and baseline", async () => {
  const catalogs = new MemoryCatalogs();
  let shouldFail = false;
  const fetchFn: typeof globalThis.fetch = () => {
    if (shouldFail) {
      return Promise.resolve(new Response("unavailable", { status: 503 }));
    }
    return Promise.resolve(new Response(JSON.stringify({ data: [{ id: "last-known-model" }] }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }));
  };

  const runtime = createRuntime({ catalogs, fetchFn });
  await runtime.syncCustomProviders([gatewayProvider()]);
  await runtime.refresh({ allowNetwork: true });
  shouldFail = true;

  const failed = await runtime.refresh({ allowNetwork: true });
  assert.equal(failed.errors.has("Gateway · Acme"), true);
  assert.deepEqual(
    runtime.models.getModels("Gateway · Acme").map((model) => model.id),
    ["configured-model", "last-known-model"],
  );
});

void test("extension providers are owner-scoped and can resolve host-owned credentials", async () => {
  const runtime = createRuntime();
  const registration: BrowserProviderRegistration = {
    id: "ext.example.provider",
    name: "Example provider",
    api: "openai-responses",
    baseUrl: "https://models.example.com/v1",
    models: [{ id: "example-model", contextWindow: 128_000, maxTokens: 16_000 }],
    modelsUrl: "https://models.example.com/v1/models",
    resolveApiKey: () => Promise.resolve("host-owned-secret"),
  };

  runtime.registerExtensionProvider("ext.example", registration);
  assert.equal(runtime.isExtensionProvider(registration.id), true);
  assert.equal(runtime.shouldProxyProvider(registration.id), true);
  assert.equal(runtime.shouldProxyProvider("openai"), false);
  assert.equal((await runtime.models.getAuth(registration.id))?.auth.apiKey, "host-owned-secret");
  assert.ok((await runtime.models.getAvailable(registration.id)).some((model) => model.id === "example-model"));

  await assert.rejects(
    runtime.unregisterExtensionProvider("ext.someone-else", registration.id),
    /not owned/,
  );
  await runtime.unregisterExtensionProvider("ext.example", registration.id);
  assert.equal(runtime.models.getProvider(registration.id), undefined);
  assert.equal(runtime.shouldProxyProvider(registration.id), false);
});

void test("credential adapter does not reinterpret OAuth records as browser API keys", async () => {
  const credentials = new ProviderCredentialsStore(new MemoryProviderKeys());

  await assert.rejects(
    credentials.modify("openai", () => Promise.resolve({
      type: "oauth",
      access: "access",
      refresh: "refresh",
      expires: Date.now() + 60_000,
    } satisfies Credential)),
    /taskpane OAuth store/,
  );
});
