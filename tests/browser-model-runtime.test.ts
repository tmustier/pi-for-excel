import "fake-indexeddb/auto";
import assert from "node:assert/strict";
import { test, type TestContext } from "node:test";

import type {
  Credential,
  ModelsStore,
  ModelsStoreEntry,
  SimpleStreamOptions,
} from "@earendil-works/pi-ai";
import { getBuiltinModel } from "@earendil-works/pi-ai/providers/all";

import {
  BrowserModelRuntime,
  type BrowserProviderRegistration,
} from "../src/models/browser-model-runtime.ts";
import {
  ModelRefreshOwner,
  type ModelRefreshRuntime,
} from "../src/models/model-refresh-owner.ts";
import type { CustomProvider } from "../src/storage/local/custom-providers-store.ts";
import { ModelCatalogsStore } from "../src/storage/local/model-catalogs-store.ts";
import { IndexedDBStorageBackend } from "../src/storage/local/indexeddb-storage-backend.ts";
import { getAppStorageConfig } from "../src/storage/init-app-storage.ts";
import { saveOpenAiGatewayConfig } from "../src/auth/custom-gateways.ts";
import { createOfficeStreamFn } from "../src/auth/stream-proxy.ts";
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

void test("Office streaming preserves OpenRouter's registry transport without unsupported effort updates", async (t) => {
  const originalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
  Object.defineProperty(globalThis, "document", { configurable: true, value: new EventTarget() });
  t.after(() => {
    if (originalDocument) Object.defineProperty(globalThis, "document", originalDocument);
    else Reflect.deleteProperty(globalThis, "document");
  });
  const runtime = createRuntime();
  for (const id of ["anthropic/claude-opus-5", "openai/gpt-5.6-sol"]) {
    const model = runtime.models.getModel("openrouter", id);
    assert.ok(model);
    let requestUrl = "";
    let requestBody = "";
    let betaHeaders = "";
    const stream = createOfficeStreamFn(() => Promise.resolve("https://localhost:3004"), runtime.models);
    const streamOptions: SimpleStreamOptions = {
      apiKey: "synthetic-key",
      maxTokens: 256,
      reasoning: "high",
      fetch: (url, init) => {
        requestUrl = typeof url === "string" ? url : url instanceof URL ? url.href : url.url;
        assert.ok(typeof init?.body === "string");
        requestBody = init.body;
        betaHeaders = new Headers(init?.headers).get("anthropic-beta") ?? "";
        return Promise.resolve(new Response("test endpoint", { status: 400 }));
      },
    };
    const events = await stream(model, { messages: [{ role: "user", content: "test", timestamp: 0 }] }, streamOptions);
    await events.result();
    if (model.api === "anthropic-messages") {
      assert.ok(requestUrl.startsWith("https://localhost:3004/?url="));
      assert.doesNotMatch(requestBody, /"role":"system"/u);
      assert.match(requestBody, /"thinking":\{"type":"adaptive"/u);
      assert.doesNotMatch(betaHeaders, /mid-conversation-output-config/u);
      await assert.rejects(createOfficeStreamFn(() => Promise.resolve(undefined), runtime.models)(
        model, { messages: [] },
      ), /requires the Pi for Excel proxy/u);
    } else {
      assert.equal(requestUrl, "https://openrouter.ai/api/v1/chat/completions");
    }
  }
});

void test("Office streaming sends every Copilot API through the configured proxy", async (t) => {
  const originalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
  Object.defineProperty(globalThis, "document", { configurable: true, value: new EventTarget() });
  t.after(() => {
    if (originalDocument) Object.defineProperty(globalThis, "document", originalDocument);
    else Reflect.deleteProperty(globalThis, "document");
  });
  const runtime = createRuntime();
  const stream = createOfficeStreamFn(() => Promise.resolve("https://localhost:3003"), runtime.models);
  for (const [id, path, plan] of [
    ["gpt-5.4-nano", "/responses", "individual"],
    ["claude-haiku-4.5", "/v1/messages?beta=true", "individual"],
    ["gpt-6-luna", "/responses", "individual"],
    ["gemini-3.8-flash", "/chat/completions", "individual"],
    ["gpt-5.4-nano", "/responses", "business"],
    ["gpt-5.4-nano", "/responses", "enterprise"],
  ] as const) {
    const model = runtime.models.getModel("github-copilot", id);
    assert.ok(model);
    const baseUrl = `https://api.${plan}.githubcopilot.com`;
    let requested = "";
    const events = await stream({ ...model, baseUrl }, { messages: [{ role: "user", content: "test", timestamp: 0 }] }, {
      apiKey: "synthetic-copilot-token", maxTokens: 32,
      fetch: (url) => {
        requested = typeof url === "string" ? url : url instanceof URL ? url.href : url.url;
        return Promise.resolve(new Response("test endpoint", { status: 400 }));
      },
    });
    await events.result();
    const request = new URL(requested);
    assert.equal(request.origin, "https://localhost:3003");
    assert.equal(request.searchParams.get("url"), `${baseUrl}${path}`);
  }
});

void test("Office continuation after Stop returns an aborted result without dispatching a request", async () => {
  const runtime = createRuntime();
  const model = runtime.models.getModel("openrouter", "openai/gpt-5.6-sol");
  assert.ok(model);
  const controller = new AbortController();
  controller.abort();
  const stream = createOfficeStreamFn(() => {
    assert.fail("Stopped continuations must not resolve a proxy or send a request");
  }, runtime.models);
  const events = await stream(model, { messages: [] }, {
    signal: controller.signal,
    fetch: () => { assert.fail("Stopped continuations must not fetch"); },
  });
  const result = await events.result();
  assert.equal(result.stopReason, "aborted");
  assert.equal(result.errorMessage, "Request aborted");
  assert.deepEqual(result.content, []);
});

function createRuntime(args?: {
  providerKeys?: MemoryProviderKeys;
  catalogs?: ModelsStore;
  fetchFn?: typeof globalThis.fetch;
  modelCatalogUrl?: (providerId: string) => string;
}): BrowserModelRuntime {
  return new BrowserModelRuntime({
    providerKeys: args?.providerKeys ?? new MemoryProviderKeys(),
    modelCatalogs: args?.catalogs ?? new MemoryCatalogs(),
    getProxyUrl: () => Promise.resolve(undefined),
    ...(args?.fetchFn !== undefined ? { fetchFn: args.fetchFn } : {}),
    ...(args?.modelCatalogUrl !== undefined ? { modelCatalogUrl: args.modelCatalogUrl } : {}),
  });
}

class GatedSyncBrowserModelRuntime extends BrowserModelRuntime {
  private firstSyncPending = true;
  private readonly firstSyncGate: Promise<void>;
  private notifyFirstSyncStarted: () => void = () => {};
  private releaseFirstSyncGate: () => void = () => {};
  readonly firstSyncStarted = new Promise<void>((resolve) => {
    this.notifyFirstSyncStarted = resolve;
  });

  constructor(fetchFn?: typeof globalThis.fetch) {
    super({
      providerKeys: new MemoryProviderKeys(),
      modelCatalogs: new MemoryCatalogs(),
      getProxyUrl: () => Promise.resolve(undefined),
      ...(fetchFn !== undefined ? { fetchFn } : {}),
    });
    this.firstSyncGate = new Promise<void>((resolve) => {
      this.releaseFirstSyncGate = resolve;
    });
  }

  releaseFirstSync(): void {
    this.releaseFirstSyncGate();
  }

  override async syncCustomProviders(customProviders: readonly CustomProvider[]): Promise<void> {
    if (this.firstSyncPending) {
      this.firstSyncPending = false;
      this.notifyFirstSyncStarted();
      await this.firstSyncGate;
    }
    await super.syncCustomProviders(customProviders);
  }
}

function gatewayProvider(options?: {
  id?: string;
  name?: string;
  providerId?: string;
  baseUrl?: string;
  apiKey?: string;
  connectDirectly?: boolean;
}): CustomProvider {
  const baseUrl = options?.baseUrl ?? "https://gateway.example.com/v1";
  return {
    id: options?.id ?? "stored-gateway",
    name: options?.name ?? "Acme gateway",
    type: "openai-completions",
    baseUrl,
    apiKey: options?.apiKey ?? "gateway-secret",
    ...(options?.connectDirectly === true ? { connectDirectly: true } : {}),
    models: [{
      id: "configured-model",
      name: "Configured model",
      api: "openai-completions",
      provider: options?.providerId ?? "Gateway · Acme",
      baseUrl,
      reasoning: false,
      input: ["text"],
      inputLimits: {
        maxRequestBytes: 8_388_608,
        images: {
          maxPerRequest: 12,
          resize: { maxWidth: 640, maxHeight: 480, maxBytes: 524_288, jpegQuality: 71 },
        },
      },
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

void test("discovery-enabled custom providers merge remote model ids and persist the catalogue", async () => {
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
  assert.equal(runtime.providerRoute("Gateway · Acme"), "proxy");
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
  assert.deepEqual(
    runtime.models.getModel("Gateway · Acme", "remote-a")?.inputLimits,
    gatewayProvider().models?.[0]?.inputLimits,
  );
});

void test("a gateway saved through the manual form uses only its configured model, including after reload", async () => {
  const providers = new Map<string, CustomProvider>();
  const registry = createRuntime().models;
  await saveOpenAiGatewayConfig({
    get: (id) => Promise.resolve(providers.get(id) ?? null),
    set: (provider) => { providers.set(provider.id, provider); return Promise.resolve(); },
    delete: (id) => { providers.delete(id); return Promise.resolve(); },
    getAll: () => Promise.resolve([...providers.values()]),
  }, {
    displayName: "Manual gateway",
    endpointUrl: "https://openrouter.ai/api/v1",
    modelId: "openai/gpt-5.6-sol",
    apiKey: "synthetic-key",
  }, registry);
  const provider = [...providers.values()][0];
  const model = provider?.models?.[0];
  assert.ok(provider);
  assert.ok(model);
  const catalogs = new MemoryCatalogs();
  catalogs.entries.set(model.provider, { models: [{ ...model, id: "formerly-discovered" }] });
  let requests = 0;
  const fetchFn: typeof fetch = () => {
    requests += 1;
    return Promise.resolve(new Response("not supported", { status: 404 }));
  };
  for (let boot = 0; boot < 2; boot += 1) {
    const runtime = createRuntime({ catalogs, fetchFn });
    await runtime.syncCustomProviders([provider]);
    await runtime.refresh({ allowNetwork: false });
    const result = await runtime.refresh({ allowNetwork: true, force: true });
    assert.equal(result.errors.size, 0);
    assert.deepEqual((await runtime.models.getAvailable(model.provider)).map((entry) => entry.id), [model.id]);
    const registered = runtime.models.getModel(model.provider, model.id);
    const known = registry.getModel("openrouter", model.id);
    assert.ok(known);
    assert.equal(registered?.contextWindow, known.contextWindow);
    assert.equal(registered?.maxTokens, known.maxTokens);
    assert.deepEqual(registered?.thinkingLevelMap, known.thinkingLevelMap);
    assert.deepEqual(registered?.cost, known.cost);
    assert.equal((await runtime.models.getAuth(model.provider))?.auth.apiKey, "synthetic-key");
  }
  assert.equal(requests, 0);
});

void test("provider discovery distinguishes omitted, explicit and null URLs", async () => {
  for (const modelsUrl of [undefined, "https://catalog.example.com/models", null]) {
    const requests: string[] = [];
    const runtime = createRuntime({ fetchFn: (input) => {
      requests.push(typeof input === "string" ? input : input instanceof URL ? input.href : input.url);
      return Promise.resolve(Response.json({ data: [{ id: "discovered" }] }));
    } });
    runtime.registerExtensionProvider("test-owner", {
      id: "test-provider",
      name: "Test provider",
      api: "openai-completions",
      baseUrl: "https://gateway.example.com/v1",
      models: [{ id: "baseline" }],
      resolveApiKey: () => Promise.resolve("synthetic-key"),
      ...(modelsUrl !== undefined ? { modelsUrl } : {}),
    });
    await runtime.refresh({ allowNetwork: true });
    assert.deepEqual(requests, modelsUrl === null ? [] : [modelsUrl ?? "https://gateway.example.com/v1/models"]);
    assert.deepEqual(runtime.models.getModels("test-provider").map((model) => model.id),
      modelsUrl === null ? ["baseline"] : ["baseline", "discovered"]);
  }
});

void test("a fresh runtime restores discovered models without network access", async () => {
  const catalogs = new MemoryCatalogs();
  const fetchFn: typeof globalThis.fetch = () => {
    return Promise.resolve(new Response(JSON.stringify({ data: [{ id: "cached-model" }] }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }));
  };

  const first = createRuntime({ catalogs, fetchFn });
  await first.syncCustomProviders([gatewayProvider()]);
  await first.refresh({ allowNetwork: true });

  const second = createRuntime({
    catalogs,
    fetchFn: () => {
      throw new Error("cache-only startup must not call the gateway");
    },
  });
  await second.syncCustomProviders([gatewayProvider()]);
  await second.refresh({ allowNetwork: false });

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
  assert.deepEqual(restored?.inputLimits, updated.models?.[0]?.inputLimits);
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

void test("oversized or malformed discovery catalogues retain the last safe cache", async () => {
  const catalogs = new MemoryCatalogs();
  let mode: "safe" | "too-many" | "long-id" | "too-large" = "safe";
  const fetchFn: typeof globalThis.fetch = () => {
    if (mode === "too-many") {
      return Promise.resolve(new Response(JSON.stringify({
        data: Array.from({ length: 2_001 }, (_, index) => ({ id: `model-${index}` })),
      }), { status: 200, headers: { "Content-Type": "application/json" } }));
    }
    if (mode === "long-id") {
      return Promise.resolve(new Response(JSON.stringify({
        data: [{ id: "x".repeat(257) }],
      }), { status: 200, headers: { "Content-Type": "application/json" } }));
    }
    if (mode === "too-large") {
      return Promise.resolve(new Response("{}", {
        status: 200,
        headers: {
          "Content-Type": "application/json",
          "Content-Length": String(2 * 1024 * 1024 + 1),
        },
      }));
    }
    return Promise.resolve(new Response(JSON.stringify({ data: [{ id: "last-safe-model" }] }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }));
  };

  const runtime = createRuntime({ catalogs, fetchFn });
  await runtime.syncCustomProviders([gatewayProvider()]);
  await runtime.refresh({ allowNetwork: true });

  for (const unsafeMode of ["too-many", "long-id", "too-large"] as const) {
    mode = unsafeMode;
    const failed = await runtime.refresh({ allowNetwork: true });
    assert.equal(failed.errors.has("Gateway · Acme"), true, unsafeMode);
    assert.deepEqual(
      runtime.models.getModels("Gateway · Acme").map((model) => model.id),
      ["configured-model", "last-safe-model"],
      unsafeMode,
    );
  }
});

void test("extension providers are owner-scoped and can resolve host-owned credentials", async () => {
  const runtime = createRuntime();
  const registration: BrowserProviderRegistration = {
    id: "ext.example.provider",
    name: "Example provider",
    api: "openai-responses",
    baseUrl: "https://models.example.com/v1",
    models: [{
      id: "example-model",
      contextWindow: 128_000,
      maxTokens: 16_000,
      inputLimits: { images: { resize: { maxWidth: 1_024 } } },
    }],
    modelsUrl: "https://models.example.com/v1/models",
    resolveApiKey: () => Promise.resolve("host-owned-secret"),
  };

  runtime.registerExtensionProvider("ext.example", registration);
  assert.equal(runtime.isExtensionProvider(registration.id), true);
  assert.equal(runtime.providerRoute(registration.id), "proxy");
  assert.equal(runtime.providerRoute("openai"), undefined);
  assert.equal((await runtime.models.getAuth(registration.id))?.auth.apiKey, "host-owned-secret");
  const extensionModel = (await runtime.models.getAvailable(registration.id)).find(
    (model) => model.id === "example-model",
  );
  assert.deepEqual(extensionModel?.inputLimits, registration.models[0]?.inputLimits);

  await assert.rejects(
    runtime.unregisterExtensionProvider("ext.someone-else", registration.id),
    /not owned/,
  );
  await runtime.unregisterExtensionProvider("ext.example", registration.id);
  assert.equal(runtime.models.getProvider(registration.id), undefined);
  assert.equal(runtime.providerRoute(registration.id), undefined);
});

void test("unregister aborts in-flight discovery before it can resurrect a catalogue", async () => {
  const catalogs = new MemoryCatalogs();
  let mode: "delayed" | "offline" = "delayed";
  let notifyStarted: () => void = () => {};
  let resolveDelayed: (response: Response) => void = () => {};
  const started = new Promise<void>((resolve) => {
    notifyStarted = resolve;
  });
  const fetchFn: typeof globalThis.fetch = (_input, init) => {
    if (mode === "offline") {
      return Promise.resolve(new Response("offline", { status: 503 }));
    }

    return new Promise<Response>((resolve, reject) => {
      resolveDelayed = resolve;
      const abort = (): void => reject(new DOMException("aborted", "AbortError"));
      if (init?.signal?.aborted) {
        abort();
        return;
      }
      init?.signal?.addEventListener("abort", abort, { once: true });
      notifyStarted();
    });
  };
  const runtime = createRuntime({ catalogs, fetchFn });
  const registration: BrowserProviderRegistration = {
    id: "ext.race.provider",
    name: "Race provider",
    api: "openai-completions",
    baseUrl: "https://race.example.com/v1",
    models: [{ id: "baseline-model", contextWindow: 32_768, maxTokens: 4_096 }],
    resolveApiKey: () => Promise.resolve("test-key"),
  };

  runtime.registerExtensionProvider("ext.race", registration);
  const refresh = runtime.refresh({ allowNetwork: true });
  await started;
  await runtime.unregisterExtensionProvider("ext.race", registration.id);
  resolveDelayed(new Response(JSON.stringify({ data: [{ id: "must-not-resurrect" }] }), {
    status: 200,
    headers: { "Content-Type": "application/json" },
  }));
  await refresh;

  assert.equal(catalogs.entries.has(registration.id), false);

  mode = "offline";
  runtime.registerExtensionProvider("ext.race", registration);
  await runtime.refresh({ allowNetwork: true });
  assert.deepEqual(
    runtime.models.getModels(registration.id).map((model) => model.id),
    ["baseline-model"],
  );
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

function waitForOwnerRevision(owner: ModelRefreshOwner, revision: number): Promise<void> {
  if (owner.snapshot().revision >= revision) return Promise.resolve();

  return new Promise((resolve) => {
    const unsubscribe = owner.subscribe((snapshot) => {
      if (snapshot.revision < revision) return;
      unsubscribe();
      resolve();
    });
  });
}

void test("concurrent model refresh callers receive the discovered catalogue", async () => {
  let notifyFetchStarted: () => void = () => {};
  const fetchStarted = new Promise<void>((resolve) => {
    notifyFetchStarted = resolve;
  });
  let releaseFetch: () => void = () => {};
  const fetchGate = new Promise<void>((resolve) => {
    releaseFetch = resolve;
  });
  const fetchFn: typeof globalThis.fetch = async () => {
    notifyFetchStarted();
    await fetchGate;
    return new Response(JSON.stringify({ data: [{ id: "remote-model" }] }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  };
  const runtime = createRuntime({ fetchFn });
  const owner = new ModelRefreshOwner({
    modelRuntime: runtime,
    loadCustomProviders: () => Promise.resolve([gatewayProvider()]),
    getRuntimes: () => [],
  });

  await owner.restoreCached();
  const first = owner.refresh(true);
  const second = owner.refresh(true);
  assert.equal(first, second, "concurrent callers should share the refresh result");
  await fetchStarted;

  releaseFetch();
  await Promise.all([first, second]);
  assert.ok(runtime.models.getModel("Gateway · Acme", "remote-model"));
});

void test("configured-provider refresh during active sync reloads the latest configuration", async () => {
  const firstProvider = gatewayProvider({
    id: "first",
    name: "First",
    providerId: "Gateway · First",
    baseUrl: "https://first.example.com/v1",
  });
  const latestProvider = gatewayProvider({
    id: "latest",
    name: "Latest",
    providerId: "Gateway · Latest",
    baseUrl: "https://latest.example.com/v1",
  });
  let configuredProviders: readonly CustomProvider[] = [firstProvider];
  const runtime = new GatedSyncBrowserModelRuntime(() => {
    return Promise.resolve(new Response(JSON.stringify({ data: [] }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }));
  });
  const owner = new ModelRefreshOwner({
    modelRuntime: runtime,
    loadCustomProviders: () => Promise.resolve(configuredProviders),
    getRuntimes: () => [],
  });

  const firstRefresh = owner.refreshConfiguredProviders(false);
  await runtime.firstSyncStarted;
  configuredProviders = [latestProvider];
  const followUpRefresh = owner.refreshConfiguredProviders(true);
  const coalescedRefresh = owner.refreshConfiguredProviders(true);
  runtime.releaseFirstSync();
  await Promise.all([firstRefresh, followUpRefresh, coalescedRefresh]);

  assert.equal(runtime.models.getProvider("Gateway · First"), undefined);
  assert.ok(runtime.models.getProvider("Gateway · Latest"));
});

void test("configured-provider refresh during active network discovery is not lost", async () => {
  const firstProvider = gatewayProvider({
    id: "first",
    name: "First",
    providerId: "Gateway · First",
    baseUrl: "https://first.example.com/v1",
  });
  const latestProvider = gatewayProvider({
    id: "latest",
    name: "Latest",
    providerId: "Gateway · Latest",
    baseUrl: "https://latest.example.com/v1",
  });
  let configuredProviders: readonly CustomProvider[] = [firstProvider];
  let notifyFetchStarted: () => void = () => {};
  const fetchStarted = new Promise<void>((resolve) => {
    notifyFetchStarted = resolve;
  });
  let releaseFetch: () => void = () => {};
  const fetchGate = new Promise<void>((resolve) => {
    releaseFetch = resolve;
  });
  const runtime = createRuntime({
    fetchFn: async () => {
      notifyFetchStarted();
      await fetchGate;
      return new Response(JSON.stringify({ data: [] }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    },
  });
  const owner = new ModelRefreshOwner({
    modelRuntime: runtime,
    loadCustomProviders: () => Promise.resolve(configuredProviders),
    getRuntimes: () => [],
  });

  await owner.refreshConfiguredProviders(false);
  const networkRefresh = owner.refresh(true);
  await fetchStarted;
  configuredProviders = [latestProvider];
  const configuredRefresh = owner.refreshConfiguredProviders(false);
  releaseFetch();
  await Promise.all([networkRefresh, configuredRefresh]);

  assert.equal(runtime.models.getProvider("Gateway · First"), undefined);
  assert.ok(runtime.models.getProvider("Gateway · Latest"));
});

void test("model refresh owner preserves healthy provider models on partial failure", async () => {
  const healthy = gatewayProvider({
    id: "healthy",
    name: "Healthy",
    providerId: "Gateway · Healthy",
    baseUrl: "https://healthy.example.com/v1",
  });
  const failing = gatewayProvider({
    id: "failing",
    name: "Failing",
    providerId: "Gateway · Failing",
    baseUrl: "https://failing.example.com/v1",
  });
  const fetchFn: typeof globalThis.fetch = (input) => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.toString() : input.url;
    if (new URL(url).hostname === "failing.example.com") {
      return Promise.resolve(new Response("unavailable", { status: 503 }));
    }
    return Promise.resolve(new Response(JSON.stringify({ data: [{ id: "healthy-remote" }] }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }));
  };
  const runtime = createRuntime({ fetchFn });
  const owner = new ModelRefreshOwner({
    modelRuntime: runtime,
    loadCustomProviders: () => Promise.resolve([healthy, failing]),
    getRuntimes: () => [],
  });

  await owner.refreshConfiguredProviders(true);

  assert.deepEqual(owner.snapshot().providerErrors, ["Gateway · Failing"]);
  assert.deepEqual(
    runtime.models.getModels("Gateway · Healthy").map((model) => model.id),
    ["configured-model", "healthy-remote"],
  );
  assert.ok(owner.snapshot().availableProviders.includes("Gateway · Healthy"));
});

void test("late credentials update idle runtimes and defer busy runtimes", async () => {
  const providerKeys = new MemoryProviderKeys();
  const runtime = createRuntime({ providerKeys });
  let restoreCredentials: () => void = () => {};
  const credentialRestore = new Promise<void>((resolve) => {
    restoreCredentials = resolve;
  });
  let idleModel = getBuiltinModel("openai", "gpt-5.6-sol");
  let busyModel = getBuiltinModel("openai", "gpt-5.6-sol");
  let busy = true;
  const targets: ModelRefreshRuntime[] = [
    {
      runtimeId: "idle",
      model: idleModel,
      isBusy: false,
      applyModel: (model) => {
        idleModel = model;
        targets[0].model = model;
      },
    },
    {
      runtimeId: "busy",
      model: busyModel,
      isBusy: busy,
      applyModel: (model) => {
        busyModel = model;
        targets[1].model = model;
      },
    },
  ];
  const owner = new ModelRefreshOwner({
    modelRuntime: runtime,
    loadCustomProviders: () => Promise.resolve([]),
    getRuntimes: () => {
      targets[1].isBusy = busy;
      return targets;
    },
  });

  await owner.startup(credentialRestore, 1);
  await waitForOwnerRevision(owner, 2);

  await providerKeys.set("openai-codex", "late-key");
  restoreCredentials();
  for (let attempt = 0; attempt < 50 && idleModel.provider !== "openai-codex"; attempt += 1) {
    await new Promise<void>((resolve) => setTimeout(resolve, 1));
  }

  assert.equal(idleModel.provider, "openai-codex");
  assert.equal(busyModel.provider, "openai");

  busy = false;
  owner.reconcileRuntimes();
  assert.equal(busyModel.provider, "openai-codex");
});

void test("a provider registered during a refresh pass gets its cached catalogue restored", async () => {
  const providerId = "ext.example.probe";
  const catalogs = new MemoryCatalogs();
  catalogs.entries.set(providerId, {
    models: [{
      id: "cached-extension-model",
      name: "Cached extension model",
      api: "openai-completions",
      provider: providerId,
      baseUrl: "https://old-extension.example.com/v1",
      reasoning: false,
      input: ["text"],
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
      contextWindow: 32_768,
      maxTokens: 4_096,
    }],
  });
  let notifyGatewayDiscoveryStarted: () => void = () => {};
  const gatewayDiscoveryStarted = new Promise<void>((resolve) => {
    notifyGatewayDiscoveryStarted = resolve;
  });
  let releaseGatewayDiscovery: () => void = () => {};
  const gatewayDiscoveryGate = new Promise<void>((resolve) => {
    releaseGatewayDiscovery = resolve;
  });
  const runtime = createRuntime({
    catalogs,
    fetchFn: async (input) => {
      const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
      if (url.startsWith("https://new-extension.example.com/")) {
        throw new Error("new extension endpoint is offline");
      }
      notifyGatewayDiscoveryStarted();
      await gatewayDiscoveryGate;
      return new Response(JSON.stringify({ data: [{ id: "network-model" }] }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    },
  });
  const owner = new ModelRefreshOwner({
    modelRuntime: runtime,
    loadCustomProviders: () => Promise.resolve([gatewayProvider()]),
    getRuntimes: () => [],
  });

  await owner.refreshConfiguredProviders(false);
  const startupNetworkPass = owner.refresh(true);
  await gatewayDiscoveryStarted;

  // An extension activating mid-pass registers its provider, then asks for a refresh.
  runtime.registerExtensionProvider("ext.example", {
    id: providerId,
    name: "Example extension provider",
    api: "openai-completions",
    baseUrl: "https://new-extension.example.com/v1",
    models: [{ id: "baseline" }],
    resolveApiKey: () => Promise.resolve("host-owned-secret"),
  });
  const activationRefresh = owner.refresh(true);
  releaseGatewayDiscovery();
  await Promise.all([startupNetworkPass, activationRefresh]);

  const restored = runtime.models.getModel(providerId, "cached-extension-model");
  assert.ok(restored, "the cached extension model should survive the coalesced refresh");
  assert.equal(restored.baseUrl, "https://new-extension.example.com/v1");
  assert.deepEqual(owner.snapshot().providerErrors, [providerId]);
});

void test("model refresh owner publishes cached startup before network discovery", async () => {
  const catalogs = new MemoryCatalogs();
  catalogs.entries.set("Gateway · Acme", {
    models: [{
      id: "cached-before-network",
      name: "Cached before network",
      api: "openai-completions",
      provider: "Gateway · Acme",
      baseUrl: "https://gateway.example.com/v1",
      reasoning: false,
      input: ["text"],
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
      contextWindow: 32_768,
      maxTokens: 4_096,
    }],
  });
  let releaseFetch: () => void = () => {};
  const fetchGate = new Promise<void>((resolve) => {
    releaseFetch = resolve;
  });
  const runtime = createRuntime({
    catalogs,
    fetchFn: async () => {
      await fetchGate;
      return new Response(JSON.stringify({ data: [{ id: "network-model" }] }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    },
  });
  const owner = new ModelRefreshOwner({
    modelRuntime: runtime,
    loadCustomProviders: () => Promise.resolve([gatewayProvider()]),
    getRuntimes: () => [],
  });

  await owner.startup(Promise.resolve(), 50);

  assert.equal(owner.snapshot().revision, 1);
  assert.ok(runtime.models.getModel("Gateway · Acme", "cached-before-network"));
  assert.equal(runtime.models.getModel("Gateway · Acme", "network-model"), undefined);

  // The network pass republishes the cached state, then the discovered catalogue.
  releaseFetch();
  await waitForOwnerRevision(owner, 3);
  assert.ok(runtime.models.getModel("Gateway · Acme", "network-model"));
});

function installDocumentShim(t: TestContext): void {
  const originalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
  Object.defineProperty(globalThis, "document", { configurable: true, value: new EventTarget() });
  t.after(() => {
    if (originalDocument) Object.defineProperty(globalThis, "document", originalDocument);
    else Reflect.deleteProperty(globalThis, "document");
  });
}

async function streamGatewayRequest(args: {
  runtime: BrowserModelRuntime;
  proxyUrl: string | undefined;
}): Promise<{ url: string; authorization: string }> {
  const model = args.runtime.models.getModel("Gateway · Acme", "configured-model");
  assert.ok(model);
  const auth = await args.runtime.models.getAuth("Gateway · Acme");
  const request = { url: "", authorization: "" };
  const stream = createOfficeStreamFn(
    () => Promise.resolve(args.proxyUrl),
    args.runtime.models,
    (providerId) => args.runtime.providerRoute(providerId),
  );
  const events = await stream(model, { messages: [{ role: "user", content: "test", timestamp: 0 }] }, {
    ...(auth?.auth.apiKey !== undefined ? { apiKey: auth.auth.apiKey } : {}),
    fetch: (url, init) => {
      request.url = typeof url === "string" ? url : url instanceof URL ? url.href : url.url;
      request.authorization = new Headers(init?.headers).get("authorization") ?? "";
      return Promise.resolve(new Response("test endpoint", { status: 400 }));
    },
  });
  await events.result();
  return request;
}

void test("keyless gateways reach the endpoint instead of failing for a missing API key", async (t) => {
  installDocumentShim(t);
  const runtime = createRuntime();
  await runtime.syncCustomProviders([gatewayProvider({ apiKey: "", baseUrl: "https://localhost:8090/v1" })]);

  const request = await streamGatewayRequest({ runtime, proxyUrl: undefined });

  assert.equal(request.url, "https://localhost:8090/v1/chat/completions");
  assert.equal(request.authorization, "Bearer pi-for-excel-keyless");
});

void test("gateways set to connect directly skip an enabled proxy", async (t) => {
  installDocumentShim(t);
  const proxied = createRuntime();
  await proxied.syncCustomProviders([gatewayProvider()]);
  assert.equal(proxied.providerRoute("Gateway · Acme"), "proxy");
  assert.match(
    (await streamGatewayRequest({ runtime: proxied, proxyUrl: "https://localhost:3004" })).url,
    /^https:\/\/localhost:3004\/\?url=/u,
  );

  const direct = createRuntime();
  await direct.syncCustomProviders([gatewayProvider({ connectDirectly: true })]);
  assert.equal(direct.providerRoute("Gateway · Acme"), "direct");
  assert.equal(
    (await streamGatewayRequest({ runtime: direct, proxyUrl: "https://localhost:3004" })).url,
    "https://gateway.example.com/v1/chat/completions",
  );

  await direct.syncCustomProviders([gatewayProvider()]);
  assert.equal(direct.providerRoute("Gateway · Acme"), "proxy");
});

void test("direct plain-http requests to another machine explain the mixed-content block", async (t) => {
  installDocumentShim(t);
  const remote = createRuntime();
  await remote.syncCustomProviders([gatewayProvider({ baseUrl: "http://192.168.1.50:8080/v1" })]);
  await assert.rejects(
    streamGatewayRequest({ runtime: remote, proxyUrl: undefined }),
    /plain http:\/\/ endpoint .*Excel blocks http:\/\/ requests/u,
  );

  const loopback = createRuntime();
  await loopback.syncCustomProviders([gatewayProvider({ baseUrl: "http://localhost:8080/v1" })]);
  assert.equal(
    (await streamGatewayRequest({ runtime: loopback, proxyUrl: undefined })).url,
    "http://localhost:8080/v1/chat/completions",
  );
});

const CATALOG_RELAY = (providerId: string): string => `https://relay.test/model-catalog/${providerId}`;
const FUTURE_LAST_MODIFIED = "Wed, 01 Jan 2031 00:00:00 GMT";

function remoteAnthropicModel(overrides: Record<string, unknown>): Record<string, unknown> {
  return {
    id: "claude-next",
    name: "Claude Next",
    api: "anthropic-messages",
    provider: "anthropic",
    baseUrl: "https://api.anthropic.com",
    reasoning: true,
    input: ["text", "image"],
    cost: { input: 3, output: 15, cacheRead: 0.3, cacheWrite: 3.75 },
    contextWindow: 1_000_000,
    maxTokens: 128_000,
    ...overrides,
  };
}

function catalogResponse(models: readonly unknown[], headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(models), {
    status: 200,
    headers: {
      "Content-Type": "application/json",
      "Last-Modified": FUTURE_LAST_MODIFIED,
      ETag: "\"v1\"",
      ...headers,
    },
  });
}

function persistentCatalogs(): ModelCatalogsStore {
  const store = new ModelCatalogsStore();
  store.setBackend(new IndexedDBStorageBackend(getAppStorageConfig(`catalog-test-${crypto.randomUUID()}`)));
  return store;
}

void test("built-in providers add pi.dev catalogue models but keep their own transport", async () => {
  const providerKeys = new MemoryProviderKeys();
  await providerKeys.set("anthropic", "test-key");
  const catalogs = persistentCatalogs();
  const requested: string[] = [];
  const validators: (string | null)[] = [];
  const bundledId = "claude-fable-5";
  const fetchFn: typeof globalThis.fetch = (input, init) => {
    requested.push(input instanceof Request ? input.url : input.toString());
    validators.push(new Headers(init?.headers).get("if-none-match"));
    return Promise.resolve(catalogResponse([
      remoteAnthropicModel({
        headers: { "x-leak": "credential" },
        compat: { forceAdaptiveThinking: true },
        thinkingLevelMap: { off: null, max: "max" },
        cost: {
          input: 3, output: 15, cacheRead: 0.3, cacheWrite: 3.75,
          tiers: [{ input: 6, output: 22.5, cacheRead: 0.6, cacheWrite: 7.5, inputTokensAbove: 200_000 }],
        },
      }),
      remoteAnthropicModel({ id: "claude-elsewhere", baseUrl: "https://collector.example.com" }),
      remoteAnthropicModel({ id: "claude-new-transport", api: "anthropic-next" }),
      remoteAnthropicModel({ id: "claude-broken", contextWindow: "lots" }),
      remoteAnthropicModel({ id: bundledId, name: "Claude Fable 5 (updated)" }),
    ]));
  };

  const runtime = createRuntime({ providerKeys, catalogs, fetchFn, modelCatalogUrl: CATALOG_RELAY });
  const result = await runtime.refresh({ allowNetwork: true, force: true });

  assert.equal(result.errors.size, 0);
  // Only configured providers are fetched, through the relay.
  assert.deepEqual(requested, ["https://relay.test/model-catalog/anthropic"]);
  assert.deepEqual(validators, [null]);

  const added = runtime.models.getModel("anthropic", "claude-next");
  assert.ok(added);
  assert.equal(added.baseUrl, "https://api.anthropic.com");
  assert.equal(added.headers, undefined);
  assert.deepEqual(added.compat, { forceAdaptiveThinking: true });
  assert.deepEqual(added.thinkingLevelMap, { off: null, max: "max" });
  assert.equal(added.cost.tiers?.[0]?.inputTokensAbove, 200_000);
  assert.equal(runtime.models.getModel("anthropic", bundledId)?.name, "Claude Fable 5 (updated)");
  for (const dropped of ["claude-elsewhere", "claude-new-transport", "claude-broken"]) {
    assert.equal(runtime.models.getModel("anthropic", dropped), undefined, dropped);
  }

  // A fresh taskpane restores the persisted overlay offline, with every field intact.
  const offline = createRuntime({
    providerKeys,
    catalogs,
    fetchFn: () => Promise.reject(new Error("cache-only restore must not fetch")),
    modelCatalogUrl: CATALOG_RELAY,
  });
  await offline.refresh({ allowNetwork: false });
  assert.deepEqual(offline.models.getModel("anthropic", "claude-next"), added);

  // The next network refresh revalidates; 304 keeps the overlay.
  const revalidated = createRuntime({
    providerKeys,
    catalogs,
    fetchFn: (_input, init) => {
      validators.push(new Headers(init?.headers).get("if-none-match"));
      return Promise.resolve(new Response(null, { status: 304 }));
    },
    modelCatalogUrl: CATALOG_RELAY,
  });
  await revalidated.refresh({ allowNetwork: true, force: true });
  assert.equal(validators.at(-1), "\"v1\"");
  assert.ok(revalidated.models.getModel("anthropic", "claude-next"));
});

void test("an unusable pi.dev catalogue leaves bundled and cached models in place", async () => {
  const providerKeys = new MemoryProviderKeys();
  await providerKeys.set("anthropic", "test-key");
  const catalogs = new MemoryCatalogs();
  const bundledIds = (runtime: BrowserModelRuntime): string[] =>
    runtime.models.getModels("anthropic").map((model) => model.id).sort();
  const baseline = bundledIds(createRuntime({ providerKeys }));

  const responses: Response[] = [
    // Self-hosted without the relay.
    new Response("Not found", { status: 404 }),
    // Requires a newer pi-ai than this build ships.
    catalogResponse([remoteAnthropicModel({ id: "needs-newer-pi" })], {
      "x-pi-model-catalog-minimum-version": "999.0.0",
    }),
    // Older than the bundled model data.
    catalogResponse([remoteAnthropicModel({ id: "stale-model" })], {
      "Last-Modified": "Thu, 01 Jan 2015 00:00:00 GMT",
    }),
  ];
  for (const response of responses) {
    const runtime = createRuntime({
      providerKeys,
      catalogs,
      fetchFn: () => Promise.resolve(response),
      modelCatalogUrl: CATALOG_RELAY,
    });
    const result = await runtime.refresh({ allowNetwork: true, force: true });
    assert.equal(result.errors.size, 0);
    assert.deepEqual(bundledIds(runtime), baseline);
  }

  // A good catalogue, then a failing request: the last good overlay survives.
  const good = createRuntime({
    providerKeys,
    catalogs,
    fetchFn: () => Promise.resolve(catalogResponse([remoteAnthropicModel({})])),
    modelCatalogUrl: CATALOG_RELAY,
  });
  await good.refresh({ allowNetwork: true, force: true });
  const failing = createRuntime({
    providerKeys,
    catalogs,
    fetchFn: () => Promise.resolve(new Response("<html>", { status: 502 })),
    modelCatalogUrl: CATALOG_RELAY,
  });
  const result = await failing.refresh({ allowNetwork: true, force: true });
  assert.equal(result.errors.size, 1);
  assert.ok(failing.models.getModel("anthropic", "claude-next"));
});

void test("a newly connected provider is available before its catalogue request settles", async () => {
  const providerKeys = new MemoryProviderKeys();
  let releaseFetch: () => void = () => {};
  const fetchGate = new Promise<void>((resolve) => {
    releaseFetch = resolve;
  });
  const runtime = createRuntime({
    providerKeys,
    modelCatalogUrl: CATALOG_RELAY,
    fetchFn: async () => {
      await fetchGate;
      return catalogResponse([remoteAnthropicModel({})]);
    },
  });
  const owner = new ModelRefreshOwner({
    modelRuntime: runtime,
    loadCustomProviders: () => Promise.resolve([]),
    getRuntimes: () => [],
  });
  await owner.startup(Promise.resolve(), 50);
  // Let startup's background pass settle so only the connection refresh can publish.
  await owner.refresh(false);
  assert.deepEqual(owner.snapshot().availableProviders, []);

  await providerKeys.set("anthropic", "test-key");
  const refresh = owner.refreshConfiguredProviders(true);
  await waitForOwner(owner, (snapshot) => snapshot.availableProviders.includes("anthropic"));
  assert.equal(runtime.models.getModel("anthropic", "claude-next"), undefined);

  releaseFetch();
  await refresh;
  assert.ok(runtime.models.getModel("anthropic", "claude-next"));
});

function waitForOwner(
  owner: ModelRefreshOwner,
  predicate: (snapshot: ReturnType<ModelRefreshOwner["snapshot"]>) => boolean,
): Promise<void> {
  if (predicate(owner.snapshot())) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      unsubscribe();
      reject(new Error("Timed out waiting for the model refresh owner"));
    }, 2_000);
    const unsubscribe = owner.subscribe((snapshot) => {
      if (!predicate(snapshot)) return;
      clearTimeout(timeout);
      unsubscribe();
      resolve();
    });
  });
}
