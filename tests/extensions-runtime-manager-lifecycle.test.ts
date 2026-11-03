import assert from "node:assert/strict";
import { test } from "node:test";
import { Agent } from "@earendil-works/pi-agent-core";
import { InMemoryModelsStore } from "@earendil-works/pi-ai";
import { Type } from "typebox";

import { loadExtension, type ExcelExtensionAPI } from "../src/commands/extension-api.ts";
import { commandRegistry } from "../src/commands/types.ts";
import { ConnectionManager } from "../src/connections/manager.ts";
import { ExtensionRuntimeManager, type ExtensionRuntimeManagerOptions } from "../src/extensions/runtime-manager.ts";
import { EXTENSIONS_REGISTRY_STORAGE_KEY } from "../src/extensions/store.ts";
import { BrowserModelRuntime } from "../src/models/browser-model-runtime.ts";

function deferred() {
  let resolve = (): void => {};
  const promise = new Promise<void>((done) => { resolve = done; });
  return { promise, resolve };
}

function createManager(overrides: Partial<ExtensionRuntimeManagerOptions> = {}) {
  const values = new Map<string, unknown>([[EXTENSIONS_REGISTRY_STORAGE_KEY, { version: 2, items: [] }]]);
  const settings = {
    get: (key: string) => Promise.resolve(values.get(key) ?? null),
    set: (key: string, value: unknown) => { values.set(key, value); return Promise.resolve(); },
  };
  const connectionManager = new ConnectionManager({ settings });
  const manager = new ExtensionRuntimeManager({
    settings,
    connectionManager,
    modelRuntime: new BrowserModelRuntime({
      providerKeys: {
        get: () => Promise.resolve(null), set: () => Promise.resolve(),
        delete: () => Promise.resolve(), list: () => Promise.resolve([]),
      },
      modelCatalogs: new InMemoryModelsStore(),
      getProxyUrl: () => Promise.resolve(undefined),
    }),
    getActiveAgent: () => null,
    refreshRuntimeTools: () => Promise.resolve(),
    refreshRuntimeModels: () => Promise.resolve(),
    reservedToolNames: new Set(),
    ...overrides,
  });
  return { manager, connectionManager };
}

function registerTool(api: ExcelExtensionAPI, name = "lifecycle_tool"): void {
  api.registerTool(name, {
    description: "Lifecycle regression tool", parameters: Type.Object({}),
    execute: () => ({ content: [{ type: "text", text: "ok" }], details: undefined }),
  });
}

for (const registry of ["tools", "models"] as const) {
  void test(`failed ${registry} refresh never leaves an extension listed as loaded`, async () => {
    let cleaned = 0;
    const fail = () => Promise.reject(new Error("refresh unavailable"));
    const { manager } = createManager({
      ...(registry === "tools" ? { refreshRuntimeTools: fail } : { refreshRuntimeModels: fail }),
      loadExtensionFromSource: (api) => {
        if (registry === "tools") registerTool(api);
        else api.models.registerProvider({
          id: "lifecycle", name: "Lifecycle", api: "openai-responses", baseUrl: "https://example.com/v1",
          allowKeyless: true, models: [{ id: "small" }],
        });
        return Promise.resolve({ deactivate: () => { cleaned += 1; return Promise.resolve(); } });
      },
    });
    await manager.initialize();
    await manager.installFromModuleSpecifier("Lifecycle", "./lifecycle.js");
    assert.equal(manager.list()[0]?.loaded, false, "failed refresh must not publish an active instance");
    assert.match(manager.list()[0]?.lastError ?? "", /refresh unavailable/);
    assert.equal(cleaned, 1);
    assert.deepEqual(manager.getRegisteredTools(), []);
  });
}

for (const action of ["disable", "uninstall"] as const) {
  void test(`${action} waits for activation and tears down the instance before returning`, async () => {
    const entered = deferred();
    const release = deferred();
    let cleaned = 0;
    const { manager } = createManager({
      loadExtensionFromSource: async (api) => {
        entered.resolve();
        await release.promise;
        registerTool(api);
        return { deactivate: () => { cleaned += 1; return Promise.resolve(); } };
      },
    });
    await manager.initialize();
    const installing = manager.installFromModuleSpecifier("Lifecycle", "./lifecycle.js");
    await entered.promise;
    const id = manager.list()[0]?.id;
    assert.ok(id);
    const stopping = action === "disable" ? manager.setExtensionEnabled(id, false) : manager.uninstallExtension(id);
    await new Promise<void>((resolve) => setImmediate(resolve));
    release.resolve();
    await Promise.all([installing, stopping]);
    assert.equal(cleaned, 1, "activation must be torn down even when stop arrives before its handle");
    assert.deepEqual(manager.getRegisteredTools(), []);
    assert.equal(manager.list().some((status) => status.loaded), false);
  });
}

void test("initialize is single-flight while stored extensions activate", async () => {
  const entered = deferred();
  const release = deferred();
  let loads = 0;
  const settings = {
    get: (key: string) => Promise.resolve(key === EXTENSIONS_REGISTRY_STORAGE_KEY ? {
      version: 2, items: [{
        id: "ext.lifecycle", name: "Lifecycle", enabled: true,
        source: { kind: "module", specifier: "./lifecycle.js" },
        createdAt: "2026-01-01T00:00:00Z", updatedAt: "2026-01-01T00:00:00Z",
      }],
    } : true),
    set: () => Promise.resolve(),
  };
  const { manager } = createManager({ settings, loadExtensionFromSource: async () => {
    loads += 1;
    entered.resolve();
    await release.promise;
    return { deactivate: () => Promise.resolve() };
  } });
  const first = manager.initialize();
  await entered.promise;
  const second = manager.initialize();
  release.resolve();
  await Promise.all([first, second]);
  assert.equal(loads, 1, "concurrent initialization must activate each extension once");
});

void test("reload serializes generations and rejects registration from retained APIs", async () => {
  const apis: ExcelExtensionAPI[] = [];
  const entered = deferred();
  const release = deferred();
  let loads = 0;
  let cleaned = 0;
  const { manager, connectionManager } = createManager({ loadExtensionFromSource: async (api) => {
    loads += 1;
    apis.push(api);
    if (loads === 2) {
      entered.resolve();
      await release.promise;
    }
    registerTool(api);
    return { deactivate: () => { cleaned += 1; return Promise.resolve(); } };
  } });
  await manager.initialize();
  const id = await manager.installFromModuleSpecifier("Lifecycle", "./lifecycle.js");
  const firstReload = manager.reloadExtension(id);
  await entered.promise;
  const secondReload = manager.reloadExtension(id);
  release.resolve();
  await Promise.all([firstReload, secondReload]);
  assert.equal(cleaned, 2, "each preceding generation must finish teardown before the next registers");
  assert.equal(manager.list()[0]?.loaded, true);
  assert.equal(manager.getRegisteredTools().length, 1);
  const staleApi = apis[0];
  assert.ok(staleApi);
  assert.throws(() => registerTool(staleApi, "stale_tool"), /no longer active|Permission denied/);
  assert.throws(() => staleApi.registerCommand("stale_command", { description: "stale", handler: () => {} }), /no longer active|Permission denied/);
  assert.throws(() => staleApi.connections.register({ id: "stale", title: "Stale", capability: "stale", authKind: "api_key", secretFields: [] }), /no longer active|Permission denied/);
  assert.equal(commandRegistry.get("stale_command"), undefined);
  assert.deepEqual(connectionManager.listRegisteredConnectionIds(), []);
  await manager.uninstallExtension(id);
  const latestApi = apis.at(-1);
  assert.ok(latestApi);
  assert.throws(() => registerTool(latestApi), /no longer active|Permission denied/);
});

void test("retained registration callbacks cannot resurrect an unloaded instance", async () => {
  const retained: { api: ExcelExtensionAPI | null } = { api: null };
  const { manager } = createManager({ loadExtensionFromSource: (api) => {
    retained.api = api;
    return Promise.resolve({ deactivate: () => Promise.resolve() });
  } });
  await manager.initialize();
  const id = await manager.installFromModuleSpecifier("Lifecycle", "./lifecycle.js");
  await manager.setExtensionEnabled(id, false);
  const api = retained.api;
  assert.ok(api);
  assert.throws(() => registerTool(api), /no longer active|Permission denied/, "retained API must reject registration after teardown");
  assert.deepEqual(manager.getRegisteredTools(), []);
});

void test("uninstall retries the failed teardown of a disabled extension", async (t) => {
  let fail = true;
  let cleanups = 0;
  let successfulCleanups = 0;
  let unsubscribes = 0;
  const agent = new Agent({ streamFn: () => { throw new Error("unused"); } });
  t.mock.method(agent, "subscribe", () => () => {
    unsubscribes += 1;
    if (fail) throw new Error("unsubscribe unavailable");
  });
  const { manager } = createManager({
    getActiveAgent: () => agent,
    loadExtensionFromSource: (api) => {
      api.onAgentEvent(() => {});
      return loadExtension(api, () => [
        () => { successfulCleanups += 1; },
        () => { cleanups += 1; if (fail) throw new Error("cleanup unavailable"); },
      ]);
    },
  });
  await manager.initialize();
  const id = await manager.installFromModuleSpecifier("Lifecycle", "./lifecycle.js");
  await assert.rejects(manager.setExtensionEnabled(id, false), /cleanup unavailable/);
  assert.match(manager.list()[0]?.lastError ?? "", /cleanup unavailable/);
  fail = false;
  await manager.uninstallExtension(id);
  assert.equal(cleanups, 2, "uninstall must finish the teardown that disabling could not complete");
  assert.equal(unsubscribes, 2, "failed subscriptions must also be released before uninstall completes");
  assert.equal(successfulCleanups, 1, "successful cleanups must not be replayed");
  assert.deepEqual(manager.list(), []);
});

void test("an extension can unregister its resources from its teardown callback", async () => {
  const { manager, connectionManager } = createManager({ loadExtensionFromSource: (api) => {
    registerTool(api);
    const connection = api.connections.register({ id: "cleanup", title: "Cleanup", capability: "cleanup", authKind: "api_key", secretFields: [] });
    return Promise.resolve({ deactivate: () => {
      api.unregisterTool("lifecycle_tool");
      api.connections.unregister(connection);
      return Promise.resolve();
    } });
  } });
  await manager.initialize();
  const id = await manager.installFromModuleSpecifier("Lifecycle", "./lifecycle.js");
  await manager.setExtensionEnabled(id, false);
  assert.deepEqual(manager.getRegisteredTools(), []);
  assert.deepEqual(connectionManager.listRegisteredConnectionIds(), []);
});
