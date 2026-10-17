import assert from "node:assert/strict";
import { test } from "node:test";

import { InMemoryModelsStore } from "@earendil-works/pi-ai";
import { Type } from "typebox";

import { commandRegistry } from "../src/commands/types.ts";
import { ConnectionManager } from "../src/connections/manager.ts";
import { loadConnectionStoreDocument, saveConnectionStoreDocument } from "../src/connections/store.ts";
import { ExtensionRuntimeManager } from "../src/extensions/runtime-manager.ts";
import { EXTENSIONS_REGISTRY_STORAGE_KEY } from "../src/extensions/store.ts";
import { BrowserModelRuntime } from "../src/models/browser-model-runtime.ts";
import type { ProviderKeysStoreLike } from "../src/storage/local/provider-credentials-store.ts";

class MemorySettingsStore {
  private readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.get(key) ?? null);
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }
}

class EmptyProviderKeys implements ProviderKeysStoreLike {
  get(): Promise<string | null> { return Promise.resolve(null); }
  set(): Promise<void> { return Promise.resolve(); }
  delete(): Promise<void> { return Promise.resolve(); }
  list(): Promise<string[]> { return Promise.resolve([]); }
}

void test("installing and uninstalling an extension exposes and removes its capabilities", async () => {
  const settings = new MemorySettingsStore();
  await settings.set(EXTENSIONS_REGISTRY_STORAGE_KEY, { version: 2, items: [] });
  const connectionManager = new ConnectionManager({ settings });
  const modelRuntime = new BrowserModelRuntime({
    providerKeys: new EmptyProviderKeys(),
    modelCatalogs: new InMemoryModelsStore(),
    getProxyUrl: () => Promise.resolve(undefined),
  });
  const manager = new ExtensionRuntimeManager({
    settings,
    connectionManager,
    modelRuntime,
    getActiveAgent: () => null,
    refreshRuntimeTools: () => Promise.resolve(),
    refreshRuntimeModels: () => Promise.resolve(),
    reservedToolNames: new Set<string>(),
    loadExtensionFromSource: (api) => {
      api.registerCommand("audit_lookup", {
        description: "Look up an audit record",
        execute: () => {},
      });
      api.registerTool("audit_lookup", {
        description: "Look up an audit record",
        parameters: Type.Object({ id: Type.String() }),
        execute: () => ({
          content: [{ type: "text", text: "record found" }],
          details: undefined,
        }),
      });
      api.connections.register({
        id: "audit-service",
        title: "Audit service",
        capability: "record lookup",
        authKind: "none",
        secretFields: [],
      });
      api.models.registerProvider({
        id: "audit-models",
        name: "Audit models",
        api: "openai-responses",
        baseUrl: "http://127.0.0.1:11434/v1",
        models: [{ id: "audit-small", contextWindow: 8_192, maxTokens: 1_024 }],
        allowKeyless: true,
      });
      return Promise.resolve({ deactivate: () => Promise.resolve() });
    },
  });

  await manager.initialize();
  const extensionId = await manager.installFromModuleSpecifier("Audit helper", "./audit-helper.js");

  const [status] = manager.list();
  assert.equal(status?.id, extensionId);
  assert.equal(status?.loaded, true);
  assert.deepEqual(status?.toolNames, ["audit_lookup"]);
  assert.deepEqual(status?.commandNames, ["audit_lookup"]);
  assert.deepEqual(status?.modelProviderIds, [`${extensionId}.audit-models`]);
  assert.equal(manager.getRegisteredTools()[0]?.name, "audit_lookup");
  assert.equal(commandRegistry.get("audit_lookup")?.source, "extension");
  assert.equal(modelRuntime.models.getProvider(`${extensionId}.audit-models`)?.name, "Audit models");
  assert.deepEqual(connectionManager.listRegisteredConnectionIds(), [`${extensionId}.audit-service`]);

  await manager.uninstallExtension(extensionId);

  assert.deepEqual(manager.list(), []);
  assert.deepEqual(manager.getRegisteredTools(), []);
  assert.equal(commandRegistry.get("audit_lookup"), undefined);
  assert.equal(modelRuntime.models.getProvider(`${extensionId}.audit-models`), undefined);
  assert.deepEqual(connectionManager.listRegisteredConnectionIds(), []);
});

function createManagerWithConnection(settings: MemorySettingsStore, connectionManager: ConnectionManager): ExtensionRuntimeManager {
  return new ExtensionRuntimeManager({
    settings,
    connectionManager,
    modelRuntime: new BrowserModelRuntime({
      providerKeys: new EmptyProviderKeys(),
      modelCatalogs: new InMemoryModelsStore(),
      getProxyUrl: () => Promise.resolve(undefined),
    }),
    getActiveAgent: () => null,
    refreshRuntimeTools: () => Promise.resolve(),
    refreshRuntimeModels: () => Promise.resolve(),
    reservedToolNames: new Set<string>(),
    loadExtensionFromSource: (api) => {
      api.connections.register({
        id: "account",
        title: "Model account",
        capability: "model inference",
        authKind: "api_key",
        secretFields: [{ id: "apiKey", label: "API key", required: true }],
      });
      return Promise.resolve({ deactivate: () => Promise.resolve() });
    },
  });
}

void test("uninstalling an extension deletes its stored connection secrets", async () => {
  const settings = new MemorySettingsStore();
  await settings.set(EXTENSIONS_REGISTRY_STORAGE_KEY, { version: 2, items: [] });
  const connectionManager = new ConnectionManager({ settings });
  const manager = createManagerWithConnection(settings, connectionManager);
  await manager.initialize();

  const extensionId = await manager.installFromModuleSpecifier("Model account", "./model-account.js");
  await connectionManager.updateSecretsFromHost(`${extensionId}.account`, { apiKey: "user-secret" });
  await saveConnectionStoreDocument(settings, {
    ...(await loadConnectionStoreDocument(settings)),
    "builtin.mcp.servers": { status: "connected", secrets: { token: "mcp-token" } },
  });
  assert.deepEqual(Object.keys(await loadConnectionStoreDocument(settings)).sort(), [
    "builtin.mcp.servers",
    `${extensionId}.account`,
  ]);

  await manager.uninstallExtension(extensionId);

  assert.deepEqual(Object.keys(await loadConnectionStoreDocument(settings)), ["builtin.mcp.servers"]);
});

void test("startup deletes stored connections of extensions that are no longer installed", async () => {
  const settings = new MemorySettingsStore();
  await settings.set(EXTENSIONS_REGISTRY_STORAGE_KEY, { version: 2, items: [] });
  const connectionManager = new ConnectionManager({ settings });
  const firstPane = createManagerWithConnection(settings, connectionManager);
  await firstPane.initialize();
  const installedId = await firstPane.installFromModuleSpecifier("Model account", "./model-account.js");
  await connectionManager.updateSecretsFromHost(`${installedId}.account`, { apiKey: "installed-secret" });

  // Storage as an earlier version left it: the same registry, a secret left by
  // an old uninstall, and no record that the cleanup has run.
  const goneId = "ext.00000000-0000-4000-8000-000000000000";
  const upgraded = new MemorySettingsStore();
  await upgraded.set(EXTENSIONS_REGISTRY_STORAGE_KEY, await settings.get(EXTENSIONS_REGISTRY_STORAGE_KEY));
  await saveConnectionStoreDocument(upgraded, {
    ...(await loadConnectionStoreDocument(settings)),
    [`${goneId}.account`]: { status: "connected", secrets: { apiKey: "orphaned-secret" } },
    "builtin.mcp.servers": { status: "connected", secrets: { token: "mcp-token" } },
  });

  const reloadedPane = createManagerWithConnection(upgraded, new ConnectionManager({ settings: upgraded }));
  await reloadedPane.initialize();

  assert.deepEqual(Object.keys(await loadConnectionStoreDocument(upgraded)).sort(), [
    "builtin.mcp.servers",
    `${installedId}.account`,
  ]);
});

for (const { name, registry } of [
  { name: "missing", registry: null },
  { name: "malformed", registry: { version: 2, items: "unreadable" } },
  { name: "future-version", registry: { version: 3, items: [] } },
  { name: "partly unreadable", registry: {
    version: 2,
    items: [{ id: "ext.kept", name: "Newer extension", source: { kind: "future-kind" } }],
  } },
]) {
  void test(`startup preserves credentials across reloads with a ${name} registry`, async () => {
    const settings = new MemorySettingsStore();
    await settings.set(EXTENSIONS_REGISTRY_STORAGE_KEY, registry);
    const connections = {
      "ext.kept.account": { status: "connected" as const, secrets: { apiKey: "installed-secret" } },
    };
    await saveConnectionStoreDocument(settings, connections);

    for (let launch = 0; launch < 2; launch += 1) {
      await createManagerWithConnection(settings, new ConnectionManager({ settings })).initialize();
      assert.deepEqual(await loadConnectionStoreDocument(settings), connections);
    }
  });
}

void test("failed migration persistence cannot replace a missing registry with defaults", async () => {
  class FailingWrites extends MemorySettingsStore {
    failWrites = false;

    override set(key: string, value: unknown): Promise<void> {
      return this.failWrites
        ? Promise.reject(new Error("storage unavailable"))
        : super.set(key, value);
    }
  }
  const settings = new FailingWrites();
  const connections = {
    "ext.kept.account": { status: "connected" as const, secrets: { apiKey: "installed-secret" } },
  };
  await saveConnectionStoreDocument(settings, connections);
  settings.failWrites = true;
  await assert.rejects(
    createManagerWithConnection(settings, new ConnectionManager({ settings })).initialize(),
    /storage unavailable/u,
  );
  assert.equal(await settings.get(EXTENSIONS_REGISTRY_STORAGE_KEY), null);

  settings.failWrites = false;
  for (let launch = 0; launch < 2; launch += 1) {
    await createManagerWithConnection(settings, new ConnectionManager({ settings })).initialize();
    assert.deepEqual(await loadConnectionStoreDocument(settings), connections);
  }
});

void test("startup cleanup of uninstalled extensions' connections runs only once", async () => {
  const settings = new MemorySettingsStore();
  await settings.set(EXTENSIONS_REGISTRY_STORAGE_KEY, { version: 2, items: [] });
  const firstOrphan = "ext.00000000-0000-4000-8000-000000000003.account";
  await saveConnectionStoreDocument(settings, {
    [firstOrphan]: { status: "connected", secrets: { apiKey: "orphaned-secret" } },
  });
  await createManagerWithConnection(settings, new ConnectionManager({ settings })).initialize();
  assert.deepEqual(await loadConnectionStoreDocument(settings), {});

  const laterRecord = "ext.00000000-0000-4000-8000-000000000004.account";
  await saveConnectionStoreDocument(settings, {
    [laterRecord]: { status: "connected", secrets: { apiKey: "later-secret" } },
  });
  await createManagerWithConnection(settings, new ConnectionManager({ settings })).initialize();

  assert.deepEqual(Object.keys(await loadConnectionStoreDocument(settings)), [laterRecord]);
});
