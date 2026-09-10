import assert from "node:assert/strict";
import { test } from "node:test";

import { InMemoryModelsStore } from "@earendil-works/pi-ai";
import { Type } from "typebox";

import { commandRegistry } from "../src/commands/types.ts";
import { ConnectionManager } from "../src/connections/manager.ts";
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
