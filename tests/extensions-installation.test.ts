import assert from "node:assert/strict";
import { test } from "node:test";

import { InMemoryModelsStore } from "@earendil-works/pi-ai";

import { ConnectionManager } from "../src/connections/manager.ts";
import {
  ExtensionRuntimeManager,
  type ExtensionRuntimeManagerOptions,
} from "../src/extensions/runtime-manager.ts";
import { setExperimentalFeatureEnabled } from "../src/experiments/flags.ts";
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

function createManager(
  options: { settings: MemorySettingsStore }
    & Pick<ExtensionRuntimeManagerOptions, "loadExtensionFromSource" | "activateInSandbox">,
): ExtensionRuntimeManager {
  const modelRuntime = new BrowserModelRuntime({
    providerKeys: new EmptyProviderKeys(),
    modelCatalogs: new InMemoryModelsStore(),
    getProxyUrl: () => Promise.resolve(undefined),
  });
  return new ExtensionRuntimeManager({
    settings: options.settings,
    connectionManager: new ConnectionManager({ settings: options.settings }),
    modelRuntime,
    getActiveAgent: () => null,
    refreshRuntimeTools: () => Promise.resolve(),
    refreshRuntimeModels: () => Promise.resolve(),
    reservedToolNames: new Set<string>(),
    ...(options.loadExtensionFromSource ? { loadExtensionFromSource: options.loadExtensionFromSource } : {}),
    ...(options.activateInSandbox ? { activateInSandbox: options.activateInSandbox } : {}),
  });
}

void test("the manager exposes the accepted source forms and their runtime modes", async () => {
  setExperimentalFeatureEnabled("extension_sandbox_runtime", false);
  const settings = new MemorySettingsStore();
  const manager = createManager({
    settings,
    loadExtensionFromSource: () => Promise.resolve({ deactivate: () => Promise.resolve() }),
    activateInSandbox: () => Promise.resolve({ deactivate: () => Promise.resolve() }),
  });
  await manager.initialize();

  await manager.installFromModuleSpecifier("Local", "./extensions/local.js");
  await manager.installFromModuleSpecifier("Blob", "blob:https://example.test/extension");
  await manager.installFromUrl("Remote", "https://example.test/extension.js");
  await manager.installFromCode("Inline", "export function activate() {}");

  assert.deepEqual(
    manager.list()
      .filter(({ id }) => !id.startsWith("builtin."))
      .map(({ name, trust, runtimeLabel, loaded }) => ({ name, trust, runtimeLabel, loaded })),
    [
      { name: "Local", trust: "local-module", runtimeLabel: "host runtime", loaded: true },
      { name: "Blob", trust: "inline-code", runtimeLabel: "sandbox iframe", loaded: true },
      { name: "Remote", trust: "remote-url", runtimeLabel: "sandbox iframe", loaded: true },
      { name: "Inline", trust: "inline-code", runtimeLabel: "sandbox iframe", loaded: true },
    ],
  );
});

void test("the manager reports refused and unavailable module sources as installation failures", async () => {
  setExperimentalFeatureEnabled("extension_sandbox_runtime", true);
  const settings = new MemorySettingsStore();
  const manager = createManager({ settings });
  await manager.initialize();

  await manager.installFromModuleSpecifier("Unsupported", "extension-package");
  await manager.installFromModuleSpecifier("Missing local", "./extensions/missing.js");

  assert.deepEqual(
    manager.list()
      .filter(({ id }) => !id.startsWith("builtin."))
      .map(({ name, loaded, lastError }) => ({ name, loaded, lastError })),
    [
      {
        name: "Unsupported",
        loaded: false,
        lastError: "Unsupported extension source \"extension-package\". Only local module specifiers (./, ../, /), blob: URLs, and inline function activators are allowed by default.",
      },
      {
        name: "Missing local",
        loaded: false,
        lastError: "Local extension module \"./extensions/missing.js\" was not bundled. Use a bundled module under src/extensions, paste code, or a remote URL (with explicit opt-in).",
      },
    ],
  );
});
