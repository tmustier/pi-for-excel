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

void test("the manager exposes the accepted source forms and their runtime modes", async (t) => {
  setExperimentalFeatureEnabled("extension_sandbox_runtime", false);
  // Remote URLs need the explicit opt-in, which lives in localStorage.
  const values = new Map<string, string>();
  Reflect.set(globalThis, "localStorage", {
    getItem: (key: string) => values.get(key) ?? null,
    setItem: (key: string, value: string) => { values.set(key, value); },
    removeItem: (key: string) => { values.delete(key); },
  });
  t.after(() => { Reflect.deleteProperty(globalThis, "localStorage"); });
  setExperimentalFeatureEnabled("remote_extension_urls", true);
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

void test("pasted code gets the untrusted defaults plus only the capabilities granted at install", async () => {
  setExperimentalFeatureEnabled("extension_sandbox_runtime", false);
  const settings = new MemorySettingsStore();
  const manager = createManager({
    settings,
    activateInSandbox: () => Promise.resolve({ deactivate: () => Promise.resolve() }),
  });
  await manager.initialize();

  const plainId = await manager.installFromCode("Plain", "export function activate() {}");
  const toolId = await manager.installFromCode("Tooling", "export function activate() {}", ["tools.register"]);
  const capabilitiesOf = (id: string) => manager.list().find((status) => status.id === id)?.effectiveCapabilities ?? [];

  assert.equal(capabilitiesOf(plainId).includes("tools.register"), false);
  assert.equal(capabilitiesOf(plainId).includes("http.fetch"), false);
  assert.deepEqual(
    capabilitiesOf(toolId).filter((capability) => !capabilitiesOf(plainId).includes(capability)),
    ["tools.register"],
  );
});
