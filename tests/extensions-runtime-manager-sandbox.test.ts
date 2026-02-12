import assert from "node:assert/strict";
import { test } from "node:test";
import { readFile } from "node:fs/promises";

import { setExperimentalFeatureEnabled } from "../src/experiments/flags.ts";
import { ExtensionRuntimeManager } from "../src/extensions/runtime-manager.ts";
import { EXTENSIONS_REGISTRY_STORAGE_KEY } from "../src/extensions/store.ts";
import {
  getDefaultPermissionsForTrust,
  type StoredExtensionPermissions,
  type StoredExtensionTrust,
} from "../src/extensions/permissions.ts";

class MemorySettingsStore {
  private readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.has(key) ? this.values.get(key) ?? null : null);
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }

  writeRaw(key: string, value: unknown): void {
    this.values.set(key, value);
  }
}

class MemoryLocalStorage {
  private readonly values = new Map<string, string>();

  getItem(key: string): string | null {
    return this.values.has(key) ? this.values.get(key) ?? null : null;
  }

  setItem(key: string, value: string): void {
    this.values.set(key, value);
  }

  removeItem(key: string): void {
    this.values.delete(key);
  }
}

const EXTENSION_SANDBOX_RUNTIME_STORAGE_KEY = "pi.experimental.extensionSandboxRuntime";

function clearLocalStorageKey(key: string): void {
  const storage = Reflect.get(globalThis, "localStorage");
  if (typeof storage !== "object" || storage === null) {
    return;
  }

  const removeItem = Reflect.get(storage, "removeItem");
  if (typeof removeItem !== "function") {
    return;
  }

  Reflect.apply(removeItem, storage, [key]);
}

function installLocalStorageStub(): () => void {
  const previous = Reflect.get(globalThis, "localStorage");
  Reflect.set(globalThis, "localStorage", new MemoryLocalStorage());

  return () => {
    if (previous === undefined) {
      Reflect.deleteProperty(globalThis, "localStorage");
      return;
    }

    Reflect.set(globalThis, "localStorage", previous);
  };
}

function createStoredEntry(input: {
  id: string;
  name: string;
  trust: StoredExtensionTrust;
  enabled?: boolean;
  permissions?: StoredExtensionPermissions;
}): Record<string, unknown> {
  const now = new Date().toISOString();
  const source = input.trust === "inline-code"
    ? {
      kind: "inline" as const,
      code: "export function activate(api) { api.toast('hi'); }",
    }
    : {
      kind: "module" as const,
      specifier: "../extensions/snake.js",
    };

  return {
    id: input.id,
    name: input.name,
    enabled: input.enabled ?? true,
    source,
    trust: input.trust,
    permissions: input.permissions ?? getDefaultPermissionsForTrust(input.trust),
    createdAt: now,
    updatedAt: now,
  };
}

void test("runtime manager source wires sandbox runtime selection through rollback kill switch", async () => {
  const source = await readFile(new URL("../src/extensions/runtime-manager.ts", import.meta.url), "utf8");

  assert.match(source, /resolveExtensionRuntimeMode\(/);
  assert.match(source, /extension_sandbox_runtime/);
  assert.match(source, /extension_widget_v2/);
  assert.match(source, /runtimeMode/);
  assert.match(source, /runtimeLabel/);
  assert.match(source, /activateExtensionInSandbox/);
  assert.match(source, /activateInSandbox/);
  assert.match(source, /widgetOwnerId/);
  assert.match(source, /widgetApiV2Enabled/);
});

void test("extensions overlay source renders runtime label in installed rows", async () => {
  const source = await readFile(new URL("../src/commands/builtins/extensions-overlay.ts", import.meta.url), "utf8");

  assert.match(source, /Runtime: \$\{status\.runtimeLabel\}/);
});

void test("untrusted extensions default to sandbox runtime when rollback kill switch is unset", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    clearLocalStorageKey(EXTENSION_SANDBOX_RUNTIME_STORAGE_KEY);

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.inline.default",
          name: "Inline Default",
          trust: "inline-code",
        }),
      ],
    });

    let hostLoadCalls = 0;
    let sandboxLoadCalls = 0;

    const manager = new ExtensionRuntimeManager({
      settings,
      getActiveAgent: () => null,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
      loadExtensionFromSource: () => {
        hostLoadCalls += 1;
        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
      activateInSandbox: () => {
        sandboxLoadCalls += 1;
        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
    });

    await manager.initialize();

    const status = manager.list()[0];
    assert.equal(status.runtimeMode, "sandbox-iframe");
    assert.equal(status.loaded, true);
    assert.equal(hostLoadCalls, 0);
    assert.equal(sandboxLoadCalls, 1);
  } finally {
    restoreLocalStorage();
  }
});

void test("rollback kill switch routes untrusted extensions back to host runtime", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    setExperimentalFeatureEnabled("extension_sandbox_runtime", false);

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.inline.rollback",
          name: "Inline Rollback",
          trust: "inline-code",
        }),
      ],
    });

    let hostLoadCalls = 0;
    let sandboxLoadCalls = 0;

    const manager = new ExtensionRuntimeManager({
      settings,
      getActiveAgent: () => null,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
      loadExtensionFromSource: () => {
        hostLoadCalls += 1;
        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
      activateInSandbox: () => {
        sandboxLoadCalls += 1;
        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
    });

    await manager.initialize();

    const status = manager.list()[0];
    assert.equal(status.runtimeMode, "host");
    assert.equal(status.loaded, true);
    assert.equal(hostLoadCalls, 1);
    assert.equal(sandboxLoadCalls, 0);
  } finally {
    restoreLocalStorage();
  }
});

void test("trusted local-module extensions stay on host runtime even when sandbox default is active", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    clearLocalStorageKey(EXTENSION_SANDBOX_RUNTIME_STORAGE_KEY);

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.local",
          name: "Local Module",
          trust: "local-module",
        }),
      ],
    });

    let hostLoadCalls = 0;
    let sandboxLoadCalls = 0;

    const manager = new ExtensionRuntimeManager({
      settings,
      getActiveAgent: () => null,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
      loadExtensionFromSource: () => {
        hostLoadCalls += 1;
        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
      activateInSandbox: () => {
        sandboxLoadCalls += 1;
        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
    });

    await manager.initialize();

    const status = manager.list()[0];
    assert.equal(status.runtimeMode, "host");
    assert.equal(status.loaded, true);
    assert.equal(hostLoadCalls, 1);
    assert.equal(sandboxLoadCalls, 0);
  } finally {
    restoreLocalStorage();
  }
});

void test("sandbox runtime source enforces capability gates and rejects unknown ui actions", async () => {
  const source = await readFile(new URL("../src/extensions/sandbox-runtime.ts", import.meta.url), "utf8");

  assert.match(source, /case "register_tool": \{[\s\S]*this\.assertCapability\("tools\.register"\)/);
  assert.match(source, /normalizeSandboxToolParameters/);
  assert.match(source, /Type\.Unsafe/);
  assert.match(source, /case "overlay_show": \{[\s\S]*this\.assertCapability\("ui\.overlay"\)/);
  assert.match(source, /case "widget_show": \{[\s\S]*this\.assertCapability\("ui\.widget"\)/);
  assert.match(source, /case "widget_upsert": \{[\s\S]*Widget API v2 is disabled/);
  assert.match(source, /upsertSandboxWidgetNode\([\s\S]*element:\s*body/);
  assert.match(source, /case "widget_clear": \{/);
  assert.match(source, /if \(method === "ui_action"\)/);
  assert.match(source, /Unknown sandbox UI action id:/);
  assert.match(source, /allowWhenDisposed:\s*true/);

  // Isolation boundary checks: strict iframe sandboxing + host-side message source/direction checks.
  assert.match(source, /setAttribute\("sandbox", "allow-scripts"\)/);
  assert.match(source, /if \(event\.source !== this\.iframe\.contentWindow\)/);
  assert.match(source, /if \(envelope\.direction !== "sandbox_to_host"\)/);
  assert.match(source, /api\.agent is not available in sandbox runtime/);
});

void test("sandbox activation failures are isolated per extension during initialize", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    setExperimentalFeatureEnabled("extension_sandbox_runtime", true);
    setExperimentalFeatureEnabled("extension_permission_gates", true);

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.fail",
          name: "Failing Extension",
          trust: "inline-code",
        }),
        createStoredEntry({
          id: "ext.ok",
          name: "Healthy Extension",
          trust: "inline-code",
        }),
      ],
    });

    const calls: string[] = [];

    const manager = new ExtensionRuntimeManager({
      settings,
      getActiveAgent: () => null,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
      activateInSandbox: (activation) => {
        calls.push(activation.instanceId);

        if (activation.instanceId.startsWith("ext.fail.")) {
          return Promise.reject(new Error("sandbox boot failed"));
        }

        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
    });

    await manager.initialize();

    const statuses = manager.list();
    const failing = statuses.find((entry) => entry.id === "ext.fail");
    const healthy = statuses.find((entry) => entry.id === "ext.ok");

    assert.ok(failing);
    assert.ok(healthy);

    assert.equal(failing.loaded, false);
    assert.match(failing.lastError ?? "", /sandbox boot failed/);

    assert.equal(healthy.loaded, true);
    assert.equal(healthy.lastError, null);

    assert.equal(calls.length, 2);
  } finally {
    restoreLocalStorage();
  }
});

void test("sandbox capability denial surfaces deterministic permission error", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    setExperimentalFeatureEnabled("extension_sandbox_runtime", true);
    setExperimentalFeatureEnabled("extension_permission_gates", true);

    const basePermissions = getDefaultPermissionsForTrust("inline-code");
    const deniedOverlayPermissions: StoredExtensionPermissions = {
      ...basePermissions,
      uiOverlay: false,
    };

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.denied",
          name: "Denied Extension",
          trust: "inline-code",
          permissions: deniedOverlayPermissions,
        }),
      ],
    });

    const manager = new ExtensionRuntimeManager({
      settings,
      getActiveAgent: () => null,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
      activateInSandbox: (activation) => {
        if (!activation.isCapabilityEnabled("ui.overlay")) {
          return Promise.reject(new Error(activation.formatCapabilityError("ui.overlay")));
        }

        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
    });

    await manager.initialize();

    const status = manager.list()[0];
    assert.equal(status.loaded, false);
    assert.match(
      status.lastError ?? "",
      /Permission denied for extension "Denied Extension": cannot show overlays\./,
    );
  } finally {
    restoreLocalStorage();
  }
});
