import assert from "node:assert/strict";
import { test } from "node:test";
import { readFile } from "node:fs/promises";

import { Type } from "@sinclair/typebox";

import { ConnectionManager } from "../src/connections/manager.ts";
import { setExperimentalFeatureEnabled } from "../src/experiments/flags.ts";
import { ExtensionRuntimeManager } from "../src/extensions/runtime-manager.ts";
import { isConnectionToolErrorDetails } from "../src/tools/tool-details.ts";
import { withConnectionPreflight } from "../src/tools/with-connection-preflight.ts";
import {
  SANDBOX_CHANNEL,
  isSandboxEnvelope,
  serializeForSandboxInlineScript,
} from "../src/extensions/sandbox/protocol.ts";
import { buildSandboxSrcdoc } from "../src/extensions/sandbox/srcdoc.ts";
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

function createConnectionManager(settings: MemorySettingsStore): ConnectionManager {
  return new ConnectionManager({ settings });
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

const EXTENSION_SANDBOX_RUNTIME_STORAGE_KEY = "pi.experimental.extensionSandboxHostFallback";

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

void test("runtime manager passes sandbox activation options and runtime metadata", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    clearLocalStorageKey(EXTENSION_SANDBOX_RUNTIME_STORAGE_KEY);
    setExperimentalFeatureEnabled("extension_sandbox_runtime", false);
    setExperimentalFeatureEnabled("extension_permission_gates", true);
    setExperimentalFeatureEnabled("extension_widget_v2", true);

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.inline.options",
          name: "Inline Options",
          trust: "inline-code",
        }),
      ],
    });

    let capturedWidgetOwnerId = "";
    let capturedWidgetApiV2Enabled = false;
    let capturedSourceKind = "";
    let capturedSourceHasToast = false;
    let overlayCapabilityAllowed = false;
    let toolsCapabilityAllowed = true;
    let capabilityErrorText = "";

    const manager = new ExtensionRuntimeManager({
      settings,
      connectionManager: createConnectionManager(settings),
      getActiveAgent: () => null,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
      activateInSandbox: (activation) => {
        capturedWidgetOwnerId = activation.widgetOwnerId ?? "";
        capturedWidgetApiV2Enabled = activation.widgetApiV2Enabled === true;
        capturedSourceKind = activation.source.kind;
        capturedSourceHasToast = activation.source.kind === "inline"
          && activation.source.code.includes("api.toast('hi')");
        overlayCapabilityAllowed = activation.isCapabilityEnabled("ui.overlay");
        toolsCapabilityAllowed = activation.isCapabilityEnabled("tools.register");
        capabilityErrorText = activation.formatCapabilityError("ui.overlay");

        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
    });

    await manager.initialize();

    const status = manager.list()[0];
    assert.equal(status.runtimeMode, "sandbox-iframe");
    assert.equal(status.runtimeLabel, "sandbox iframe");
    assert.equal(capturedWidgetOwnerId, "ext.inline.options");
    assert.equal(capturedWidgetApiV2Enabled, true);
    assert.equal(capturedSourceKind, "inline");
    assert.equal(capturedSourceHasToast, true);
    assert.equal(overlayCapabilityAllowed, true);
    assert.equal(toolsCapabilityAllowed, false);
    assert.match(capabilityErrorText, /cannot show overlays\./);
  } finally {
    restoreLocalStorage();
  }
});

void test("extensions hub plugins source includes runtime metadata for installed rows", async () => {
  const source = await readFile(new URL("../src/commands/builtins/extensions-hub-plugins.ts", import.meta.url), "utf8");

  assert.match(source, /description:\s*`\$\{status\.sourceLabel\} · \$\{status\.runtimeLabel\}`/);
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
      connectionManager: createConnectionManager(settings),
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
    setExperimentalFeatureEnabled("extension_sandbox_runtime", true);

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
      connectionManager: createConnectionManager(settings),
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
      connectionManager: createConnectionManager(settings),
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

void test("host runtime extension tools include source provenance in descriptions", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    clearLocalStorageKey(EXTENSION_SANDBOX_RUNTIME_STORAGE_KEY);

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.apollo",
          name: "Apollo Helper",
          trust: "local-module",
        }),
      ],
    });

    const manager = new ExtensionRuntimeManager({
      settings,
      connectionManager: createConnectionManager(settings),
      getActiveAgent: () => null,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
      loadExtensionFromSource: (api) => {
        api.registerTool("apollo_enrich", {
          description: "Enrich contacts with Apollo API",
          parameters: Type.Object({
            linkedin_url: Type.String(),
          }),
          execute: () => ({
            content: [{ type: "text", text: "ok" }],
            details: undefined,
          }),
        });

        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
      activateInSandbox: () => {
        throw new Error("sandbox runtime should not be used for local-module extensions");
      },
    });

    await manager.initialize();

    const tools = manager.getRegisteredTools();
    assert.equal(tools.length, 1);

    const description = tools[0]?.description ?? "";
    assert.match(description, /Enrich contacts with Apollo API/);
    assert.match(description, /Source: extension "Apollo Helper" \(ext\.apollo\)\./);
  } finally {
    restoreLocalStorage();
  }
});

void test("host runtime extension tools tolerate missing descriptions", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    clearLocalStorageKey(EXTENSION_SANDBOX_RUNTIME_STORAGE_KEY);

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.apollo.missing-description",
          name: "Apollo Helper",
          trust: "local-module",
        }),
      ],
    });

    const manager = new ExtensionRuntimeManager({
      settings,
      connectionManager: createConnectionManager(settings),
      getActiveAgent: () => null,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
      loadExtensionFromSource: (api) => {
        const toolWithoutDescription = {
          parameters: Type.Object({
            linkedin_url: Type.String(),
          }),
          execute: () => ({
            content: [{ type: "text", text: "ok" }],
            details: undefined,
          }),
        };

        Reflect.apply(api.registerTool, api, ["apollo_enrich", toolWithoutDescription]);

        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
      activateInSandbox: () => {
        throw new Error("sandbox runtime should not be used for local-module extensions");
      },
    });

    await manager.initialize();

    const tools = manager.getRegisteredTools();
    assert.equal(tools.length, 1);

    const description = tools[0]?.description ?? "";
    assert.equal(description, "Source: extension \"Apollo Helper\" (ext.apollo.missing-description).");
  } finally {
    restoreLocalStorage();
  }
});

void test("host runtime extension tool errors include extension ownership context", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    clearLocalStorageKey(EXTENSION_SANDBOX_RUNTIME_STORAGE_KEY);

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.apollo.error",
          name: "Apollo Helper",
          trust: "local-module",
        }),
      ],
    });

    const manager = new ExtensionRuntimeManager({
      settings,
      connectionManager: createConnectionManager(settings),
      getActiveAgent: () => null,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
      loadExtensionFromSource: (api) => {
        api.registerTool("apollo_enrich", {
          description: "Enrich contacts with Apollo API",
          parameters: Type.Object({
            linkedin_url: Type.String(),
          }),
          execute: () => {
            throw new Error("tool.execute is not a function");
          },
        });

        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
      activateInSandbox: () => {
        throw new Error("sandbox runtime should not be used for local-module extensions");
      },
    });

    await manager.initialize();

    const tool = manager.getRegisteredTools()[0];
    assert.ok(tool);
    if (!tool) {
      return;
    }

    await assert.rejects(
      async () => tool.execute("call-1", { linkedin_url: "https://linkedin.com/in/alice" }),
      /\[Extension Apollo Helper\] Tool "apollo_enrich" failed: tool\.execute is not a function/,
    );
  } finally {
    restoreLocalStorage();
  }
});

void test("sandbox srcdoc builder emits expected bridge hooks and config", () => {
  const html = buildSandboxSrcdoc({
    instanceId: "ext.inline.srcdoc",
    extensionName: "Inline Srcdoc",
    source: {
      kind: "inline",
      code: "export function activate(api) { api.toast('hello'); }",
    },
    widgetApiV2Enabled: true,
  });

  assert.match(html, /"instanceId":"ext\.inline\.srcdoc"/);
  assert.match(html, /"widgetApiV2Enabled":true/);
  assert.match(html, /if \(method === "ui_action"\)/);
  assert.match(html, /Unknown sandbox UI action id:/);
  assert.match(html, /api\.agent is not available in sandbox runtime/);
  assert.match(html, /placement: payload\.placement === "above-input" \|\| payload\.placement === "below-input"/);
  assert.match(html, /payload\.minHeightPx === null/);
});

void test("sandbox protocol helpers validate envelope shapes and escape inline script payloads", () => {
  const validRequest: unknown = {
    channel: SANDBOX_CHANNEL,
    instanceId: "ext.inline.proto",
    direction: "sandbox_to_host",
    kind: "request",
    requestId: "req-1",
    method: "register_tool",
  };

  const invalidDirection: unknown = {
    channel: SANDBOX_CHANNEL,
    instanceId: "ext.inline.proto",
    direction: "sideways",
    kind: "request",
    requestId: "req-1",
    method: "register_tool",
  };

  const invalidKind: unknown = {
    channel: SANDBOX_CHANNEL,
    instanceId: "ext.inline.proto",
    direction: "sandbox_to_host",
    kind: "invalid",
    requestId: "req-1",
    method: "register_tool",
  };

  assert.equal(isSandboxEnvelope(validRequest), true);
  assert.equal(isSandboxEnvelope(invalidDirection), false);
  assert.equal(isSandboxEnvelope(invalidKind), false);

  const serialized = serializeForSandboxInlineScript({ html: "<div>safe</div>" });
  assert.equal(serialized.includes("\\u003cdiv>safe\\u003c/div>"), true);
});

void test("sandbox runtime host source retains isolation boundary guards", async () => {
  const hostSource = await readFile(new URL("../src/extensions/sandbox-runtime.ts", import.meta.url), "utf8");

  assert.match(hostSource, /setAttribute\("sandbox", "allow-scripts"\)/);
  assert.match(hostSource, /if \(event\.source !== this\.iframe\.contentWindow\)/);
  assert.match(hostSource, /if \(envelope\.direction !== "sandbox_to_host"\)/);
  assert.match(hostSource, /if \(!isSandboxEnvelope\(envelope\)\)/);
});

void test("sandbox activation failures are isolated per extension during initialize", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    setExperimentalFeatureEnabled("extension_sandbox_runtime", false);
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
      connectionManager: createConnectionManager(settings),
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
    setExperimentalFeatureEnabled("extension_sandbox_runtime", false);
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
      connectionManager: createConnectionManager(settings),
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

void test("host runtime extensions can register required connections after tool registration during activation", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    clearLocalStorageKey(EXTENSION_SANDBOX_RUNTIME_STORAGE_KEY);

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.apollo.req",
          name: "Apollo Requires Connection",
          trust: "local-module",
        }),
      ],
    });

    const manager = new ExtensionRuntimeManager({
      settings,
      connectionManager: createConnectionManager(settings),
      getActiveAgent: () => null,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
      loadExtensionFromSource: (api) => {
        api.registerTool("apollo_lookup", {
          description: "Lookup company via Apollo",
          parameters: Type.Object({
            company: Type.String(),
          }),
          requiresConnection: "apollo",
          execute: () => ({
            content: [{ type: "text", text: "ok" }],
            details: undefined,
          }),
        });

        api.connections.register({
          id: "apollo",
          title: "Apollo",
          capability: "company and contact enrichment",
          authKind: "api_key",
          secretFields: [{ id: "apiKey", label: "API key", required: true }],
        });

        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
      activateInSandbox: () => {
        throw new Error("sandbox runtime should not be used for local-module extensions");
      },
    });

    await manager.initialize();

    const status = manager.list()[0];
    assert.equal(status.loaded, true);
    assert.equal(status.lastError, null);

    const tool = manager.getRegisteredTools()[0];
    assert.ok(tool);
    assert.deepEqual(Reflect.get(tool, "requiresConnection"), ["ext.apollo.req.apollo"]);
  } finally {
    restoreLocalStorage();
  }
});

void test("host runtime extensions fail activation when required connections were never registered", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    clearLocalStorageKey(EXTENSION_SANDBOX_RUNTIME_STORAGE_KEY);

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.apollo.req.missing",
          name: "Apollo Missing Connection",
          trust: "local-module",
        }),
      ],
    });

    const manager = new ExtensionRuntimeManager({
      settings,
      connectionManager: createConnectionManager(settings),
      getActiveAgent: () => null,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
      loadExtensionFromSource: (api) => {
        api.registerTool("apollo_lookup", {
          description: "Lookup company via Apollo",
          parameters: Type.Object({
            company: Type.String(),
          }),
          requiresConnection: "apollo",
          execute: () => ({
            content: [{ type: "text", text: "ok" }],
            details: undefined,
          }),
        });

        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
      activateInSandbox: () => {
        throw new Error("sandbox runtime should not be used for local-module extensions");
      },
    });

    await manager.initialize();

    const status = manager.list()[0];
    assert.equal(status.loaded, false);
    assert.match(status.lastError ?? "", /requires an invalid connection/);
  } finally {
    restoreLocalStorage();
  }
});

void test("connection preflight blocks missing connections before tool execution", async () => {
  const settings = new MemorySettingsStore();
  const connectionManager = createConnectionManager(settings);

  connectionManager.registerDefinition("ext.apollo", {
    id: "ext.apollo.apollo",
    title: "Apollo",
    capability: "company and contact enrichment",
    authKind: "api_key",
    secretFields: [{ id: "apiKey", label: "API key", required: true }],
  });

  let executed = false;

  const tool = {
    name: "apollo_lookup",
    label: "Apollo Lookup",
    description: "Lookup company profiles",
    parameters: Type.Object({
      company: Type.String(),
    }),
    execute: () => {
      executed = true;
      return {
        content: [{ type: "text", text: "should not execute" }],
        details: undefined,
      };
    },
  };

  Reflect.set(tool, "requiresConnection", ["ext.apollo.apollo"]);

  const tools = withConnectionPreflight([tool], {
    connectionManager,
  });

  const result = await tools[0].execute("tool-call-1", { company: "Acme" });

  assert.equal(executed, false);
  assert.equal(isConnectionToolErrorDetails(result.details), true);
  if (isConnectionToolErrorDetails(result.details)) {
    assert.equal(result.details.errorCode, "missing_connection");
    assert.equal(result.details.connectionId, "ext.apollo.apollo");
  }
});

void test("connection preflight maps runtime auth failures to connection_auth_failed", async () => {
  const settings = new MemorySettingsStore();
  const connectionManager = createConnectionManager(settings);

  connectionManager.registerDefinition("ext.apollo", {
    id: "ext.apollo.apollo",
    title: "Apollo",
    capability: "company and contact enrichment",
    authKind: "api_key",
    secretFields: [{ id: "apiKey", label: "API key", required: true }],
  });

  await connectionManager.setSecrets("ext.apollo", "ext.apollo.apollo", {
    apiKey: "top-secret-api-key",
  });

  const tool = {
    name: "apollo_lookup",
    label: "Apollo Lookup",
    description: "Lookup company profiles",
    parameters: Type.Object({
      company: Type.String(),
    }),
    execute: () => {
      throw new Error("401 Unauthorized: invalid API key top-secret-api-key");
    },
  };

  Reflect.set(tool, "requiresConnection", ["ext.apollo.apollo"]);

  const tools = withConnectionPreflight([tool], {
    connectionManager,
  });

  const result = await tools[0].execute("tool-call-2", { company: "Acme" });
  assert.equal(isConnectionToolErrorDetails(result.details), true);

  if (isConnectionToolErrorDetails(result.details)) {
    assert.equal(result.details.errorCode, "connection_auth_failed");
    assert.ok(!result.details.reason?.includes("top-secret-api-key"));
  }

  const state = await connectionManager.getState("ext.apollo.apollo");
  assert.equal(state?.status, "error");
  assert.ok(!state?.lastError?.includes("top-secret-api-key"));
});

void test("connection preflight attributes auth failures to the matching required connection", async () => {
  const settings = new MemorySettingsStore();
  const connectionManager = createConnectionManager(settings);

  connectionManager.registerDefinition("ext.multi", {
    id: "ext.multi.crm",
    title: "CRM",
    capability: "contact sync",
    authKind: "api_key",
    secretFields: [{ id: "apiKey", label: "API key", required: true }],
  });

  connectionManager.registerDefinition("ext.multi", {
    id: "ext.multi.billing",
    title: "Billing",
    capability: "invoice sync",
    authKind: "api_key",
    secretFields: [{ id: "apiKey", label: "API key", required: true }],
  });

  await connectionManager.setSecrets("ext.multi", "ext.multi.crm", {
    apiKey: "crm-key",
  });
  await connectionManager.setSecrets("ext.multi", "ext.multi.billing", {
    apiKey: "billing-key",
  });

  const tool = {
    name: "sync_everything",
    label: "Sync everything",
    description: "Sync CRM and billing systems",
    parameters: Type.Object({
      dryRun: Type.Boolean(),
    }),
    execute: () => {
      throw new Error("403 Forbidden for ext.multi.billing: invalid token billing-key");
    },
  };

  Reflect.set(tool, "requiresConnection", ["ext.multi.crm", "ext.multi.billing"]);

  const tools = withConnectionPreflight([tool], {
    connectionManager,
  });

  const result = await tools[0].execute("tool-call-3", { dryRun: false });
  assert.equal(isConnectionToolErrorDetails(result.details), true);

  if (isConnectionToolErrorDetails(result.details)) {
    assert.equal(result.details.errorCode, "connection_auth_failed");
    assert.equal(result.details.connectionId, "ext.multi.billing");
    assert.ok(!result.details.reason?.includes("billing-key"));
  }

  const crmState = await connectionManager.getState("ext.multi.crm");
  const billingState = await connectionManager.getState("ext.multi.billing");

  assert.equal(crmState?.status, "connected");
  assert.equal(billingState?.status, "error");
});
