import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import {
  InMemoryModelsStore,
  type ImageContent,
  type TextContent,
} from "@earendil-works/pi-ai";
import { Agent } from "@earendil-works/pi-agent-core";

import { ConnectionManager } from "../src/connections/manager.ts";
import { CONNECTION_STORE_KEY } from "../src/connections/store.ts";
import { setExperimentalFeatureEnabled } from "../src/experiments/flags.ts";
import { ExtensionRuntimeManager } from "../src/extensions/runtime-manager.ts";
import {
  activateExtensionInSandbox,
  type SandboxActivationOptions,
} from "../src/extensions/sandbox-runtime.ts";
import { BrowserModelRuntime } from "../src/models/browser-model-runtime.ts";
import type { ProviderKeysStoreLike } from "../src/storage/local/provider-credentials-store.ts";
import { failOnUnexpectedStream } from "./fail-on-unexpected-stream.ts";
import { decodeToolDetailsOfKind } from "../src/tools/tool-details.ts";
import { withConnectionPreflight } from "../src/tools/with-connection-preflight.ts";
import {
  SANDBOX_BOOTSTRAP_KIND,
  SANDBOX_CHANNEL,
  isSandboxBootstrapEnvelope,
  isSandboxEnvelope,
  serializeForSandboxInlineScript,
} from "../src/extensions/sandbox/protocol.ts";
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

class TransientExtensionReadSettings extends MemorySettingsStore {
  private failNextRead = false;

  armReadFailure(): void {
    this.failNextRead = true;
  }

  override get(key: string): Promise<unknown> {
    if (this.failNextRead) {
      this.failNextRead = false;
      return Promise.reject(new Error("transient runtime registry read failure"));
    }
    return super.get(key);
  }
}

class FailingConnectionStoreSettings extends MemorySettingsStore {
  private failNextConnectionStoreWrite = false;

  armConnectionStoreFailure(): void {
    this.failNextConnectionStoreWrite = true;
  }

  override set(key: string, value: unknown): Promise<void> {
    if (this.failNextConnectionStoreWrite && key === CONNECTION_STORE_KEY) {
      this.failNextConnectionStoreWrite = false;
      return Promise.reject(new Error("simulated connection store failure"));
    }

    return super.set(key, value);
  }
}

function createConnectionManager(settings: MemorySettingsStore): ConnectionManager {
  return new ConnectionManager({ settings });
}

class EmptyProviderKeys implements ProviderKeysStoreLike {
  get(): Promise<string | null> {
    return Promise.resolve(null);
  }

  set(): Promise<void> {
    return Promise.resolve();
  }

  delete(): Promise<void> {
    return Promise.resolve();
  }

  list(): Promise<string[]> {
    return Promise.resolve([]);
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

void test("runtime manager retries initialization after a transient registry read failure", async () => {
  const settings = new TransientExtensionReadSettings();
  settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
    version: 2,
    items: [createStoredEntry({
      id: "ext.persisted.disabled",
      name: "Persisted disabled",
      trust: "inline-code",
      enabled: false,
    })],
  });
  const manager = new ExtensionRuntimeManager({
    settings,
    connectionManager: createConnectionManager(settings),
    getActiveAgent: () => null,
    refreshRuntimeTools: () => Promise.resolve(),
    refreshRuntimeModels: () => Promise.resolve(),
    reservedToolNames: new Set<string>(),
  });
  settings.armReadFailure();

  await assert.rejects(() => manager.initialize(), /transient runtime registry read failure/u);
  assert.deepEqual(manager.list(), []);

  await manager.initialize();
  assert.equal(manager.list()[0]?.id, "ext.persisted.disabled");
});

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

void test("host and sandbox bridges retain distinct authority and owner boundaries", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  try {
    clearLocalStorageKey(EXTENSION_SANDBOX_RUNTIME_STORAGE_KEY);

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.local.authority",
          name: "Local Authority",
          trust: "local-module",
        }),
        createStoredEntry({
          id: "ext.inline.authority",
          name: "Sandbox Authority",
          trust: "inline-code",
        }),
      ],
    });

    const agent = new Agent({
      streamFn: failOnUnexpectedStream,
      initialState: {
        messages: [],
        tools: [],
      },
    });
    let hostRawAgent: Agent | null = null;
    let hostConnectionId = "";
    let sandboxConnectionId = "";
    let sandboxHasRawAgentCallback = true;

    const manager = new ExtensionRuntimeManager({
      settings,
      connectionManager: createConnectionManager(settings),
      getActiveAgent: () => agent,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
      loadExtensionFromSource: (api) => {
        hostRawAgent = api.agent.raw;
        hostConnectionId = api.connections.register({
          id: "ext.other.host-secret",
          title: "Host connection",
          capability: "host test",
          authKind: "api_key",
          secretFields: [],
        });
        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
      activateInSandbox: (activation) => {
        sandboxHasRawAgentCallback = Reflect.has(activation, "getAgent");
        sandboxConnectionId = activation.registerConnection({
          id: "ext.other.sandbox-secret",
          title: "Sandbox connection",
          capability: "sandbox test",
          authKind: "api_key",
          secretFields: [],
        });
        return Promise.resolve({
          deactivate: () => Promise.resolve(),
        });
      },
    });

    await manager.initialize();

    assert.equal(hostRawAgent, agent);
    assert.equal(hostConnectionId, "ext.local.authority.ext.other.host-secret");
    assert.equal(sandboxHasRawAgentCallback, false);
    assert.equal(sandboxConnectionId, "ext.inline.authority.ext.other.sandbox-secret");
    assert.deepEqual(manager.list().map((status) => status.loaded), [true, true]);
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

void test("extension tool revision increments on schema-stable reloads", async () => {
  const restoreLocalStorage = installLocalStorageStub();

  const readFirstText = (parts: ReadonlyArray<TextContent | ImageContent>): string | null => {
    for (const part of parts) {
      if (part.type === "text") {
        return part.text;
      }
    }

    return null;
  };

  try {
    clearLocalStorageKey(EXTENSION_SANDBOX_RUNTIME_STORAGE_KEY);

    const settings = new MemorySettingsStore();
    settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
      version: 2,
      items: [
        createStoredEntry({
          id: "ext.apollo.revision",
          name: "Apollo Revision",
          trust: "local-module",
        }),
      ],
    });

    let executeVersion = 0;
    const refreshRevisionSnapshots: number[] = [];

    let manager: ExtensionRuntimeManager | null = null;
    manager = new ExtensionRuntimeManager({
      settings,
      connectionManager: createConnectionManager(settings),
      getActiveAgent: () => null,
      refreshRuntimeTools: () => {
        if (!manager) {
          throw new Error("manager not initialized");
        }
        refreshRevisionSnapshots.push(manager.getToolRevision());
        return Promise.resolve();
      },
      reservedToolNames: new Set<string>(),
      loadExtensionFromSource: (api) => {
        executeVersion += 1;
        const currentVersion = executeVersion;

        api.registerTool("apollo_enrich", {
          description: "Enrich contacts with Apollo API",
          parameters: Type.Object({
            linkedin_url: Type.String(),
          }),
          execute: () => ({
            content: [{ type: "text", text: `v${currentVersion}` }],
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

    assert.equal(manager.getToolRevision(), 0);

    await manager.initialize();

    const revisionAfterInitialize = manager.getToolRevision();
    assert.equal(revisionAfterInitialize > 0, true);
    assert.equal(refreshRevisionSnapshots.includes(revisionAfterInitialize), true);

    const initialTool = manager.getRegisteredTools()[0];
    assert.ok(initialTool);
    if (!initialTool) {
      return;
    }

    const initialResult = await initialTool.execute("call-1", { linkedin_url: "https://linkedin.com/in/alice" });
    const initialText = readFirstText(initialResult.content);
    assert.equal(initialText, "v1");

    await manager.reloadExtension("ext.apollo.revision");

    const revisionAfterReload = manager.getToolRevision();
    assert.equal(revisionAfterReload > revisionAfterInitialize, true);
    assert.equal(refreshRevisionSnapshots.includes(revisionAfterReload), true);

    for (let i = 1; i < refreshRevisionSnapshots.length; i += 1) {
      const previous = refreshRevisionSnapshots[i - 1];
      const current = refreshRevisionSnapshots[i];
      assert.equal(current >= previous, true);
    }

    const reloadedTool = manager.getRegisteredTools()[0];
    assert.ok(reloadedTool);
    if (!reloadedTool) {
      return;
    }

    const reloadedResult = await reloadedTool.execute("call-2", { linkedin_url: "https://linkedin.com/in/alice" });
    const reloadedText = readFirstText(reloadedResult.content);
    assert.equal(reloadedText, "v2");
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

void test("sandbox protocol helpers validate envelope shapes and escape inline script payloads", () => {
  const validBootstrap: unknown = {
    channel: SANDBOX_CHANNEL,
    instanceId: "ext.inline.proto",
    direction: "host_to_sandbox",
    kind: SANDBOX_BOOTSTRAP_KIND,
  };

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

  assert.equal(isSandboxBootstrapEnvelope(validBootstrap), true);
  assert.equal(isSandboxBootstrapEnvelope(validRequest), false);
  assert.equal(isSandboxEnvelope(validRequest), true);
  assert.equal(isSandboxEnvelope(invalidDirection), false);
  assert.equal(isSandboxEnvelope(invalidKind), false);

  const serialized = serializeForSandboxInlineScript({ html: "<div>safe</div>" });
  assert.equal(serialized.includes("\\u003cdiv>safe\\u003c/div>"), true);
});

void test("sandbox activation uses a script-only iframe and dedicated, direction-checked message port", async () => {
  const previousDocument = Reflect.get(globalThis, "document");
  const previousWindow = Reflect.get(globalThis, "window");
  const previousHTMLElement = Reflect.get(globalThis, "HTMLElement");
  let iframe: HTMLElement | null = null;
  let sandboxPort: MessagePort | null = null;
  let activationSettled = false;
  const getIframe = (): HTMLElement | null => iframe;
  const getSandboxPort = (): MessagePort | null => sandboxPort;

  const body = {
    appendChild: (element: HTMLElement): HTMLElement => {
      iframe = element;
      return element;
    },
  };
  const fakeDocument = {
    body,
    createElement: (tagName: string): HTMLElement => {
      assert.equal(tagName, "iframe");
      const target = new EventTarget();
      const attributes = new Map<string, string>();
      const element = target as unknown as HTMLElement;
      Reflect.set(element, "style", {});
      Reflect.set(element, "setAttribute", (name: string, value: string) => attributes.set(name, value));
      Reflect.set(element, "getAttribute", (name: string) => attributes.get(name) ?? null);
      Reflect.set(element, "remove", () => {});
      Reflect.set(element, "contentWindow", {
        postMessage: (bootstrap: unknown, targetOrigin: string, transfer: MessagePort[]) => {
          assert.equal(targetOrigin, "*");
          assert.equal(transfer.length, 1);
          sandboxPort = transfer[0] ?? null;
          assert.deepEqual(bootstrap, {
            channel: SANDBOX_CHANNEL,
            instanceId: "ext.sandbox.boundary",
            direction: "host_to_sandbox",
            kind: SANDBOX_BOOTSTRAP_KIND,
          });
        },
      });
      return element;
    },
    getElementById: () => null,
  };

  const options: SandboxActivationOptions = {
    instanceId: "ext.sandbox.boundary",
    extensionName: "Boundary contract",
    source: { kind: "inline", code: "export function activate() {}" },
    registerCommand: () => {},
    registerTool: () => {},
    unregisterTool: () => {},
    subscribeAgentEvents: () => () => {},
    llmComplete: () => Promise.reject(new Error("not expected")),
    httpFetch: () => Promise.reject(new Error("not expected")),
    storageGet: () => Promise.resolve(null),
    storageSet: () => Promise.resolve(),
    storageDelete: () => Promise.resolve(),
    storageKeys: () => Promise.resolve([]),
    clipboardWriteText: () => Promise.resolve(),
    injectAgentContext: () => {},
    steerAgent: () => {},
    followUpAgent: () => {},
    listSkills: () => Promise.resolve([]),
    readSkill: () => Promise.reject(new Error("not expected")),
    installSkill: () => Promise.resolve(),
    uninstallSkill: () => Promise.resolve(),
    downloadFile: () => {},
    registerConnection: () => "connection-id",
    unregisterConnection: () => {},
    listConnections: () => Promise.resolve([]),
    getConnection: () => Promise.resolve(null),
    getConnectionSecrets: () => Promise.resolve(null),
    setConnectionSecrets: () => Promise.resolve(),
    clearConnectionSecrets: () => Promise.resolve(),
    markConnectionValidated: () => Promise.resolve(),
    markConnectionInvalid: () => Promise.resolve(),
    markConnectionStatus: () => Promise.resolve(),
    registerModelProvider: () => "provider-id",
    unregisterModelProvider: () => {},
    refreshModelProviders: () => Promise.resolve(),
    isCapabilityEnabled: () => false,
    formatCapabilityError: (capability) => `Denied ${capability}`,
    toast: () => {},
  };

  Reflect.set(globalThis, "document", fakeDocument);
  Reflect.set(globalThis, "window", {});
  Reflect.set(globalThis, "HTMLElement", EventTarget);

  try {
    const activation = activateExtensionInSandbox(options).then((handle) => {
      activationSettled = true;
      return handle;
    });

    const mountedIframe = getIframe();
    assert.ok(mountedIframe);
    assert.equal(mountedIframe.getAttribute("sandbox"), "allow-scripts");
    assert.equal(mountedIframe.getAttribute("aria-hidden"), "true");
    mountedIframe.dispatchEvent(new Event("load"));
    const transferredPort = getSandboxPort();
    assert.ok(transferredPort);

    transferredPort.postMessage({
      channel: SANDBOX_CHANNEL,
      instanceId: options.instanceId,
      direction: "host_to_sandbox",
      kind: "event",
      event: "ready",
      data: null,
    });
    await new Promise<void>((resolve) => setTimeout(resolve, 20));
    assert.equal(activationSettled, false, "host-direction messages must not activate the sandbox");

    transferredPort.postMessage({
      channel: SANDBOX_CHANNEL,
      instanceId: options.instanceId,
      direction: "sandbox_to_host",
      kind: "event",
      event: "ready",
      data: null,
    });
    const handle = await activation;

    const deactivateRequest = new Promise<Record<string, unknown>>((resolve) => {
      transferredPort.addEventListener("message", (event: MessageEvent<unknown>) => {
        const envelope = event.data;
        if (typeof envelope !== "object" || envelope === null || Array.isArray(envelope)) return;
        const payload = envelope as Record<string, unknown>;
        if (payload.kind !== "request" || payload.method !== "deactivate") return;
        resolve(payload);
        transferredPort.postMessage({
          channel: SANDBOX_CHANNEL,
          instanceId: options.instanceId,
          direction: "sandbox_to_host",
          kind: "response",
          requestId: payload.requestId,
          ok: true,
          result: null,
        });
      });
      transferredPort.start();
    });

    const deactivated = handle.deactivate();
    const request = await deactivateRequest;
    assert.equal(request.direction, "host_to_sandbox");
    await deactivated;
  } finally {
    if (previousDocument === undefined) Reflect.deleteProperty(globalThis, "document");
    else Reflect.set(globalThis, "document", previousDocument);
    if (previousWindow === undefined) Reflect.deleteProperty(globalThis, "window");
    else Reflect.set(globalThis, "window", previousWindow);
    if (previousHTMLElement === undefined) Reflect.deleteProperty(globalThis, "HTMLElement");
    else Reflect.set(globalThis, "HTMLElement", previousHTMLElement);
    getSandboxPort()?.close();
  }
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
  const connectionError = decodeToolDetailsOfKind(result.details, "connection_error");
  assert.ok(connectionError);
  assert.equal(connectionError.errorCode, "missing_connection");
  assert.equal(connectionError.connectionId, "ext.apollo.apollo");
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
  const connectionError = decodeToolDetailsOfKind(result.details, "connection_error");
  assert.ok(connectionError);
  assert.equal(connectionError.errorCode, "connection_auth_failed");
  assert.ok(!connectionError.reason?.includes("top-secret-api-key"));

  const text = result.content[0]?.type === "text" ? result.content[0].text : "";
  assert.ok(!text.includes("top-secret-api-key"));

  const state = await connectionManager.getState("ext.apollo.apollo");
  assert.equal(state?.status, "error");
  assert.ok(!state?.lastError?.includes("top-secret-api-key"));
});

void test("connection preflight still redacts auth failures when status persistence fails", async () => {
  const settings = new FailingConnectionStoreSettings();
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

  settings.armConnectionStoreFailure();

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

  const result = await tools[0].execute("tool-call-2b", { company: "Acme" });
  const connectionError = decodeToolDetailsOfKind(result.details, "connection_error");
  assert.ok(connectionError);
  assert.equal(connectionError.errorCode, "connection_auth_failed");
  assert.equal(connectionError.status, "error");
  assert.ok(!connectionError.reason?.includes("top-secret-api-key"));

  const text = result.content[0]?.type === "text" ? result.content[0].text : "";
  assert.ok(!text.includes("top-secret-api-key"));

  const state = await connectionManager.getState("ext.apollo.apollo");
  assert.equal(state?.status, "connected");
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
  const connectionError = decodeToolDetailsOfKind(result.details, "connection_error");
  assert.ok(connectionError);
  assert.equal(connectionError.errorCode, "connection_auth_failed");
  assert.equal(connectionError.connectionId, "ext.multi.billing");
  assert.ok(!connectionError.reason?.includes("billing-key"));

  const crmState = await connectionManager.getState("ext.multi.crm");
  const billingState = await connectionManager.getState("ext.multi.billing");

  assert.equal(crmState?.status, "connected");
  assert.equal(billingState?.status, "error");
});

void test("extension model providers use host-owned connection secrets and unload cleanly", async () => {
  const settings = new MemorySettingsStore();
  settings.writeRaw(EXTENSIONS_REGISTRY_STORAGE_KEY, {
    version: 2,
    items: [createStoredEntry({
      id: "ext.provider.test",
      name: "Provider extension",
      trust: "local-module",
    })],
  });

  let authorization = "";
  const fetchFn: typeof globalThis.fetch = (_input, init) => {
    authorization = new Headers(init?.headers).get("authorization") ?? "";
    return Promise.resolve(new Response(JSON.stringify({ data: [{ id: "discovered-model" }] }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }));
  };
  const modelRuntime = new BrowserModelRuntime({
    providerKeys: new EmptyProviderKeys(),
    modelCatalogs: new InMemoryModelsStore(),
    getProxyUrl: () => Promise.resolve(undefined),
    fetchFn,
  });
  const connectionManager = createConnectionManager(settings);

  const manager = new ExtensionRuntimeManager({
    settings,
    connectionManager,
    modelRuntime,
    getActiveAgent: () => null,
    refreshRuntimeTools: () => Promise.resolve(),
    refreshRuntimeModels: async () => {
      await modelRuntime.refresh({ allowNetwork: true });
    },
    reservedToolNames: new Set<string>(),
    loadExtensionFromSource: async (api) => {
      api.connections.register({
        id: "account",
        title: "Model account",
        capability: "model inference",
        authKind: "api_key",
        secretFields: [{ id: "apiKey", label: "API key", required: true }],
        httpAuth: {
          placement: "header",
          headerName: "Authorization",
          valueTemplate: "Bearer {apiKey}",
          allowedHosts: ["models.example.com"],
        },
      });
      await api.connections.setSecrets("account", { apiKey: "host-secret" });
      await api.connections.markValidated("account");
      api.models.registerProvider({
        id: "acme",
        name: "Acme models",
        api: "openai-responses",
        baseUrl: "https://models.example.com/v1",
        models: [{ id: "baseline-model", contextWindow: 128_000, maxTokens: 16_000 }],
        connection: "account",
        apiKeySecret: "apiKey",
      });

      return {
        deactivate: () => Promise.resolve(),
      };
    },
  });

  await manager.initialize();

  const providerId = "ext.provider.test.acme";
  assert.equal(authorization, "Bearer host-secret");
  assert.equal((await modelRuntime.models.getAuth(providerId))?.auth.apiKey, "host-secret");
  assert.deepEqual(
    modelRuntime.models.getModels(providerId).map((model) => model.id),
    ["baseline-model", "discovered-model"],
  );
  assert.deepEqual(manager.list()[0]?.modelProviderIds, [providerId]);

  await manager.setExtensionEnabled("ext.provider.test", false);
  assert.equal(modelRuntime.models.getProvider(providerId), undefined);
});
