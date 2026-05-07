import assert from "node:assert/strict";
import { test } from "node:test";

import type { Agent } from "@earendil-works/pi-agent-core";

import { ConnectionManager } from "../src/connections/manager.ts";
import { buildRuntimeManagerActivationBridge } from "../src/extensions/runtime-manager-activation.ts";
import {
  getDefaultPermissionsForTrust,
  type ExtensionCapability,
} from "../src/extensions/permissions.ts";
import type { StoredExtensionEntry } from "../src/extensions/store.ts";

class MemorySettingsStore {
  private readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.get(key));
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }
}

function createEntry(id: string): StoredExtensionEntry {
  const now = new Date().toISOString();
  return {
    id,
    name: "Acme Extension",
    enabled: true,
    source: {
      kind: "inline",
      code: "export function activate() {}",
    },
    trust: "inline-code",
    permissions: getDefaultPermissionsForTrust("inline-code"),
    createdAt: now,
    updatedAt: now,
  };
}

function buildBridge(args: {
  entry: StoredExtensionEntry;
  settings: MemorySettingsStore;
  connectionManager: ConnectionManager;
  getRequiredActiveAgent?: () => Agent;
  afterInjectAgentContext?: () => Promise<void> | void;
  runExtensionHttpFetch: (url: string, options?: {
    method?: "GET" | "POST" | "PUT" | "PATCH" | "DELETE" | "HEAD";
    headers?: Record<string, string>;
    body?: string;
    timeoutMs?: number;
    connection?: string;
  }) => Promise<{
    status: number;
    statusText: string;
    headers: Record<string, string>;
    body: string;
  }>;
}) {
  const getRequiredActiveAgent = args.getRequiredActiveAgent ?? (() => {
    throw new Error("agent access should not be used in this test");
  });

  return buildRuntimeManagerActivationBridge({
    entry: args.entry,
    settings: args.settings,
    connectionManager: args.connectionManager,
    getRequiredActiveAgent,
    afterInjectAgentContext: args.afterInjectAgentContext,
    runExtensionLlmCompletion: () => Promise.resolve({
      content: "",
      model: "test/model",
    }),
    runExtensionHttpFetch: args.runExtensionHttpFetch,
    writeExtensionClipboard: async () => {},
    triggerExtensionDownload: () => {},
    isCapabilityEnabled: (_capability: ExtensionCapability) => true,
    formatCapabilityError: (capability: ExtensionCapability) => `denied:${capability}`,
    showToastMessage: () => {},
    widgetApiV2Enabled: false,
  });
}

void test("injectContext appends a message and triggers host sync hook", async () => {
  const settings = new MemorySettingsStore();
  const entry = createEntry("ext.inject");
  const connectionManager = new ConnectionManager({ settings });
  const messages: Agent["state"]["messages"] = [];
  let syncCalls = 0;

  const agent = {
    state: { messages },
    steer: () => {},
    followUp: () => {},
  } as unknown as Agent;

  const bridge = buildBridge({
    entry,
    settings,
    connectionManager,
    getRequiredActiveAgent: () => agent,
    afterInjectAgentContext: () => {
      syncCalls += 1;
    },
    runExtensionHttpFetch: () => Promise.resolve({
      status: 200,
      statusText: "OK",
      headers: {},
      body: "ok",
    }),
  });

  bridge.host.injectAgentContext?.("context please");
  await Promise.resolve();

  assert.equal(syncCalls, 1);
  assert.equal(messages.length, 1);
  assert.equal(messages[0]?.role, "user");

  const firstContent = messages[0]?.content[0];
  assert.ok(firstContent && firstContent.type === "text");
  if (!firstContent || firstContent.type !== "text") {
    assert.fail("expected extension context injection to append a text message");
  }

  const injectedMessageJson = JSON.stringify(messages[0]);
  assert.match(injectedMessageJson, /\[Extension Acme Extension\]/);
  assert.match(injectedMessageJson, /context please/);
});

void test("connection-aware http fetch injects auth header for allowed hosts", async () => {
  const settings = new MemorySettingsStore();
  const entry = createEntry("ext.acme");
  const connectionManager = new ConnectionManager({ settings });

  connectionManager.registerDefinition(entry.id, {
    id: "ext.acme.acme",
    title: "Acme API",
    capability: "query Acme records",
    authKind: "api_key",
    secretFields: [{ id: "apiKey", label: "API key", required: true }],
    httpAuth: {
      placement: "header",
      headerName: "Authorization",
      valueTemplate: "Bearer {apiKey}",
      allowedHosts: ["api.acme.com"],
    },
  });

  await connectionManager.setSecrets(entry.id, "ext.acme.acme", {
    apiKey: "top-secret",
  });

  let capturedHeaders: Record<string, string> | undefined;
  let capturedConnectionId = "";

  const bridge = buildBridge({
    entry,
    settings,
    connectionManager,
    runExtensionHttpFetch: (_url, options) => {
      capturedHeaders = options?.headers;
      capturedConnectionId = options?.connection ?? "";
      return Promise.resolve({
        status: 200,
        statusText: "OK",
        headers: {},
        body: "ok",
      });
    },
  });

  await bridge.host.httpFetch("https://api.acme.com/v1/search?q=test", {
    method: "GET",
    headers: {
      Accept: "application/json",
    },
    connection: "acme",
  });

  assert.equal(capturedConnectionId, "ext.acme.acme");
  assert.equal(capturedHeaders?.Authorization, "Bearer top-secret");
  assert.equal(capturedHeaders?.Accept, "application/json");
});

void test("connection-aware http fetch blocks hosts outside allowedHosts", async () => {
  const settings = new MemorySettingsStore();
  const entry = createEntry("ext.blocked");
  const connectionManager = new ConnectionManager({ settings });

  connectionManager.registerDefinition(entry.id, {
    id: "ext.blocked.api",
    title: "Blocked API",
    capability: "query blocked records",
    authKind: "api_key",
    secretFields: [{ id: "apiKey", label: "API key", required: true }],
    httpAuth: {
      placement: "header",
      headerName: "Authorization",
      valueTemplate: "Bearer {apiKey}",
      allowedHosts: ["api.allowed.com"],
    },
  });

  await connectionManager.setSecrets(entry.id, "ext.blocked.api", {
    apiKey: "blocked-secret",
  });

  const bridge = buildBridge({
    entry,
    settings,
    connectionManager,
    runExtensionHttpFetch: () => {
      return Promise.reject(new Error("http fetch should not be reached for blocked hosts"));
    },
  });

  await assert.rejects(
    async () => bridge.host.httpFetch("https://evil.example.com/data", {
      connection: "api",
    }),
    /not allowed for this connection/i,
  );
});

void test("connection-aware http fetch marks auth failures and throws structured connection error", async () => {
  const settings = new MemorySettingsStore();
  const entry = createEntry("ext.auth");
  const connectionManager = new ConnectionManager({ settings });

  connectionManager.registerDefinition(entry.id, {
    id: "ext.auth.api",
    title: "Auth API",
    capability: "query auth records",
    authKind: "api_key",
    secretFields: [{ id: "apiKey", label: "API key", required: true }],
    httpAuth: {
      placement: "header",
      headerName: "Authorization",
      valueTemplate: "Bearer {apiKey}",
      allowedHosts: ["api.auth.com"],
    },
  });

  await connectionManager.setSecrets(entry.id, "ext.auth.api", {
    apiKey: "auth-secret",
  });

  const bridge = buildBridge({
    entry,
    settings,
    connectionManager,
    runExtensionHttpFetch: () => {
      return Promise.resolve({
        status: 401,
        statusText: "Unauthorized",
        headers: {},
        body: "unauthorized",
      });
    },
  });

  await assert.rejects(
    async () => bridge.host.httpFetch("https://api.auth.com/v1/data", {
      connection: "api",
    }),
    /failed authentication/i,
  );

  const state = await connectionManager.getState("ext.auth.api");
  assert.equal(state?.status, "error");
});
