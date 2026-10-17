import assert from "node:assert/strict";
import { test } from "node:test";
import { setTimeout as delay } from "node:timers/promises";

import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createRegistry, Harness, InboxDoc, MemoryStorage, type AgentEvent } from "@earendil-works/pi-durable";
import { ConversationPresenter } from "../src/durable/conversation-view.ts";
import { createAssistantMessageEventStream, createModels, fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai";
import { createExtensionAPI } from "../src/commands/extension-api.ts";

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
  getActiveConversation?: () => ConversationPresenter;
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
  const getActiveConversation = args.getActiveConversation ?? (() => {
    throw new Error("agent access should not be used in this test");
  });

  return buildRuntimeManagerActivationBridge({
    entry: args.entry,
    settings: args.settings,
    connectionManager: args.connectionManager,
    getActiveConversation,
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

void test("injected extension context is available to the next agent request", async () => {
  const settings = new MemorySettingsStore();
  const entry = createEntry("ext.inject");
  const connectionManager = new ConnectionManager({ settings });
  const faux = fauxProvider();
  const models = createModels();
  faux.setResponses([fauxAssistantMessage("done")]);
  let outboundContext = "";
  models.setProvider({ ...faux.provider,
    streamSimple: (model, transcript, options) => {
      outboundContext = JSON.stringify(transcript);
      return faux.provider.streamSimple(model, transcript, options);
    },
  });
  const harness = await Harness.open(new MemoryStorage(), { registry: createRegistry(), models }, context);
  const model = faux.getModel();
  const conversation = await harness.root(context, { agent: { model: { provider: model.provider, modelId: model.id } } });
  const presenter = await ConversationPresenter.open(harness, conversation, models, model);

  const bridge = buildBridge({
    entry,
    settings,
    connectionManager,
    getActiveConversation: () => presenter,
    runExtensionHttpFetch: () => Promise.resolve({
      status: 200,
      statusText: "OK",
      headers: {},
      body: "ok",
    }),
  });

  const api = createExtensionAPI(bridge.host);
  const events: AgentEvent[] = [];
  const unsubscribe = api.onAgentEvent((event) => events.push(event));
  try {
    assert.equal(api.agent.raw, conversation);
    await api.agent.injectContext("context please");
    assert.equal(presenter.current.busy, false);
    assert.match(JSON.stringify(api.agent.snapshot.messages), /context please/);
    const submission = await conversation.submit({ type: "input", content: "continue" }, context);
    await submission.wait(context);
    assert.match(outboundContext, /\[Extension Acme Extension\].*context please/u);
    for (let attempt = 0; attempt < 100 && !events.some((event) => event.type === "message_end" && event.entry.kind === "extension.context"); attempt++) {
      await delay(10);
    }
    assert.ok(events.some((event) => event.type === "message_end" && event.entry.kind === "extension.context"));
  } finally {
    unsubscribe();
    await presenter.dispose();
    await harness.close(context);
  }
});

void test("extension input uses native busy admission modes and passive writes", async () => {
  const settings = new MemorySettingsStore();
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: Infinity });
  const source = createAssistantMessageEventStream();
  models.setProvider({ ...faux.provider, streamSimple: () => source });
  const model = faux.getModel();
  const harness = await Harness.open(new MemoryStorage(), { registry: createRegistry(), models }, context);
  const conversation = await harness.root(context, { agent: { model: { provider: model.provider, modelId: model.id } } });
  const presenter = await ConversationPresenter.open(harness, conversation, models, model);
  const bridge = buildBridge({
    entry: createEntry("ext.busy"), settings,
    connectionManager: new ConnectionManager({ settings }),
    getActiveConversation: () => presenter,
    runExtensionHttpFetch: () => Promise.reject(new Error("not expected")),
  });
  const api = createExtensionAPI(bridge.host);
  try {
    await conversation.submit({ type: "input", content: "running" }, context);
    await api.agent.steer("correct course");
    await api.agent.followUp("next task");
    await api.agent.injectContext("passive note");
    const inbox = await harness.snapshot(InboxDoc, conversation.id, context);
    assert.deepEqual(inbox?.items.map((item) => item.mode), ["steer", "followUp", "write"]);
    assert.match(JSON.stringify(inbox), /correct course/);
    assert.match(JSON.stringify(inbox), /next task/);
    assert.match(JSON.stringify(inbox), /passive note/);
    await conversation.abort(context);
  } finally {
    source.end(fauxAssistantMessage("Stopped"));
    await presenter.dispose();
    await harness.close(context);
  }
  await assert.rejects(api.agent.injectContext("closed write"));
  await assert.rejects(api.agent.steer("closed input"));
  await assert.rejects(api.agent.followUp("closed follow-up"));
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
