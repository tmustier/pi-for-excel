import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxProvider, fauxAssistantMessage } from "@earendil-works/pi-ai/providers/faux";
import { MemoryStorage } from "@earendil-works/pi-durable";
import { adaptTools, openDurableRuntime, routeModels } from "../src/durable/runtime.ts";
import { createAllTools } from "../src/tools/index.ts";
import { createOfficeStreamFn, getPayloadSnapshots } from "../src/auth/stream-proxy.ts";
import { setDebugEnabled } from "../src/debug/debug.ts";
import { ProviderCredentialsStore } from "../src/storage/local/provider-credentials-store.ts";
import { createAssistantMessageEventStream } from "@earendil-works/pi-ai";
import { createModelRecoveryLifecycle, prepareDurableReload, type ModelRecoveryLifecycle } from "../src/durable/model-recovery.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import type { MutationHost } from "../src/durable/mutations.ts";

const mutationHost: MutationHost = {
  prepare: () => Promise.reject(new Error("Unexpected workbook mutation")),
  observe: () => Promise.reject(new Error("Unexpected workbook observation")),
};

for (const navigation of ["beforeunload", "pagehide", "explicit reload", "none"]) {
  void test(`routed Durable request remains recoverable after ${navigation}`, async () => {
    // Keep the disposable memory store available to a second Harness, like a
    // persistent backend reopened by a new page. This is not a storage test.
    class ReopenableMemory extends MemoryStorage {
      override async close(): Promise<void> {}
    }
    const storage = new ReopenableMemory();
    const target = new EventTarget();
    const models = createModels();
    const faux = fauxProvider({ models: [{ id: "recovery" }], tokensPerSecond: Infinity });
    models.setProvider(faux.provider);
    const model = faux.getModel();
    let lifecycle: ModelRecoveryLifecycle | undefined;
    let started: (() => void) | undefined;
    const requestStarted = new Promise<void>((resolve) => { started = resolve; });
    const source = createAssistantMessageEventStream();
    const runtime = await openDurableRuntime({
      storage, workbookId: "recovery", mutationHost, coordinator: createWorkbookCoordinator(), autoCompaction: false,
      models: routeModels(models, () => { started?.(); return source; }, () => lifecycle),
      onOpened: (harness) => {
        lifecycle = createModelRecoveryLifecycle(() => harness.close(context), target);
      },
    });
    const conversation = await runtime.harness.root(context, { agent: { model: { provider: model.provider, modelId: model.id } } });
    await runtime.install(conversation, { tools: [], systemPrompt: "Recovery test" }, (messages) => Promise.resolve(messages));
    try {
      const submitted = await conversation.submit({ type: "input", content: "Resume me", requestId: "once" }, context);
      await requestStarted;
      source.push({ type: "error", reason: "error", error: {
        ...fauxAssistantMessage(""), stopReason: "error", errorMessage: "Load failed",
      } });
      if (navigation === "none") {
        const settled = await submitted.wait(context);
        assert.equal(settled.status, "unanswered");
        assert.match(JSON.stringify(await conversation.context(context)), /Load failed/);
        return;
      }
      setTimeout(() => {
        if (navigation === "explicit reload") prepareDurableReload(target);
        else target.dispatchEvent(new Event(navigation));
      }, 0);
      await runtime.harness.closed;
      assert.equal((await storage.submission(submitted.id, context))?.status, "placed");
      lifecycle?.dispose();
      faux.setResponses([fauxAssistantMessage("Recovered answer")]);
      const reopened = await openDurableRuntime({
        storage, workbookId: "recovery", models, mutationHost, coordinator: createWorkbookCoordinator(), autoCompaction: false,
      });
      const restored = await reopened.harness.root(context);
      await reopened.install(restored, { tools: [], systemPrompt: "Recovery test" }, (messages) => Promise.resolve(messages));
      try {
        assert.equal(restored.id, conversation.id);
        const receipt = await reopened.harness.submission(submitted.id, context);
        assert.ok(receipt);
        assert.equal((await receipt.wait(context)).status, "done");
        const entries = (await restored.entries({}, 100, undefined, context)).items;
        assert.equal(entries.filter((entry) => entry.kind === "pi.user").length, 1);
        assert.match(JSON.stringify(entries), /Recovered answer/);
        assert.doesNotMatch(JSON.stringify(entries), /Load failed/);
      } finally { await reopened.harness.close(context); }
    } finally {
      lifecycle?.dispose();
      await runtime.harness.close(context);
    }
  });
}

void test("only exclusively read-only workbook tools opt into Durable replay", () => {
  const tools = adaptTools(createAllTools());
  assert.deepEqual(tools.filter((tool) => tool.replay === "safe").map((tool) => tool.name).sort(),
    ["get_workbook_overview", "read_range", "search_workbook", "trace_dependencies", "explain_formula"].sort());
  for (const name of ["write_cells", "fill_formula", "modify_structure", "charts", "comments", "view_settings", "instructions", "files", "python_run", "execute_office_js"]) {
    const tool = tools.find((candidate) => candidate.name === name);
    assert.ok(tool, name);
    assert.equal(tool.replay, undefined, `${name} keeps Durable's unsafe default`);
  }
});

void test("Durable preserves custom provider and model auth headers at provider dispatch without persisting secrets", async (t) => {
  const originalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
  Object.defineProperty(globalThis, "document", { configurable: true, value: new EventTarget() });
  t.after(() => {
    if (originalDocument) Object.defineProperty(globalThis, "document", originalDocument);
    else Reflect.deleteProperty(globalThis, "document");
  });
  const secret = "custom-header-credential-secret";
  const modelSecret = "custom-model-header-secret";
  const provider = "custom-header-test";
  const credentials = new ProviderCredentialsStore({
    get: () => Promise.resolve(secret), set: async () => {}, delete: async () => {}, list: () => Promise.resolve([provider]),
  });
  const models = createModels({ credentials });
  const faux = fauxProvider({ provider, models: [{ id: "headers" }], tokensPerSecond: Infinity });
  const model = { ...faux.getModel(), headers: { "x-model-token": modelSecret, "x-header-precedence": "model" } };
  const captured: Array<{ headers: Record<string, string> | undefined; transcript: string }> = [];
  models.setProvider({ ...faux.provider, models: [model], getModels: () => [model],
    auth: { apiKey: { name: "Custom header", resolve: ({ credential }) => Promise.resolve({
      auth: { headers: { "x-provider-token": credential?.key ?? "missing", "x-header-precedence": "provider" } },
    }) } },
    streamSimple: (requestModel, transcript, options) => {
      captured.push({ headers: options?.headers, transcript: JSON.stringify(transcript) });
      return faux.provider.streamSimple(requestModel, transcript, options);
    },
  });
  faux.setResponses([fauxAssistantMessage("header transport complete")]);
  const storage = new MemoryStorage();
  const writes: string[] = [];
  const commit = storage.commit.bind(storage);
  storage.commit = (batch, ctx) => {
    writes.push(JSON.stringify(batch));
    return commit(batch, ctx);
  };
  const runtime = await openDurableRuntime({
    storage, workbookId: "header-test", mutationHost, coordinator: createWorkbookCoordinator(), autoCompaction: false,
    models: routeModels(models, createOfficeStreamFn(() => Promise.resolve(undefined), models)),
  });
  try {
    const conversation = await runtime.harness.root(context, { agent: { model: { provider, modelId: model.id } } });
    await runtime.install(conversation, { tools: [], systemPrompt: "Header test" }, (messages) => Promise.resolve(messages));
    const submission = await conversation.submit({ type: "input", content: "Hello headers" }, context);
    assert.equal((await submission.wait(context)).status, "done");
    assert.equal(captured.length, 1);
    // This is the provider callback after the real underlying Models.applyAuth,
    // not the Office wrapper's options (which only explicitly forward apiKey).
    assert.deepEqual(captured[0]?.headers, {
      "x-provider-token": secret, "x-model-token": modelSecret, "x-header-precedence": "model",
    });
    const status = JSON.stringify(await runtime.status(conversation));
    assert.match(status, /header transport complete/);
    assert.ok(writes.length > 0);
    for (const value of [secret, modelSecret]) {
      assert.ok(!writes.join("\n").includes(value), "No credential in any Durable storage commit");
      assert.ok(!status.includes(value), "No credential in persisted conversation history");
      assert.ok(!captured[0]?.transcript.includes(value), "No credential in provider transcript");
    }
  } finally { await runtime.harness.close(context); }
});

for (const provider of ["direct-test", "anthropic", "openai-codex"]) {
  void test(`Durable uses browser credentials and Office transport for ${provider}`, async (t) => {
    const originalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
    const originalWindow = Object.getOwnPropertyDescriptor(globalThis, "window");
    const values = new Map<string, string>();
    Object.defineProperty(globalThis, "window", { configurable: true, value: { localStorage: {
      getItem: (key: string) => values.get(key) ?? null,
      setItem: (key: string, value: string) => values.set(key, value),
    } } });
    Object.defineProperty(globalThis, "document", { configurable: true, value: new EventTarget() });
    setDebugEnabled(true);
    t.after(() => {
      setDebugEnabled(false);
      if (originalWindow) Object.defineProperty(globalThis, "window", originalWindow);
      else Reflect.deleteProperty(globalThis, "window");
      if (originalDocument) Object.defineProperty(globalThis, "document", originalDocument);
      else Reflect.deleteProperty(globalThis, "document");
    });
    const previousFetch = globalThis.fetch;
    globalThis.fetch = (url) => {
      assert.equal(typeof url === "string" ? url : url instanceof URL ? url.href : url.url, "https://localhost:3456/healthz");
      return Promise.resolve(new Response("", { headers: { "x-pi-for-excel-proxy": "1", "x-pi-for-excel-codex-websocket-bridge": "1" } }));
    };
    t.after(() => { globalThis.fetch = previousFetch; });
    const secret = "sk-ant-oat-synthetic-test";
    const credentials = new ProviderCredentialsStore({
      get: () => Promise.resolve(secret), set: async () => {}, delete: async () => {}, list: () => Promise.resolve([provider]),
    });
    const models = createModels({ credentials });
    const faux = fauxProvider({ provider, models: [{ id: provider === "openai-codex" ? "gpt-5.6-luna" : "test" }], tokensPerSecond: Infinity });
    const captured: { url: string; sessionId: string | undefined; transport: string | undefined; key: string | undefined }[] = [];
    models.setProvider({ ...faux.provider,
      auth: { apiKey: { name: "Test", resolve: ({ credential }) => Promise.resolve({ auth: { ...(credential?.key ? { apiKey: credential.key } : {}) } }) } },
      streamSimple: (model, transcript, options) => {
        captured.push({ url: model.baseUrl, sessionId: options?.sessionId, transport: options?.transport, key: options?.apiKey });
        assert.ok(transcript.messages.some((message) => message.role === "system"));
        return faux.provider.streamSimple(model, transcript, options);
      },
    });
    faux.setResponses([fauxAssistantMessage("transport complete"), fauxAssistantMessage("second turn")]);
    const model = faux.getModel();
    const runtime = await openDurableRuntime({
      storage: new MemoryStorage(), workbookId: "transport-test",
      mutationHost, coordinator: createWorkbookCoordinator(), autoCompaction: false,
      models: routeModels(models, createOfficeStreamFn(() => Promise.resolve("https://localhost:3456"), models)),
    });
    const conversation = await runtime.harness.root(context, { agent: { model: { provider, modelId: model.id } } });
    const capabilities = { tools: [], systemPrompt: "Transport test" };
    await runtime.install(conversation, capabilities, (messages) => Promise.resolve(messages));
    try {
      const submission = await conversation.submit({ type: "input", content: "Hello" }, context);
      assert.equal((await submission.wait(context)).status, "done");
      assert.equal(captured.length, 1);
      const request = captured[0];
      assert.ok(request);
      assert.equal(request.key, secret);
      assert.ok(request.sessionId);
      if (provider === "direct-test") assert.equal(request.url, model.baseUrl);
      else assert.ok(request.url.startsWith("https://localhost:3456/?"));
      if (provider === "openai-codex") {
        assert.match(request.url, /pi_transport=codex-websocket/);
        assert.equal(request.transport, "sse");
        assert.match(request.sessionId, /^[0-9a-f]{8}-[0-9a-f]{4}-7/);
      }
      await runtime.install(conversation, capabilities, (messages) => Promise.resolve(messages));
      const repeated = await conversation.submit({ type: "input", content: "Again" }, context);
      assert.equal((await repeated.wait(context)).status, "done");
      assert.equal(captured[1]?.sessionId, request.sessionId);
      assert.deepEqual(getPayloadSnapshots().filter((snapshot) => snapshot.sessionId === request.sessionId)
        .map((snapshot) => snapshot.prefixChangeReasons), [[], []]);
      assert.doesNotMatch(JSON.stringify(await runtime.status(conversation)), /synthetic-test/);
    } finally { await runtime.harness.close(context); }
  });
}
