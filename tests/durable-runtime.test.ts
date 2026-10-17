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
      storage, workbookId: "recovery", model: { provider: model.provider, modelId: model.id },
      models: routeModels(models, () => { started?.(); return source; }, () => lifecycle),
      onOpened: (harness) => {
        lifecycle = createModelRecoveryLifecycle(() => harness.close(context), target);
      },
      buildCapabilities: () => Promise.resolve({ tools: [], systemPrompt: "Recovery test" }),
    });
    try {
      const submitted = await runtime.conversation.submit({ type: "input", content: "Resume me", requestId: "once" }, context);
      await requestStarted;
      source.push({ type: "error", reason: "error", error: {
        ...fauxAssistantMessage(""), stopReason: "error", errorMessage: "Load failed",
      } });
      if (navigation === "none") {
        const settled = await submitted.wait(context);
        assert.equal(settled.status, "unanswered");
        assert.match(JSON.stringify(await runtime.conversation.context(context)), /Load failed/);
        return;
      }
      setTimeout(() => {
        if (navigation === "explicit reload") prepareDurableReload(target);
        else target.dispatchEvent(new Event(navigation));
      }, 0);
      await runtime.harness.closed;
      lifecycle?.dispose();
      faux.setResponses([fauxAssistantMessage("Recovered answer")]);
      const reopened = await openDurableRuntime({
        storage, workbookId: "recovery", model: { provider: model.provider, modelId: model.id }, models,
        buildCapabilities: () => Promise.resolve({ tools: [], systemPrompt: "Recovery test" }),
      });
      try {
        assert.equal(reopened.conversation.id, runtime.conversation.id);
        const receipt = await reopened.harness.submission(submitted.id, context);
        assert.ok(receipt);
        assert.equal((await receipt.wait(context)).status, "done");
        const entries = (await reopened.conversation.entries({}, 100, undefined, context)).items;
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
      model: { provider, modelId: model.id },
      models: routeModels(models, createOfficeStreamFn(() => Promise.resolve("https://localhost:3456"), models)),
      buildCapabilities: () => Promise.resolve({ tools: [], systemPrompt: "Transport test" }),
    });
    try {
      const submission = await runtime.conversation.submit({ type: "input", content: "Hello" }, context);
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
      await runtime.refresh();
      const repeated = await runtime.conversation.submit({ type: "input", content: "Again" }, context);
      assert.equal((await repeated.wait(context)).status, "done");
      assert.equal(captured[1]?.sessionId, request.sessionId);
      assert.deepEqual(getPayloadSnapshots().filter((snapshot) => snapshot.sessionId === request.sessionId)
        .map((snapshot) => snapshot.prefixChangeReasons), [[], []]);
      assert.doesNotMatch(JSON.stringify(await runtime.status()), /synthetic-test/);
    } finally { await runtime.harness.close(context); }
  });
}
