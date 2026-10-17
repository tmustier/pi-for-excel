import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { createAssistantMessageEventStream } from "@earendil-works/pi-ai";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import type { Api, Model } from "@earendil-works/pi-ai";
import { createRegistry, Harness, InboxDoc, MemoryStorage } from "@earendil-works/pi-durable";
import { Type } from "typebox";
import { ConversationPresenter } from "../src/durable/conversation-view.ts";
import { DurableSessions } from "../src/durable/sessions.ts";
import type { SessionRuntime } from "../src/taskpane/session-runtime-manager.ts";
import { executeExtensionVerificationCommand as execute, type ExtensionVerificationOptions } from "../src/taskpane/background-extension-verification.ts";

const model: Model<Api> = {
  id: "bridge-test", name: "Bridge test", provider: "test", api: "openai-completions", baseUrl: "https://example.com",
  reasoning: true, input: ["text"], cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 }, contextWindow: 32000, maxTokens: 4000,
};

async function fixture(holdGeneration = false) {
  const models = createModels();
  const source = createAssistantMessageEventStream();
  const faux = fauxProvider();
  if (holdGeneration) models.setProvider({ ...faux.provider, streamSimple: () => source });
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const sessions = new DurableSessions(harness, models, model);
  const conversation = await sessions.create(holdGeneration ? faux.getModel() : model, "off");
  const view = await ConversationPresenter.open(harness, conversation, models, model);
  let refreshes = 0;
  let rawCalls = 0;
  const runtime: SessionRuntime = {
    runtimeId: "native", view, title: "Test", hasExplicitTitle: false, lockState: "idle",
    tools: [{ name: "write_cells", label: "Write", description: "test", parameters: Type.Object({}),
      execute: () => { rawCalls++; return Promise.reject(new Error("Must not bypass installed tools")); } }],
    queueDisplay: { takeQueuedInputs: () => Promise.resolve([]), editQueuedInput: () => Promise.resolve(false), attach: () => {}, detach: () => {}, dispose: () => {} },
    refreshCapabilities: () => { refreshes++; return Promise.resolve(); }, dispose: () => {},
  };
  // These commands never touch extension/connection/model managers. Keep the fixture
  // limited to its real Durable execution owner and catalogue, not an Agent facade.
  // eslint-disable-next-line @typescript-eslint/consistent-type-assertions -- Unused managers intentionally absent in this focused command fixture.
  const options = { getActiveRuntime: () => runtime, getSessions: () => sessions,
    durable: { sessions, execute: (type: string, json: string) => Promise.resolve({ type, json }) },
  } as ExtensionVerificationOptions;
  return { harness, conversation, runtime, options, refreshes: () => refreshes, rawCalls: () => rawCalls,
    close: async () => { source.end(fauxAssistantMessage("Stopped")); await view.dispose(); await harness.close(context); } };
}

void test("bridge inspects and exports complete paginated history after reset", async () => {
  const f = await fixture();
  try {
    await f.harness.commit(async (tx) => {
      for (let index = 0; index < 205; index++) await tx.appendEntry(f.conversation.id, {
        kind: "test.message", model: [{ role: "user", content: `message-${index}`, timestamp: index }],
      });
      await tx.appendEntry(f.conversation.id, { kind: "test.message", model: [{
        role: "assistant", content: [{ type: "text", text: "Old answer" }], api: model.api, model: model.id, provider: model.provider,
        stopReason: "stop", timestamp: 210, usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0,
          cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } },
      }] });
    }, context);
    await f.conversation.reset(undefined, context);
    assert.deepEqual(await execute("assertLastAssistantText", { expected: "Old answer" }, f.options), {
      matches: true, expectedLength: 10, actualLength: 10,
    });
    const session = await f.options.getSessions?.().loadSession(String(f.conversation.id));
    assert.equal(session?.messages.length, 206);
    assert.deepEqual(await execute("getMessages", {}, f.options), { messages: session?.messages });
    assert.deepEqual(await execute("exportTranscript", {}, f.options), session);
    assert.deepEqual(await execute("listSessions", {}, f.options), { sessions: await f.options.durable.sessions.getAllMetadata() });
  } finally { await f.close(); }
});

void test("bridge uses native configure and inbox, and abort withdraws queued input", async () => {
  const f = await fixture(true);
  try {
    assert.deepEqual(await execute("configure", { thinkingLevel: "high" }, f.options), { configured: true });
    assert.equal((await f.conversation.agent(context)).thinkingLevel, "high");
    assert.equal(f.refreshes(), 1);
    await assert.rejects(execute("configure", { tools: [] }, f.options), /Invalid configure/);
    await f.conversation.submit({ type: "input", content: "Running" }, context);
    const submission = await f.conversation.submit({ type: "input", content: "Queued" }, context);
    const inbox = await f.harness.snapshot(InboxDoc, f.conversation.id, context);
    assert.equal(inbox?.items.length, 1);
    assert.deepEqual(await execute("getPendingMessages", {}, f.options), { pending: inbox?.items ?? [] });
    await f.runtime.view.refresh();
    await assert.rejects(execute("configure", { thinkingLevel: "low" }, f.options), /idle active chat/);
    assert.deepEqual(await execute("abort", {}, f.options), { aborted: true });
    assert.deepEqual(await execute("getPendingMessages", {}, f.options), { pending: [] });
    const receipt = await submission.wait(context);
    assert.equal(receipt.status, "unanswered");
  } finally { await f.close(); }
});

void test("tool bridge fails closed without installed dispatch and never executes raw tools", async () => {
  const f = await fixture();
  try {
    await assert.rejects(execute("executeTool", { name: "write_cells", args: {} }, f.options), /Installed Durable/);
    let dispatched = 0;
    f.options.executeTool = (runtime, name, args) => {
      dispatched++;
      assert.equal(runtime, f.runtime);
      assert.equal(name, "write_cells");
      assert.deepEqual(args, { cells: [] });
      return Promise.resolve({ gated: true });
    };
    assert.deepEqual(await execute("executeTool", { name: "write_cells", args: { cells: [] } }, f.options), { gated: true });
    await assert.rejects(execute("executeTool", { name: "missing", args: {} }, f.options), /not available/);
    assert.equal(dispatched, 1);
    assert.equal(f.rawCalls(), 0);
    assert.deepEqual(await execute("durableStatus", null, f.options), { type: "durableStatus", json: "{}" });
  } finally { await f.close(); }
});

void test("reload seals Durable immediately before location.reload", async () => {
  const f = await fixture();
  const previous = Object.getOwnPropertyDescriptor(globalThis, "window");
  const events: string[] = [];
  let scheduled: (() => void) | undefined;
  const target = Object.assign(new EventTarget(), {
    setTimeout: (callback: () => void, delay: number) => {
      assert.equal(delay, 100);
      scheduled = callback;
      return 1;
    },
    location: { reload: () => events.push("reload") },
  });
  target.addEventListener("pi-durable-suspend", () => events.push("suspend"));
  Object.defineProperty(globalThis, "window", { value: target, configurable: true });
  try {
    assert.deepEqual(await execute("reloadTaskpane", null, f.options), { scheduled: true });
    assert.deepEqual(events, []);
    assert.ok(scheduled);
    scheduled();
    assert.deepEqual(events, ["suspend", "reload"]);
  } finally {
    if (previous) Object.defineProperty(globalThis, "window", previous);
    else Reflect.deleteProperty(globalThis, "window");
    await f.close();
  }
});
