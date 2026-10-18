import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels, fauxProvider } from "@earendil-works/pi-ai";
import { createRegistry, Harness, MemoryStorage } from "@earendil-works/pi-durable";
import { ConversationPresenter } from "../src/durable/conversation-view.ts";
import { DurableSessions, SessionInfoDoc } from "../src/durable/sessions.ts";
import { createQueueDisplay } from "../src/taskpane/queue-display.ts";
import { SessionRuntimeManager, type SessionRuntime } from "../src/taskpane/session-runtime-manager.ts";

void test("native runtime lifecycle owns tab identity, rename, subscriptions and capability refresh", async () => {
  const models = createModels();
  const faux = fauxProvider();
  models.setProvider(faux.provider);
  const model = faux.getModel();
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const sessions = new DurableSessions(harness, models, model);
  const views: ConversationPresenter[] = [];
  let refreshes = 0;
  let disposals = 0;
  const manager = new SessionRuntimeManager({ sessions, createRuntime: async (options) => {
    const conversation = options.conversationId
      ? await sessions.resolve(options.conversationId) : await sessions.create(model, "off");
    assert.ok(conversation);
    const view = await ConversationPresenter.open(harness, conversation, models, model);
    views.push(view);
    const runtime: SessionRuntime = {
      runtimeId: `tab-${views.length}`, view, tools: [], title: "", hasExplicitTitle: false,
      lockState: "idle", refreshCapabilities: () => { refreshes += 1; return Promise.resolve(); },
      dispose: () => { disposals += 1; void view.dispose(); },
      queueDisplay: createQueueDisplay({ view }),
    };
    return runtime;
  } });
  try {
    // Startup refreshes can arrive before any tabs exist; they must not strand
    // the in-flight slot and suppress later extension capability updates.
    await manager.refreshCapabilities();
    const first = await manager.createRuntime({ activate: true, autoRestoreLatest: false });
    const second = await manager.createRuntime({ activate: false, autoRestoreLatest: false });
    assert.deepEqual(manager.snapshotTabs().map((tab) => tab.title), ["Chat 1", "Chat 2"]);
    assert.equal(manager.findRuntimeBySessionId(String(first.view.conversation.id)), first);
    await manager.renameRuntime(first.runtimeId, "Renamed");
    assert.equal(manager.snapshotTabs()[0]?.title, "Renamed");
    assert.equal((await harness.snapshot(SessionInfoDoc, first.view.conversation.id, context))?.title, "Renamed");
    await manager.refreshCapabilities();
    assert.equal(refreshes, 2);
    let release: () => void = () => {};
    const gate = new Promise<void>((resolve) => { release = resolve; });
    let passes = 0;
    first.refreshCapabilities = async () => { passes += 1; if (passes === 1) await gate; };
    const pending = manager.refreshCapabilities();
    const coalesced = manager.refreshCapabilities();
    release();
    await Promise.all([pending, coalesced]);
    assert.equal(passes, 2);
    assert.equal(manager.moveRuntime(second.runtimeId, -1), true);
    assert.deepEqual(manager.snapshotTabs().map((tab) => tab.title), ["Chat 2", "Renamed"]);
    manager.setRuntimeLockState(first.runtimeId, "holding_lock");
    assert.equal(manager.snapshotTabs()[1]?.lockState, "holding_lock");
    let emissions = 0;
    const unsubscribe = manager.subscribe(() => { emissions += 1; });
    await first.view.refresh();
    assert.equal(emissions, 2);
    manager.closeRuntime(first.runtimeId);
    const afterClose = emissions;
    await first.view.refresh();
    assert.equal(emissions, afterClose);
    assert.equal(disposals, 1);
    assert.equal(manager.getActiveRuntime(), second);
    manager.closeRuntime(second.runtimeId);
    assert.equal(disposals, 1, "last tab remains open");
    unsubscribe();
  } finally {
    for (const view of views) await view.dispose();
    await harness.close(context);
  }
});

void test("a refresh requested while an empty startup pass settles reaches newly registered tabs", async () => {
  const models = createModels();
  const faux = fauxProvider();
  models.setProvider(faux.provider);
  const model = faux.getModel();
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const sessions = new DurableSessions(harness, models, model);
  const conversation = await sessions.create(model, "off");
  const view = await ConversationPresenter.open(harness, conversation, models, model);
  let refreshes = 0;
  const manager = new SessionRuntimeManager({ sessions, createRuntime: () => {
    throw new Error("This contract registers an existing runtime");
  } });
  try {
    const startup = manager.refreshCapabilities();
    manager.registerRuntime({
      runtimeId: "first-tab", view, tools: [], title: "", hasExplicitTitle: false,
      lockState: "idle", refreshCapabilities: () => { refreshes += 1; return Promise.resolve(); },
      dispose: () => {}, queueDisplay: createQueueDisplay({ view }),
    });
    const requested = manager.refreshCapabilities();
    await Promise.all([startup, requested]);
    assert.equal(refreshes, 1, "the coalesced caller awaits the newly requested pass");
    await manager.refreshCapabilities();
    assert.equal(refreshes, 2, "subsequent passes remain available");
  } finally {
    await view.dispose();
    await harness.close(context);
  }
});
