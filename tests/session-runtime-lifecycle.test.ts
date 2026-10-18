import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createAssistantMessageEventStream, createModels, fauxProvider, fauxAssistantMessage } from "@earendil-works/pi-ai";
import { createRegistry, Harness, MemoryStorage } from "@earendil-works/pi-durable";
import { ConversationPresenter } from "../src/durable/conversation-view.ts";
import { DurableSessions } from "../src/durable/sessions.ts";
import { createQueueDisplay } from "../src/taskpane/queue-display.ts";
import { SessionRuntimeManager } from "../src/taskpane/session-runtime-manager.ts";

void test("model selection configures native conversations and forks immutable ancestry", async () => {
  const models = createModels();
  const faux = fauxProvider({ models: [{ id: "first" }, { id: "second" }] });
  models.setProvider(faux.provider);
  const firstModel = faux.getModel("first");
  const secondModel = faux.getModel("second");
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const sessions = new DurableSessions(harness, models, firstModel);
  const views: ConversationPresenter[] = [];
  const manager = new SessionRuntimeManager({ sessions, createRuntime: async (options) => {
    const conversation = options.conversationId ? await sessions.resolve(options.conversationId)
      : await sessions.create(firstModel, "high");
    assert.ok(conversation);
    const view = await ConversationPresenter.open(harness, conversation, models, firstModel);
    views.push(view);
    const session = await sessions.loadSession(String(conversation.id));
    return { runtimeId: `runtime-${views.length}`, view, tools: [], title: session?.title ?? "",
      hasExplicitTitle: options.conversationId !== undefined, lockState: "idle",
      refreshCapabilities: async () => {}, dispose: () => { void view.dispose(); },
      queueDisplay: createQueueDisplay({ view }) };
  } });
  try {
    const runtime = await manager.createRuntime({ activate: true, autoRestoreLatest: false });
    assert.equal((await manager.selectModel({ runtimeId: runtime.runtimeId, nextModel: secondModel, behavior: "fork" })).outcome, "updated");
    assert.equal((await runtime.view.conversation.agent(context)).model?.modelId, secondModel.id);
    assert.equal((await manager.selectModel({ runtimeId: runtime.runtimeId, nextModel: secondModel, behavior: "fork" })).outcome, "unchanged");
    await harness.commit((tx) => tx.appendEntry(runtime.view.conversation.id, {
      kind: "test.message", model: [{ role: "user", content: "Keep immutable", timestamp: 1 }],
    }), context);
    // Reopen the projection to observe the complete committed timeline.
    await runtime.view.dispose();
    runtime.view = await ConversationPresenter.open(harness, runtime.view.conversation, models, firstModel);
    views.push(runtime.view);
    const result = await manager.selectModel({ runtimeId: runtime.runtimeId, nextModel: firstModel, behavior: "fork" });
    assert.equal(result.outcome, "forked");
    if (result.outcome !== "forked") throw new Error("Expected fork");
    assert.equal((await harness.commit((tx) => tx.conversation(result.runtime.view.conversation.id), context))?.parent?.conversationId,
      runtime.view.conversation.id);
    assert.deepEqual(result.runtime.view.current.messages, runtime.view.current.messages);
    assert.equal(result.runtime.view.current.thinkingLevel, "high");
    assert.equal(result.runtime.view.current.model.id, firstModel.id);
    assert.equal((await runtime.view.conversation.agent(context)).model?.modelId, secondModel.id);
    assert.equal(manager.getActiveRuntime(), result.runtime);
    assert.equal(manager.snapshotTabs()[1]?.title, result.title);
    const source = createAssistantMessageEventStream();
    let started: () => void = () => {};
    const requestStarted = new Promise<void>((resolve) => { started = resolve; });
    models.setProvider({ ...faux.provider, streamSimple: () => { started(); return source; } });
    const running = await runtime.view.conversation.submit({ type: "input", content: "Continue after tab close" }, context);
    await requestStarted;
    const queued = await runtime.view.conversation.submit({ type: "input", content: "Queued" }, context);
    await runtime.view.refresh();
    assert.equal(manager.snapshotTabs()[0]?.isBusy, true);
    assert.equal(runtime.view.current.inbox.length, 1);
    assert.equal((await manager.selectModel({ runtimeId: runtime.runtimeId, nextModel: firstModel, behavior: "fork" })).outcome, "busy");
    manager.closeRuntime(runtime.runtimeId);
    // Closing the UI must not abort either active or queued Durable submissions.
    faux.setResponses([fauxAssistantMessage("Queued completed")]);
    models.setProvider(faux.provider);
    source.push({ type: "done", reason: "stop", message: fauxAssistantMessage("Finished after closing") });
    assert.equal((await running.wait(context)).status, "done");
    assert.equal((await queued.wait(context)).status, "done");
  } finally {
    for (const view of views) await view.dispose();
    await harness.close(context);
  }
});
