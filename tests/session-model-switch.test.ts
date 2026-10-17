import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels, fauxProvider } from "@earendil-works/pi-ai";
import { createRegistry, Harness, MemoryStorage } from "@earendil-works/pi-durable";
import { ConversationPresenter } from "../src/durable/conversation-view.ts";
import { DurableSessions } from "../src/durable/sessions.ts";
import { createQueueDisplay } from "../src/taskpane/queue-display.ts";
import { SessionRuntimeManager } from "../src/taskpane/session-runtime-manager.ts";

for (const behavior of ["inPlace", "fork"] as const) {
  void test(`${behavior} model selection survives native projection reattachment`, async () => {
    const models = createModels();
    const faux = fauxProvider({ models: [{ id: "original" }, { id: "selected" }] });
    models.setProvider(faux.provider);
    const originalModel = faux.getModel("original");
    const selectedModel = faux.getModel("selected");
    const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
    const sessions = new DurableSessions(harness, models, originalModel);
    const views: ConversationPresenter[] = [];
    const manager = new SessionRuntimeManager({ sessions, createRuntime: async (options) => {
      const conversation = options.conversationId ? await sessions.resolve(options.conversationId)
        : await sessions.create(originalModel, "off");
      assert.ok(conversation);
      const view = await ConversationPresenter.open(harness, conversation, models, originalModel);
      views.push(view);
      return { runtimeId: `tab-${views.length}`, view, tools: [], queueDisplay: createQueueDisplay({ view }),
        title: "", hasExplicitTitle: false, lockState: "idle", refreshCapabilities: async () => {},
        dispose: () => { void view.dispose(); } };
    } });
    try {
      const conversation = await sessions.create(originalModel, "off");
      await harness.commit((tx) => tx.appendEntry(conversation.id, {
        kind: "test.message", model: [{ role: "user", content: "Original question", timestamp: 1 }],
      }), context);
      const original = await manager.createRuntime({ activate: true, autoRestoreLatest: false, conversationId: String(conversation.id) });
      const result = await manager.selectModel({ runtimeId: original.runtimeId, nextModel: selectedModel, behavior });
      assert.equal(result.outcome, behavior === "fork" ? "forked" : "updated");
      const selected = result.outcome === "forked" ? result.runtime : original;
      const id = String(selected.view.conversation.id);
      if (behavior === "fork") assert.notEqual(id, String(conversation.id));
      else assert.equal(id, String(conversation.id));
      await selected.view.dispose();
      const reopened = await manager.createRuntime({ activate: true, autoRestoreLatest: false, conversationId: id });
      assert.equal(reopened.view.current.model.id, selectedModel.id);
      assert.match(JSON.stringify(reopened.view.current.messages), /Original question/);
      assert.equal((await sessions.loadSession(String(conversation.id)))?.model.id,
        behavior === "fork" ? originalModel.id : selectedModel.id);
      const refreshedMetadata = { ...selectedModel, reasoning: !selectedModel.reasoning };
      assert.equal((await manager.selectModel({ runtimeId: reopened.runtimeId, nextModel: refreshedMetadata, behavior: "fork" })).outcome, "updated");
      assert.equal(String(reopened.view.conversation.id), id, "metadata refresh never forks");
    } finally {
      for (const view of views) await view.dispose();
      await harness.close(context);
    }
  });
}
