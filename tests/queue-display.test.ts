import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createAssistantMessageEventStream } from "@earendil-works/pi-ai";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import { createRegistry, Harness, InboxDoc, MemoryStorage } from "@earendil-works/pi-durable";
import { ConversationPresenter } from "../src/durable/conversation-view.ts";
import { createQueueDisplay } from "../src/taskpane/queue-display.ts";
import { restoreQueuedMessagesToEditor } from "../src/taskpane/keyboard-shortcuts.ts";

void test("queue display edits and restores the native inbox without shadow queue state", async () => {
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: Infinity });
  const source = createAssistantMessageEventStream();
  models.setProvider({ ...faux.provider, streamSimple: () => source });
  const model = faux.getModel();
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const conversation = await harness.root(context, { agent: { model: { provider: model.provider, modelId: model.id } } });
  const view = await ConversationPresenter.open(harness, conversation, models, model);
  const display = createQueueDisplay({ view });
  try {
    const running = await conversation.submit({ type: "input", content: "running" }, context);
    const first = await conversation.submit({ type: "input", content: "one", whenBusy: "steer" }, context);
    const second = await conversation.submit({ type: "input", content: "two", whenBusy: "followUp" }, context);
    await conversation.submit({ type: "write", entry: { kind: "test.write", model: [{ role: "user", content: "note", timestamp: 1 }] } }, context);
    assert.equal(await display.editQueuedInput(second.id, "edited two"), true);
    assert.equal(await display.editQueuedInput(running.id, "cannot edit placed"), false);
    const taken = await display.takeQueuedInputs();
    assert.deepEqual(taken.map((item) => [item.id, item.mode, item.content]), [
      [first.id, "steer", "one"], [second.id, "followUp", "edited two"],
    ]);
    assert.deepEqual(await display.takeQueuedInputs(), []);
    assert.deepEqual((await harness.snapshot(InboxDoc, conversation.id, context))?.items.map((item) => item.mode), ["write"]);
    assert.equal((await second.wait(context)).status, "unanswered");
    assert.equal(await display.editQueuedInput(second.id, "cannot resurrect"), false);
    await conversation.submit({ type: "input", content: "three", whenBusy: "followUp" }, context);
    await conversation.submit({ type: "input", content: "four", whenBusy: "steer" }, context);
    const input = { value: "draft" };
    assert.equal(await restoreQueuedMessagesToEditor({ getInput: () => input, getTextarea: () => undefined }, display), 2);
    assert.equal(input.value, "three\n\nfour\n\ndraft");
  } finally {
    display.dispose();
    await view.dispose();
    source.end(fauxAssistantMessage("Stopped"));
    await harness.close(context);
  }
});
