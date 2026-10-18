import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createAssistantMessageEventStream } from "@earendil-works/pi-ai";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import { createRegistry, Harness, InboxDoc, MemoryStorage } from "@earendil-works/pi-durable";
import { editQueuedInput, takeQueuedInputs } from "../src/durable/inbox.ts";

void test("queued edits retain submission identity and position; withdrawal settles only queued inputs atomically", async () => {
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: Infinity });
  const source = createAssistantMessageEventStream();
  let first = true;
  models.setProvider({ ...faux.provider, streamSimple: (model, transcript, options) => {
    if (first) { first = false; return source; }
    return faux.provider.streamSimple(model, transcript, options);
  } });
  const model = faux.getModel();
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const conversation = await harness.root(context, { agent: { model: { provider: model.provider, modelId: model.id } } });
  try {
    const running = await conversation.submit({ type: "input", content: "running" }, context);
    const one = await conversation.submit({ type: "input", content: "one", requestId: "one" }, context);
    const two = await conversation.submit({ type: "input", content: "two", requestId: "two" }, context);
    const three = await conversation.submit({ type: "input", content: "three", whenBusy: "steer" }, context);
    const note = await conversation.submit({ type: "write", entry: {
      kind: "test.context", model: [{ role: "user", content: "passive", timestamp: 1 }],
    } }, context);
    assert.equal(await editQueuedInput(conversation, two.id, [{ type: "text", text: "edited two" }], context), true);
    const inbox = await harness.snapshot(InboxDoc, conversation.id, context);
    assert.deepEqual(inbox?.items.map((item) => [item.id, item.mode]), [
      [one.id, "followUp"], [two.id, "followUp"], [three.id, "steer"], [note.id, "write"],
    ]);
    assert.equal((await conversation.submit({ type: "input", content: "duplicate", requestId: "two" }, context)).id, two.id);
    assert.equal(await editQueuedInput(conversation, running.id, "must not replace active input", context), false);
    const withdrawn = await takeQueuedInputs(conversation, context);
    assert.deepEqual(withdrawn.map((item) => item.content), ["one", [{ type: "text", text: "edited two" }], "three"]);
    assert.deepEqual((await harness.snapshot(InboxDoc, conversation.id, context))?.items.map((item) => item.id), [note.id]);
    for (const submission of [one, two, three]) {
      const receipt = await submission.wait(context);
      assert.equal(receipt.status, "unanswered");
      if (receipt.status === "unanswered") assert.equal(receipt.reason, "aborted");
    }
    faux.setResponses([fauxAssistantMessage("Final")]);
    source.push({ type: "done", reason: "stop", message: fauxAssistantMessage("Initial") });
    await conversation.waitForIdle(context);
    assert.equal((await running.wait(context)).status, "done");
    assert.equal((await note.wait(context)).status, "done");
    assert.doesNotMatch(JSON.stringify((await conversation.context(context)).messages), /edited two|must not replace/);
  } finally {
    source.end(fauxAssistantMessage("Stopped"));
    await harness.close(context);
  }
});
