import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import { createRegistry, Harness, MemoryStorage, type Conversation } from "@earendil-works/pi-durable";
import { compactionForModel } from "../src/durable/compaction-policy.ts";

async function seed(conversation: Conversation, tokens: number): Promise<void> {
  // Passive entries avoid generation usage estimates obscuring the threshold.
  for (let i = 0; i < 4; i++) {
    await conversation.submit({ type: "write", entry: {
      kind: "test.history", model: [{ role: "user", content: "x".repeat(Math.floor(tokens * 3.5 / 4)), timestamp: i }],
    } }, context);
  }
}

void test("interleaved real Harness conversations use model thresholds and keep summary budget independent", { timeout: 15000 }, async () => {
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: Infinity, models: [
    { id: "small", contextWindow: 32_000, maxTokens: 32_000 },
    { id: "large", contextWindow: 200_000, maxTokens: 32_000 },
  ] });
  const summaries: { model: string; budget: number | undefined; cache: string | undefined }[] = [];
  faux.setResponses(Array.from({ length: 5 }, () => fauxAssistantMessage("summary or answer")));
  models.setProvider({ ...faux.provider, streamSimple: (model, transcript, options) => {
    if (options?.cacheRetention === "none") summaries.push({ model: model.id, budget: options.maxTokens, cache: options.cacheRetention });
    return faux.provider.streamSimple(model, transcript, options);
  } });
  const harness = await Harness.open(new MemoryStorage(), {
    models, registry: createRegistry(), settings: { compactionForModel },
  }, context);
  try {
    const small = await harness.root(context, { agent: { model: { provider: faux.getModel().provider, modelId: "small" } } });
    const large = await harness.createConversation({ ownership: { kind: "root" }, agent: { model: { provider: faux.getModel().provider, modelId: "large" } } }, context);
    await seed(small, 24_000);
    await seed(large, 169_000);
    const run = async (conversation: Conversation): Promise<void> => {
      const submission = await conversation.submit({ type: "input", content: "continue" }, context);
      assert.equal((await submission.wait(context)).status, "done");
    };
    await run(large);
    assert.equal(summaries.length, 0, "below large quality cap must not compact in background");
    await run(small);
    assert.deepEqual(summaries.map((item) => [item.model, item.budget]), [["small", 12_800]]);
    // New large history is above the quality cap but below window-reserve.
    const capped = await harness.createConversation({ ownership: { kind: "root" }, agent: { model: { provider: faux.getModel().provider, modelId: "large" } } }, context);
    await seed(capped, 175_000);
    await run(capped);
    assert.deepEqual(summaries.map((item) => [item.model, item.budget]), [["small", 12_800], ["large", 13_107]]);
  } finally { await harness.close(context); }
});

for (const mode of ["manual", "overflow"] as const) {
  void test(`native ${mode} uses fresh model policy, not harness-wide selection`, { timeout: 15000 }, async () => {
    const models = createModels();
    const faux = fauxProvider({ tokensPerSecond: Infinity, models: [{ id: "small", contextWindow: 32_000, maxTokens: 32_000 }] });
    let summaries = 0;
    models.setProvider({ ...faux.provider, streamSimple: (model, transcript, options) => {
      if (options?.cacheRetention === "none") {
        summaries++;
        assert.equal(options.maxTokens, 12_800);
      }
      return faux.provider.streamSimple(model, transcript, options);
    } });
    // Global policy deliberately forbids overflow and keeps too much history.
    // Overrides must reach generation classification AND CompactionTask.select.
    const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry(), settings: {
      compaction: { enabled: false, keepRecentTokens: 1_000_000 },
      compactionForModel: (model) => ({ ...compactionForModel(model), keepRecentTokens: 4_000, enabled: mode === "overflow" }),
    } }, context);
    try {
      const conversation = await harness.root(context, { agent: { model: { provider: faux.getModel().provider, modelId: "small" } } });
      await seed(conversation, 15_000); // Below blocking threshold, but a selectable prefix.
      if (mode === "manual") {
        faux.setResponses([fauxAssistantMessage("summary")]);
        const task = await conversation.compact(undefined, context);
        assert.equal((await harness.waitForTask(task, context)).state.outcome.status, "completed");
        await conversation.waitForIdle(context);
      } else {
        faux.setResponses([
          fauxAssistantMessage("", { stopReason: "error", errorMessage: "maximum context length is 32000 tokens" }),
          fauxAssistantMessage("summary"),
          fauxAssistantMessage("recovered"),
        ]);
        const submission = await conversation.submit({ type: "input", content: "continue" }, context);
        const receipt = await submission.wait(context);
        assert.equal(receipt.status, "done", JSON.stringify(receipt));
        assert.equal(faux.state.callCount, 3, "native overflow performs one summary and one retry");
      }
      assert.equal(summaries, 1);
    } finally { await harness.close(context); }
  });
}
