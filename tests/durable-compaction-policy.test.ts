import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createAssistantMessageEventStream } from "@earendil-works/pi-ai";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import { createRegistry, Harness, LiveDoc, MemoryStorage, type Conversation } from "@earendil-works/pi-durable";
import { compactionForModel } from "../src/durable/compaction-policy.ts";
import { routeModels } from "../src/durable/runtime.ts";
import { createOfficeStreamFn } from "../src/auth/stream-proxy.ts";
import { installFakeDom } from "./fixtures/fake-dom.ts";

async function seed(conversation: Conversation, tokens: number): Promise<void> {
  // Passive entries avoid generation usage estimates obscuring the threshold.
  for (let i = 0; i < 4; i++) {
    await conversation.submit({ type: "write", entry: {
      kind: "test.history", model: [{ role: "user", content: "x".repeat(Math.floor(tokens * 3.5 / 4)), timestamp: i }],
    } }, context);
  }
}

for (const providerPreference of [undefined, "long"] as const) {
  for (const preference of [undefined, "none", "short", "long"] as const) {
    void test(`summary retention through Office transport: request=${preference ?? "unset"}, provider=${providerPreference ?? "unset"}`, async () => {
      const dom = installFakeDom();
      const models = createModels();
      const faux = fauxProvider({ tokensPerSecond: Infinity });
      let retention: string | undefined;
      models.setProvider({ ...faux.provider,
        auth: { apiKey: { name: "Faux", resolve: () => Promise.resolve({ auth: {},
          ...(providerPreference ? { env: { PI_CACHE_RETENTION: providerPreference } } : {}),
        }) } },
        streamSimple: (model, transcript, options) => {
          retention = options?.cacheRetention;
          assert.ok(transcript.messages.some((message) => message.role === "system" && message.content.includes("context summarization assistant")));
          return faux.provider.streamSimple(model, transcript, options);
        },
      });
      faux.setResponses([fauxAssistantMessage("summary")]);
      const office = createOfficeStreamFn(() => Promise.resolve(undefined), models);
      let transportError = "";
      const routed = routeModels(models, async (model, transcript, options) => {
        try { return await office(model, transcript, options); }
        catch (error) { transportError = String(error); throw error; }
      });
      const harness = await Harness.open(new MemoryStorage(), { models: routed, registry: createRegistry(), settings: {
        stream: preference ? { cacheRetention: preference } : {},
        compaction: { keepRecentTokens: 100 },
      } }, context);
      try {
        const conversation = await harness.root(context, { agent: { model: { provider: faux.getModel().provider, modelId: faux.getModel().id } } });
        await seed(conversation, 4_000);
        const task = await conversation.compact(undefined, context);
        const outcome = (await harness.waitForTask(task, context)).state.outcome;
        assert.equal(outcome.status, "completed", `${JSON.stringify(outcome)} ${transportError}`);
        assert.equal(retention, preference ?? providerPreference ?? "none");
      } finally { await harness.close(context); dom.restore(); }
    });
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

for (const ending of ["second-overflow", "cancel-summary"] as const) {
  void test(`native overflow settles inputs and accepts the next turn after ${ending}`, { timeout: 15000 }, async () => {
    const models = createModels();
    const faux = fauxProvider({ tokensPerSecond: Infinity, models: [{ id: "small", contextWindow: 32_000, maxTokens: 32_000 }] });
    const firstError = "maximum context length is 32000 tokens (initial)";
    const secondError = "maximum context length is 32000 tokens (retry still too large)";
    let summaries = 0;
    let requests = 0;
    let notifySummaryStarted = (): void => {};
    const summaryStarted = new Promise<void>((resolve) => { notifySummaryStarted = resolve; });
    const blockedSummary = createAssistantMessageEventStream();
    models.setProvider({ ...faux.provider, streamSimple: (model, transcript, options) => {
      if (options?.cacheRetention === "none") {
        summaries++;
        assert.equal(options.maxTokens, 12_800);
        notifySummaryStarted();
        if (ending === "cancel-summary") {
          options.signal?.addEventListener("abort", () => blockedSummary.end(fauxAssistantMessage("", { stopReason: "aborted" })), { once: true });
          return blockedSummary;
        }
      } else requests++;
      return faux.provider.streamSimple(model, transcript, options);
    } });
    const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry(), settings: {
      compaction: { enabled: false, keepRecentTokens: 1_000_000 },
      compactionForModel: (model) => ({ ...compactionForModel(model), enabled: true, keepRecentTokens: 4_000 }),
    } }, context);
    try {
      const conversation = await harness.root(context, { agent: { model: { provider: faux.getModel().provider, modelId: "small" } } });
      await seed(conversation, 15_000);
      faux.setResponses([
        fauxAssistantMessage("", { stopReason: "error", errorMessage: firstError }),
        ...(ending === "second-overflow" ? [
          fauxAssistantMessage("summary"),
          fauxAssistantMessage("", { stopReason: "error", errorMessage: secondError }),
        ] : []),
      ]);
      const input = await conversation.submit({ type: "input", content: "continue" }, context);
      if (ending === "cancel-summary") {
        await summaryStarted;
        await conversation.abort(context);
      }
      const receipt = await input.wait(context);
      assert.equal(receipt.status, "unanswered", JSON.stringify(receipt));
      if (receipt.status === "unanswered") {
        assert.equal(receipt.reason, ending === "second-overflow" ? "model_error" : "aborted");
        if (ending === "second-overflow") assert.equal(receipt.detail, secondError);
      }
      await conversation.waitForIdle(context);
      assert.equal(summaries, 1, "one overflow compaction only, even if retry overflows");
      assert.equal(requests, ending === "second-overflow" ? 2 : 1);
      const live = await harness.snapshot(LiveDoc, conversation.id, context);
      assert.equal(live?.run, undefined);
      assert.equal(live?.generation, undefined);
      assert.equal(live?.compactions, undefined, "cancelled blocking child must not leave live compaction status");
      if (ending === "second-overflow") {
        const view = await conversation.context(context);
        // Native context deliberately excludes error/aborted messages from model
        // requests; the raw transcript must still retain the visible failure.
        assert.ok(view.entries.flatMap((entry) => entry.model ?? []).some((message) => message.role === "assistant" && message.stopReason === "error" && message.errorMessage === secondError), "second overflow remains visible in native transcript");
        assert.ok(!view.messages.some((message) => message.role === "assistant" && message.stopReason === "error"));
      }
      faux.setResponses([fauxAssistantMessage("next turn accepted")]);
      const next = await conversation.submit({ type: "input", content: "try a new turn" }, context);
      assert.equal((await next.wait(context)).status, "done");
      await conversation.waitForIdle(context);
      assert.equal(summaries, 1);
      assert.equal(requests, ending === "second-overflow" ? 3 : 2);
    } finally {
      blockedSummary.end(fauxAssistantMessage("", { stopReason: "aborted" }));
      await harness.close(context);
    }
  });
}
