import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createAssistantMessageEventStream } from "@earendil-works/pi-ai";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import { MemoryStorage } from "@earendil-works/pi-durable";
import { createModelRecoveryLifecycle } from "../src/durable/model-recovery.ts";
import { openDurableRuntime, routeModels } from "../src/durable/runtime.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import type { MutationHost } from "../src/durable/mutations.ts";

const mutationHost: MutationHost = {
  prepare: () => Promise.reject(new Error("Unexpected workbook mutation")),
  observe: () => Promise.reject(new Error("Unexpected workbook observation")),
};

void test("native cancellation without any lifecycle event persists a resumable retry instead of a final failed submission", { timeout: 15000 }, async () => {
  // A disposable backend kept across two Harness instances. Browser and Excel
  // acceptance cover actual IndexedDB; this test owns the scheduling contract.
  class ReopenableMemory extends MemoryStorage {
    override async close(): Promise<void> {}
  }
  const storage = new ReopenableMemory();
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: Infinity });
  models.setProvider(faux.provider);
  const model = faux.getModel();
  const target = new EventTarget();
  const lifecycle = createModelRecoveryLifecycle(() => Promise.resolve(), target);
  const source = createAssistantMessageEventStream();
  const runtime = await openDurableRuntime({
    storage, workbookId: "native-recovery", mutationHost, coordinator: createWorkbookCoordinator(), autoCompaction: false,
    models: routeModels(models, () => source, () => lifecycle),
  });
  const conversation = await runtime.harness.root(context, { agent: { model: { provider: model.provider, modelId: model.id } } });
  await runtime.install(conversation, { tools: [], systemPrompt: "Test request recovery." }, (messages) => Promise.resolve(messages));
  let published: (() => void) | undefined;
  const networkErrorCommitted = new Promise<void>((resolve) => { published = resolve; });
  const unsubscribe = runtime.harness.subscribeCommits((publication) => {
    if (publication.changes.some((write) => write.type === "entry"
      && write.value.model?.some((message) => message.role === "assistant"
        && message.errorMessage?.includes("Load failed")))) published?.();
  });
  let submissionId;
  try {
    const submitted = await conversation.submit({
      type: "input", content: "Finish this request after the pane returns.", requestId: "native-once",
    }, context);
    submissionId = submitted.id;
    source.push({ type: "error", reason: "error", error: {
      ...fauxAssistantMessage(""), stopReason: "error", errorMessage: "Load failed",
    } });
    await networkErrorCommitted;
    // No beforeunload/pagehide/explicit-reload event was delivered. The native
    // host disappears only after its cancelled fetch error has been committed.
    await runtime.harness.close(context);
    const persisted = await storage.submission(submitted.id, context);
    assert.equal(persisted?.status, "placed", "the original input must remain unanswered, not terminally failed");
    assert.equal(lifecycle.suspending, false);
  } finally {
    unsubscribe();
    lifecycle.dispose();
    await runtime.harness.close(context);
  }
  assert.ok(submissionId);
  faux.setResponses([fauxAssistantMessage("Recovered after native pane closure.")]);
  const reopened = await openDurableRuntime({
    storage, models, workbookId: "native-recovery", mutationHost, coordinator: createWorkbookCoordinator(), autoCompaction: false,
  });
  const restored = await reopened.harness.root(context);
  await reopened.install(restored, { tools: [], systemPrompt: "Test request recovery." }, (messages) => Promise.resolve(messages));
  try {
    assert.equal(restored.id, conversation.id);
    const original = await reopened.harness.submission(submissionId, context);
    assert.ok(original);
    assert.equal((await original.wait(context)).status, "done");
    const entries = (await restored.entries({}, 100, undefined, context)).items;
    assert.equal(entries.filter((entry) => entry.kind === "pi.user").length, 1);
    assert.match(JSON.stringify(entries), /Recovered after native pane closure/);
    assert.match(JSON.stringify(entries), /Network error: Load failed/);
  } finally { await reopened.harness.close(context); }
});
