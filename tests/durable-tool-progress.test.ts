import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import type { AgentEvent, AgentTool, AgentToolUpdateCallback } from "@earendil-works/pi-agent-core";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxProvider, fauxAssistantMessage, fauxToolCall } from "@earendil-works/pi-ai/providers/faux";
import { MemoryStorage, watchEvents } from "@earendil-works/pi-durable";
import { openDurableRuntime } from "../src/durable/runtime.ts";
import { DurableSessions } from "../src/durable/sessions.ts";
import { legacyEventProjection } from "../src/commands/extension-api-compat.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import { withWorkbookCoordinator } from "../src/tools/with-workbook-coordinator.ts";

function latch() {
  let release = (): void => {};
  const promise = new Promise<void>((resolve) => { release = resolve; });
  return { promise, release };
}

for (const toolName of ["read_range", "extension_progress"]) {
  void test(`native ${toolName} progress preserves partials and drains before the exact final receipt`, { timeout: 10000 }, async () => {
    const models = createModels();
    const faux = fauxProvider({ tokensPerSecond: Infinity });
    models.setProvider(faux.provider);
    const model = faux.getModel();
    faux.setResponses([fauxAssistantMessage(fauxToolCall(toolName, { range: "B2" }, { id: "same-provider-id" }), { stopReason: "toolUse" }), fauxAssistantMessage("done")]);
    const coordinator = createWorkbookCoordinator();
    const runtime = await openDurableRuntime({ storage: new MemoryStorage(), models, workbookId: "book", coordinator,
      autoCompaction: false, mutationHost: {
        prepare: () => Promise.resolve({ intent: {}, target: {}, before: { workbookId: "book", structure: {}, state: {}, complete: false, limitation: "opaque" }, expected: null }),
        observe: () => Promise.resolve({ workbookId: "book", structure: {}, state: {}, complete: false, limitation: "opaque" }),
      } });
    const firstVisible = latch();
    const ended = latch();
    const events: AgentEvent[] = [];
    let late: AgentToolUpdateCallback | undefined;
    const final = { content: [{ type: "text" as const, text: "final only" }],
      details: toolName === "read_range" ? undefined : { completed: true } };
    const tool: AgentTool = { name: toolName, label: "Progress", description: "Progress", parameters: Type.Object({ range: Type.String() }),
      execute: async (_id, _args, _signal, onUpdate) => {
        late = onUpdate;
        const partial = { content: [{ type: "text" as const, text: "first snapshot" },
          { type: "image" as const, data: "image-data", mimeType: "image/png" }], details: { stage: 1 } };
        onUpdate?.(partial);
        partial.details.stage = 999;
        await firstVisible.promise;
        onUpdate?.({ content: [{ type: "text", text: "last snapshot" }], details: { stage: 2 } });
        return final;
      } };
    const conversation = await runtime.harness.root(context, { agent: { model: { provider: model.provider, modelId: model.id } } });
    await runtime.install(conversation, { tools: [tool], systemPrompt: "", bindMutationGate: (gate, id) =>
      withWorkbookCoordinator([tool], coordinator, { getWorkbookId: () => Promise.resolve("book"), getSessionId: () => String(conversation.id) },
        undefined, { mutationGate: gate, getMutationId: () => id }) }, (messages) => Promise.resolve(messages));
    const stream = await watchEvents(runtime.harness, conversation.id, context);
    const project = legacyEventProjection();
    stream.start((batch) => {
      for (const native of batch) for (const event of project(native)) {
        events.push(event);
        if (event.type === "tool_execution_update") firstVisible.release();
        if (event.type === "tool_execution_end") ended.release();
      }
      return Promise.resolve();
    });
    try {
      const submission = await conversation.submit({ type: "input", content: "Go" }, context);
      assert.equal((await submission.wait(context)).status, "done");
      await ended.promise;
      const updates = events.filter((event) => event.type === "tool_execution_update");
      assert.equal(updates.length, 2);
      assert.deepEqual(updates.map((event) => [event.toolCallId, event.toolName]), [["same-provider-id", toolName], ["same-provider-id", toolName]]);
      for (const update of updates) assert.deepEqual(update.args, { range: "B2" });
      assert.deepEqual([updates[0]?.partialResult, updates[1]?.partialResult], [
        { content: [{ type: "text", text: "first snapshot" }, { type: "image", data: "image-data", mimeType: "image/png" }], details: { stage: 1 } },
        { content: [{ type: "text", text: "last snapshot" }], details: { stage: 2 } },
      ]);
      const transcript = await conversation.context(context);
      const receipt = transcript.messages.find((message) => message.role === "toolResult");
      assert.ok(receipt?.role === "toolResult");
      assert.deepEqual(receipt.content, final.content);
      assert.deepEqual(receipt.details, final.details);
      assert.doesNotMatch(JSON.stringify(receipt), /snapshot|agent-tool-progress/);
      const catalogue = new DurableSessions(runtime.harness, models, model);
      try {
        const exported = await catalogue.loadSession(String(conversation.id));
        assert.ok(exported);
        const exportedReceipt = exported.messages.find((message) => message.role === "toolResult");
        assert.ok(exportedReceipt?.role === "toolResult");
        assert.deepEqual(exportedReceipt.content, final.content);
        assert.deepEqual(exportedReceipt.details, final.details);
        assert.doesNotMatch(JSON.stringify(exported), /snapshot|agent-tool-progress|image-data/);
      } finally { catalogue.dispose(); }
      assert.ok(events.findIndex((event) => event === updates[1]) < events.findIndex((event) => event.type === "tool_execution_end"));
      late?.({ content: [{ type: "text", text: "stale-after-final" }], details: {} });
      await conversation.commit(() => {}, context);
      assert.doesNotMatch(JSON.stringify(await conversation.context(context)), /stale-after-final/);
      assert.equal(events.filter((event) => event.type === "tool_execution_update").length, 2);
    } finally { firstVisible.release(); await stream.stop(); await runtime.harness.close(context); }
  });
}

void test("cancelled tool callbacks cannot publish into a subsequent conversation using the same call ID", { timeout: 10000 }, async () => {
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: Infinity });
  models.setProvider(faux.provider);
  const model = faux.getModel();
  const writeCall = () => fauxAssistantMessage(fauxToolCall("read_range", {}, { id: "reused" }), { stopReason: "toolUse" });
  faux.setResponses([writeCall(), writeCall(), fauxAssistantMessage("done")]);
  const runtime = await openDurableRuntime({ storage: new MemoryStorage(), models, workbookId: "book", coordinator: createWorkbookCoordinator(), autoCompaction: false,
    mutationHost: { prepare: () => Promise.reject(new Error("read only")), observe: () => Promise.reject(new Error("read only")) } });
  const visible = latch();
  const aborted = latch();
  let oldCallback: AgentToolUpdateCallback | undefined;
  let calls = 0;
  const tool: AgentTool = { name: "read_range", label: "Read", description: "Read", parameters: Type.Object({}),
    execute: async (_id, _args, signal, update) => {
      if (++calls === 1) {
        oldCallback = update;
        update?.({ content: [{ type: "text", text: "before cancellation" }], details: {} });
        await new Promise<void>((resolve) => signal?.addEventListener("abort", () => {
          update?.({ content: [{ type: "text", text: "stale-on-abort" }], details: {} });
          resolve();
        }, { once: true }));
      } else {
        oldCallback?.({ content: [{ type: "text", text: "stale-cross-conversation" }], details: {} });
        update?.({ content: [{ type: "text", text: "second progress" }], details: { second: true } });
      }
      return { content: [{ type: "text", text: "final" }], details: {} };
    } };
  const first = await runtime.harness.root(context, { agent: { model: { provider: model.provider, modelId: model.id } } });
  const second = await runtime.harness.createConversation({ ownership: { kind: "ownerless" }, agent: { model: { provider: model.provider, modelId: model.id } } }, context);
  const events: AgentEvent[] = [];
  const streams = [];
  for (const conversation of [first, second]) {
    await runtime.install(conversation, { tools: [tool], systemPrompt: "" }, (messages) => Promise.resolve(messages));
    const stream = await watchEvents(runtime.harness, conversation.id, context);
    const project = legacyEventProjection();
    stream.start((batch) => {
      for (const native of batch) for (const event of project(native)) {
        events.push(event);
        if (event.type === "tool_execution_update") visible.release();
        if (event.type === "tool_execution_end") aborted.release();
      }
      return Promise.resolve();
    });
    streams.push(stream);
  }
  try {
    const original = await first.submit({ type: "input", content: "cancel me" }, context);
    await visible.promise;
    await first.abort(context);
    assert.equal((await original.wait(context)).status, "unanswered");
    await aborted.promise;
    const next = await second.submit({ type: "input", content: "next" }, context);
    assert.equal((await next.wait(context)).status, "done");
    for (const stream of streams) await stream.stop();
    assert.doesNotMatch(JSON.stringify(events), /stale-/);
    const updates = events.filter((event) => event.type === "tool_execution_update");
    assert.equal(updates.length, 2);
    assert.match(JSON.stringify(updates[1]), /second progress/);
    assert.doesNotMatch(JSON.stringify((await second.context(context)).messages), /agent-tool-progress|before cancellation|stale-/);
  } finally { for (const stream of streams) await stream.stop(); await runtime.harness.close(context); }
});
