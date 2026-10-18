import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider, fauxToolCall } from "@earendil-works/pi-ai/providers/faux";
import { MemoryStorage } from "@earendil-works/pi-durable";
import type { AgentTool } from "@earendil-works/pi-agent-core";
import { MutationDoc, openDurableRuntime } from "../src/durable/runtime.ts";
import { parseMutationRecord, type MutationObservation } from "../src/durable/mutations.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import { withWorkbookCoordinator } from "../src/tools/with-workbook-coordinator.ts";

function latch() {
  let release = (): void => {};
  const promise = new Promise<void>((resolve) => { release = resolve; });
  return { promise, release };
}

void test("shared coordinator waits for exact task acknowledgement across conversations and reused raw IDs", { timeout: 15000 }, async () => {
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: Infinity });
  const secondRequest = latch();
  let requests = 0;
  models.setProvider({ ...faux.provider, streamSimple: (...args) => {
    requests += 1;
    const stream = faux.provider.streamSimple(...args);
    if (requests === 2) secondRequest.release();
    return stream;
  } });
  const model = faux.getModel();
  // Every turn and conversation deliberately receives the same provider ID.
  const write = () => fauxAssistantMessage({ ...fauxToolCall("write_cells", {}), id: "call_1" }, { stopReason: "toolUse" });
  faux.setResponses([write(), write(), fauxAssistantMessage("first done"), fauxAssistantMessage("second done"),
    write(), fauxAssistantMessage("third done")]);
  const coordinator = createWorkbookCoordinator();
  let value = 0;
  const observe = (): MutationObservation => ({ workbookId: "book", structure: { sheetId: "sheet" },
    state: value, complete: true, limitation: null });
  const rawIds: string[] = [];
  const tool: AgentTool = { name: "write_cells", label: "Write", description: "Write a cell", parameters: Type.Object({}),
    execute: (id) => {
      rawIds.push(id);
      value += 1;
      return Promise.resolve({ content: [{ type: "text" as const, text: `written ${value}` }], details: {} });
    } };
  const runtime = await openDurableRuntime({ storage: new MemoryStorage(), models, workbookId: "book",
    coordinator, autoCompaction: false, mutationHost: {
      prepare: (call) => Promise.resolve({ intent: call.params, target: { sheetId: "sheet" }, before: observe(),
        expected: { ...observe(), state: value + 1 } }),
      observe: () => Promise.resolve(observe()),
    } });
  const firstReturned = latch();
  const releaseFirst = latch();
  const secondReturned = latch();
  const releaseSecond = latch();
  try {
    const first = await runtime.harness.root(context, { agent: { model: { provider: model.provider, modelId: model.id } } });
    const second = await runtime.harness.createConversation({ ownership: { kind: "ownerless" },
      agent: { model: { provider: model.provider, modelId: model.id } } }, context);
    for (const conversation of [first, second]) {
      await runtime.install(conversation, { tools: [tool], systemPrompt: "Write once then finish.",
        bindMutationGate: (gate, mutationId) => withWorkbookCoordinator([tool], coordinator, {
          getWorkbookId: () => Promise.resolve("book"), getSessionId: () => String(conversation.id),
        }, undefined, { mutationGate: gate, getMutationId: () => mutationId }).map((wrapped) => ({ ...wrapped,
          execute: async (...args: Parameters<typeof wrapped.execute>) => {
            const result = await wrapped.execute(...args);
            // The real gate has returned and released the coordinator, but the
            // adapter has not yet returned this result to its Durable ToolTask.
            if (rawIds.length === 1) { firstReturned.release(); await releaseFirst.promise; }
            else if (rawIds.length === 2) { secondReturned.release(); await releaseSecond.promise; }
            return result;
          },
        })),
      }, (messages) => Promise.resolve(messages));
    }
    const a = await first.submit({ type: "input", content: "First write" }, context);
    await Promise.race([firstReturned.promise, a.wait(context).then(async () => {
      assert.fail(`First write never reached its returned boundary: ${JSON.stringify((await first.context(context)).messages)}`);
    })]);
    const b = await second.submit({ type: "input", content: "Second write" }, context);
    await secondRequest.promise;
    await runtime.acknowledgeCommittedResults();
    const pendingFirst = await runtime.mutations.pending();
    assert.equal(pendingFirst.length, 1);
    assert.equal(pendingFirst[0]?.phase, "returned");
    assert.deepEqual(rawIds, ["call_1"]);
    releaseFirst.release();
    await secondReturned.promise;
    // The first successful receipt has the same raw ID, but must not acknowledge
    // the still-uncommitted second task's returned record.
    await runtime.acknowledgeCommittedResults();
    const pendingSecond = await runtime.mutations.pending();
    assert.equal(pendingSecond.length, 1);
    const record = pendingSecond[0];
    assert.ok(record);
    assert.equal(record.phase, "returned");
    assert.notEqual(record.mutationId, pendingFirst[0]?.mutationId);
    assert.equal(record.toolCallId, pendingFirst[0]?.toolCallId);
    const review = await runtime.mutations.reconcile(record.mutationId);
    await runtime.mutations.resolve(record.mutationId, review.observationToken);
    releaseSecond.release();
    assert.equal((await a.wait(context)).status, "done");
    assert.equal((await b.wait(context)).status, "done");
    await runtime.acknowledgeCommittedResults();
    const resolved = (await runtime.harness.snapshot(MutationDoc, record.mutationId, context))?.record;
    assert.ok(resolved);
    assert.equal(parseMutationRecord(resolved).phase, "resolved", "late acknowledgement must preserve explicit resolution");
    const c = await first.submit({ type: "input", content: "Another turn" }, context);
    assert.equal((await c.wait(context)).status, "done");
    await runtime.acknowledgeCommittedResults();
    const tasks = await runtime.harness.commit((tx) => tx.scanTasks({}, 100), context);
    const docs = await Promise.all(tasks.items.filter((task) => task.kind === "pi.tool").map((task) => runtime.harness.snapshot(MutationDoc, `task:${task.id}`, context)));
    const records = docs.flatMap((doc) => doc?.record ? [parseMutationRecord(doc.record)] : []);
    assert.equal(records.length, 3);
    assert.equal(new Set(records.map((item) => item.mutationId)).size, 3);
    assert.equal(records.filter((item) => item.phase === "resultCommitted").length, 2);
    const bindings = docs.flatMap((doc) => doc?.binding ? [[`task:${doc.binding.taskId}`, doc.binding] as const] : []);
    assert.equal(bindings.length, 3);
    for (const [key, binding] of bindings) {
      assert.equal(key, `task:${binding.taskId}`);
      assert.equal(binding.rawCallId, "call_1");
    }
    assert.equal(new Set(bindings.map(([, binding]) => binding.conversationId)).size, 2);
    assert.deepEqual(rawIds, ["call_1", "call_1", "call_1"]);
    for (const conversation of [first, second]) {
      const transcript = await conversation.context(context);
      const results = transcript.messages.filter((message) => message.role === "toolResult");
      assert.equal(results.length, conversation.id === first.id ? 2 : 1);
      for (const result of results) {
        assert.equal(result.toolCallId, "call_1");
        assert.equal(result.isError, false);
      }
    }
    assert.equal(value, 3);
    assert.deepEqual(await runtime.mutations.pending(), []);
  } finally {
    releaseFirst.release();
    releaseSecond.release();
    await runtime.harness.close(context);
  }
});
