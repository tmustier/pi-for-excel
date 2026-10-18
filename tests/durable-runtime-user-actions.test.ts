import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider, fauxToolCall } from "@earendil-works/pi-ai/providers/faux";
import { MemoryStorage } from "@earendil-works/pi-durable";
import { DirectToolResultEntry, MutationDoc, openDurableRuntime } from "../src/durable/runtime.ts";
import { parseMutationRecord } from "../src/durable/mutations.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import { withWorkbookCoordinator } from "../src/tools/with-workbook-coordinator.ts";
import type { AgentTool } from "@earendil-works/pi-agent-core";

function latch() {
  let release = (): void => {};
  const promise = new Promise<void>((resolve) => { release = resolve; });
  return { promise, release };
}

for (const directFirst of [true, false]) {
  void test(`${directFirst ? "direct" : "model"} write holds admission until its receipt commits before another conversation writes`, { timeout: 15000 }, async () => {
    const models = createModels();
    const faux = fauxProvider({ tokensPerSecond: Infinity });
    models.setProvider(faux.provider);
    const model = faux.getModel();
    faux.setResponses([
      fauxAssistantMessage({ ...fauxToolCall("write_cells", {}), id: "call_1" }, { stopReason: "toolUse" }),
      fauxAssistantMessage("done"),
    ]);
    const coordinator = createWorkbookCoordinator();
    let value = 0;
    const rawIds: string[] = [];
    const observe = () => ({ workbookId: "book", structure: {}, state: value, complete: true, limitation: null });
    const tool: AgentTool = { name: "write_cells", label: "Write", description: "Write", parameters: Type.Object({}),
      execute: (id) => {
        rawIds.push(id);
        value++;
        return Promise.resolve({ content: [{ type: "text" as const, text: `written ${value}` }], details: {} });
      } };
    const runtime = await openDurableRuntime({ storage: new MemoryStorage(), models, workbookId: "book",
      coordinator, autoCompaction: false, mutationHost: {
        prepare: (call) => Promise.resolve({ intent: call.params, target: {}, before: observe(), expected: { ...observe(), state: value + 1 } }),
        observe: () => Promise.resolve(observe()),
      } });
    const returned = [latch(), latch()] as const;
    const releases = [latch(), latch()] as const;
    try {
      const direct = await runtime.harness.root(context);
      const modelConversation = await runtime.harness.createConversation({ ownership: { kind: "ownerless" },
        agent: { model: { provider: model.provider, modelId: model.id } } }, context);
      for (const conversation of [direct, modelConversation]) {
        await runtime.install(conversation, { tools: [tool], systemPrompt: "Write once then finish.",
          bindMutationGate: (gate, id) => withWorkbookCoordinator([tool], coordinator, {
            getWorkbookId: () => Promise.resolve("book"), getSessionId: () => String(conversation.id),
          }, undefined, { mutationGate: gate, getMutationId: () => id }).map((wrapped) => ({ ...wrapped,
            execute: async (...args: Parameters<typeof wrapped.execute>) => {
              const result = await wrapped.execute(...args);
              // Pause outside the real gate/lock, before the Durable task can commit its receipt.
              const index = rawIds.length - 1;
              returned[index]?.release();
              await releases[index]?.promise;
              return result;
            },
          })),
        }, (messages) => Promise.resolve(messages));
      }
      const runModel = async () => {
        const turn = await modelConversation.submit({ type: "input", content: "Write once" }, context);
        assert.equal((await turn.wait(context)).status, "done");
      };
      const runDirect = () => runtime.executeTool(direct, "write_cells", {});
      const first = directFirst ? runDirect() : runModel();
      await returned[0].promise;
      const second = directFirst ? runModel() : runDirect();
      // Observe both actual tasks before yielding for dispatch; a model request
      // alone would not establish that a competing tool write was scheduled.
      while (true) {
        const scheduled = await runtime.harness.commit((tx) => tx.scanTasks({}, 100), context);
        if (scheduled.items.filter((task) => task.kind === "pi.tool" || task.kind === "excel.direct-tool").length === 2) break;
        await new Promise<void>((resolve) => setImmediate(resolve));
      }
      await new Promise<void>((resolve) => setImmediate(resolve));
      await runtime.acknowledgeCommittedResults();
      const pendingFirst = await runtime.mutations.pending();
      assert.equal(pendingFirst.length, 1);
      assert.equal(pendingFirst[0]?.phase, "returned");
      assert.equal(value, 1, "the competing write must wait despite the released workbook lock");
      releases[0].release();
      await returned[1].promise;
      await runtime.acknowledgeCommittedResults();
      const pendingSecond = await runtime.mutations.pending();
      assert.equal(pendingSecond.length, 1);
      assert.equal(pendingSecond[0]?.phase, "returned");
      assert.notEqual(pendingSecond[0]?.mutationId, pendingFirst[0]?.mutationId);
      releases[1].release();
      await Promise.all([first, second]);
      await runtime.acknowledgeCommittedResults();
      assert.equal(value, 2);
      assert.deepEqual(await runtime.mutations.pending(), []);
      const tasks = await runtime.harness.commit((tx) => tx.scanTasks({}, 100), context);
      const writes = tasks.items.filter((task) => task.kind === "pi.tool" || task.kind === "excel.direct-tool");
      assert.equal(writes.length, 2);
      assert.equal(new Set(writes.map((task) => task.id)).size, 2);
      const conversations = new Set<string>();
      for (const task of writes) {
        const doc = await runtime.harness.snapshot(MutationDoc, `task:${task.id}`, context);
        assert.ok(doc?.record);
        assert.equal(parseMutationRecord(doc.record).phase, "resultCommitted");
        assert.equal(doc.binding?.taskId, task.id);
        assert.ok(doc.binding);
        conversations.add(doc.binding.conversationId);
      }
      assert.equal(conversations.size, 2);
      const entries = await direct.entries({}, 100, undefined, context);
      const receipts = entries.items.filter(DirectToolResultEntry.is);
      assert.equal(receipts.length, 1);
      assert.equal(receipts[0]?.data?.status, "completed");
      const results = (await modelConversation.context(context)).messages.filter((message) => message.role === "toolResult");
      assert.equal(results.length, 1);
      assert.equal(results[0]?.isError, false);
    } finally {
      for (const release of releases) release.release();
      await runtime.harness.close(context);
    }
  });
}

void test("direct native task writes receipt and acknowledgement without a model or assistant", async (t) => {
  const coordinator = createWorkbookCoordinator();
  let value = 0;
  const observation = () => ({ workbookId: "book", structure: {}, state: value, complete: true, limitation: null });
  const runtime = await openDurableRuntime({ storage: new MemoryStorage(), models: createModels(), workbookId: "book",
    coordinator, autoCompaction: false, mutationHost: {
      prepare: (call) => Promise.resolve({ intent: call.params, target: {}, before: observation(), expected: { ...observation(), state: value + 1 } }),
      observe: () => Promise.resolve(observation()),
    } });
  try {
    const conversation = await runtime.harness.root(context);
    const tool: AgentTool = { name: "write_cells", label: "Write", description: "Write", parameters: Type.Object({}),
      execute: () => { value++; return Promise.resolve({ content: [{ type: "text" as const, text: "written" }], details: {} }); } };
    await runtime.install(conversation, { tools: [tool], systemPrompt: "", bindMutationGate: (gate, id) =>
      withWorkbookCoordinator([tool], coordinator, { getWorkbookId: () => Promise.resolve("book"), getSessionId: () => String(conversation.id) },
        undefined, { mutationGate: gate, getMutationId: () => id }) }, (messages) => Promise.resolve(messages));
    await runtime.executeTool(conversation, "write_cells", {});
    assert.equal(value, 1);
    assert.deepEqual(await runtime.mutations.pending(), []);
    const entries = await conversation.entries({}, 100, undefined, context);
    assert.equal(entries.items.filter(DirectToolResultEntry.is).length, 1);
    assert.equal((await conversation.context(context)).messages.filter((message) => message.role === "assistant").length, 0);
    await assert.rejects(runtime.executeTool(conversation, "write_cells", null));
    assert.equal(value, 1);
    const tasks = await runtime.harness.commit((tx) => tx.scanTasks({}, 100), context);
    const firstTask = tasks.items[0];
    assert.ok(firstTask);
    const settled = (await runtime.harness.snapshot(MutationDoc, `task:${firstTask.id}`, context))?.record;
    assert.ok(settled);
    await runtime.harness.commit(async (tx) => {
      for (let i = 0; i < 300; i++) (await tx.doc(MutationDoc, `cold:${i}`, null)).record = settled;
    }, context);
    const snapshot = t.mock.method(runtime.harness, "snapshot");
    await runtime.executeTool(conversation, "write_cells", {});
    assert.equal(value, 2);
    const coldReads = snapshot.mock.calls.filter(({ arguments: args }) => args[0] === MutationDoc && String(args[1]).startsWith("cold:")).length;
    assert.equal(coldReads, 0, "settled forensic documents must not be loaded on the next mutation");
  } finally { await runtime.harness.close(context); }
});

for (const restart of [false, true]) {
  void test(`direct task ${restart ? "restart" : "cancellation"} never repeats native invocation`, async () => {
    class ReopenableStorage extends MemoryStorage { override async close(): Promise<void> {} }
    const storage = new ReopenableStorage();
    const options = { storage, models: createModels(), workbookId: "book", coordinator: createWorkbookCoordinator(), autoCompaction: false,
      mutationHost: {
        prepare: () => Promise.reject(new Error("Read tool must not prepare a mutation")),
        observe: () => Promise.reject(new Error("Read tool must not observe a mutation")),
      } };
    let calls = 0;
    let entered = (): void => {};
    const started = new Promise<void>((resolve) => { entered = resolve; });
    const tool: AgentTool = { name: "read_range", label: "Read", description: "Read", parameters: Type.Object({}),
      execute: async (_id, _args, signal) => {
        calls++;
        entered();
        await new Promise<void>((_resolve, reject) => signal?.addEventListener("abort", () => reject(new Error("Interrupted")), { once: true }));
        return { content: [], details: {} };
      } };
    let runtime = await openDurableRuntime(options);
    try {
      let conversation = await runtime.harness.root(context);
      await runtime.install(conversation, { tools: [tool], systemPrompt: "" }, (messages) => Promise.resolve(messages));
      const result = runtime.executeTool(conversation, "read_range", {}).catch(() => null);
      await started;
      if (restart) {
        await runtime.harness.close(context);
        await result;
        runtime = await openDurableRuntime(options);
        conversation = await runtime.harness.root(context);
        await runtime.install(conversation, { tools: [tool], systemPrompt: "" }, (messages) => Promise.resolve(messages));
        runtime.harness.resume();
        const tasks = await runtime.harness.commit((tx) => tx.scanTasks({}, 100), context);
        const task = tasks.items[0];
        assert.ok(task);
        await runtime.harness.waitForTask(task.id, context);
      } else {
        await conversation.abort(context);
        await result;
      }
      assert.equal(calls, 1);
      const entries = await conversation.entries({}, 100, undefined, context);
      const receipt = entries.items.find(DirectToolResultEntry.is);
      assert.equal(receipt?.data?.status, "aborted");
    } finally { await runtime.harness.close(context); }
  });
}
