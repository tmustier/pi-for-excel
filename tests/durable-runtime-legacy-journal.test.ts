import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { createRegistry, defineDoc, Harness, MemoryStorage } from "@earendil-works/pi-durable";
import { MutationDoc, PendingMutationsDoc, openDurableRuntime } from "../src/durable/runtime.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import type { JsonValue } from "@earendil-works/chord";

void test("legacy V1/V2 evidence is retained; stale observation rejected and explicit resolution preserves originals", async () => {
  class ReopenableMemoryStorage extends MemoryStorage { override async close(): Promise<void> {} }
  const storage = new ReopenableMemoryStorage();
  const old = defineDoc<{ records: Record<string, JsonValue> }>({ kind: "excel.workbook-mutations", version: 1,
    scope: "session", initial: () => ({ records: {} }) });
  const originals = { one: { version: 1, evidence: "unchanged original one" }, two: { version: 2, evidence: "unchanged original two" } };
  const seed = await Harness.open(storage, { models: createModels(), registry: createRegistry() }, context);
  await seed.commit(async (tx) => { (await tx.doc(old)).records = originals; }, context);
  await seed.close(context);
  let value = 1;
  const runtime = await openDurableRuntime({ storage, models: createModels(), workbookId: "book", autoCompaction: false,
    coordinator: createWorkbookCoordinator(), mutationHost: {
      prepare: () => Promise.resolve({ intent: {}, target: {}, before: { workbookId: "book", structure: {}, state: value, complete: true, limitation: null }, expected: null }),
      observe: () => Promise.resolve({ workbookId: "book", structure: {}, state: value, complete: true, limitation: null }),
    } });
  try {
    assert.equal((await runtime.reviews()).length, 2);
    const review = await runtime.mutations.reconcile("one");
    value++;
    await assert.rejects(runtime.mutations.resolve("one", review.observationToken), /changed/);
    const fresh = await runtime.mutations.reconcile("one");
    await runtime.mutations.resolve("one", fresh.observationToken);
    assert.deepEqual((await runtime.harness.snapshot(MutationDoc, "one", context))?.original, originals.one);
    assert.deepEqual((await runtime.harness.snapshot(old, context))?.records, originals);
    assert.deepEqual(Object.keys((await runtime.harness.snapshot(PendingMutationsDoc, context))?.ids ?? {}), ["two"]);
  } finally { await runtime.harness.close(context); }
});
