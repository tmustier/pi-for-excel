import assert from "node:assert/strict";
import { test } from "node:test";
import { createMutationController, type MutationFaultPoint, type MutationRecord } from "../src/durable/mutations.ts";
import { createWriteCellsTool, type WriteCellsGrid } from "../src/tools/write-cells.ts";

function fixture(allowFaultInjection = true) {
  const records = new Map<string, MutationRecord>();
  const saves: MutationRecord[] = [];
  const store = {
    load: (id: string): Promise<MutationRecord | null> => Promise.resolve(structuredClone(records.get(id) ?? null)),
    save: (record: MutationRecord): Promise<void> => {
      saves.push(structuredClone(record));
      records.set(record.toolCallId, structuredClone(record));
      return Promise.resolve();
    },
  };
  const controller = createMutationController({ workbookId: "opaque-book-id", store, allowFaultInjection });
  let values: WriteCellsGrid = [["old", ""]];
  let pending: WriteCellsGrid | null = null;
  let writes = 0;
  let rejectVerification = false;
  const range = {
    address: "Original!A1:B1",
    load: (): void => {},
    get values(): WriteCellsGrid { return structuredClone(values); },
    set values(grid: WriteCellsGrid) { pending = structuredClone(grid); },
    get formulas(): WriteCellsGrid { return values.map((row) => row.map((v) => typeof v === "string" && v.startsWith("=") ? v : "")); },
  };
  const sheet = { id: "stable-sheet-id", name: "Original", load: (): void => {}, getRange: () => range };
  const context = {
    workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet } },
    sync: (): Promise<void> => {
      if (pending) { values = pending; pending = null; writes += 1; }
      else if (writes && rejectVerification) throw new Error("verification unavailable");
      return Promise.resolve();
    },
  };
  const tool = createWriteCellsTool({
    ...controller.hooks,
    appendAuditEntry: async () => {},
    appendRecoverySnapshot: () => Promise.resolve(null),
  });
  return {
    controller, store, records, saves, sheet,
    writes: () => writes,
    read: (record: MutationRecord) => {
      assert.equal(record.workbookId, "opaque-book-id");
      assert.equal(record.sheetId, sheet.id);
      return Promise.resolve({ values: range.values, formulas: range.formulas });
    },
    edit: () => { values = [["user edit", ""]]; },
    rejectVerification: () => { rejectVerification = true; },
    execute: () => tool.execute("call-1", { start_cell: "A1", values: [["new", "=1+1"]], allow_overwrite: true }),
    host: { run: <T>(callback: (ctx: typeof context) => Promise<T>): Promise<T> => callback(context) },
  };
}

async function withHost(action: (f: ReturnType<typeof fixture>) => Promise<void>, enabled = true): Promise<void> {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "Excel");
  const f = fixture(enabled);
  Reflect.set(globalThis, "Excel", f.host);
  try { await action(f); } finally {
    if (previous) Object.defineProperty(globalThis, "Excel", previous);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
}

for (const point of ["beforeSync", "afterSync", "beforeResult"] satisfies MutationFaultPoint[]) {
  void test(`write_cells pauses once at ${point} with committed intent and independent reconciliation`, async () => {
    await withHost(async (f) => {
      f.controller.configureFault(point);
      let completed = false;
      const execution = f.execute().then((result) => { completed = true; return result; });
      for (let i = 0; i < 100 && !f.controller.status().paused; i += 1) {
        await new Promise<void>((resolve) => setImmediate(resolve));
      }
      assert.deepEqual(f.controller.status(), { armed: null, paused: point, toolCallId: "call-1" });
      assert.equal(completed, false);
      assert.equal(f.writes(), point === "beforeSync" ? 0 : 1);
      const record = f.records.get("call-1");
      assert.ok(record);
      assert.deepEqual(record.beforeValues, [["old", ""]]);
      assert.deepEqual(record.beforeFormulas, [["", ""]]);
      assert.deepEqual(record.intendedValues, [["new", "=1+1"]]);
      assert.equal(record.range, "A1:B1");
      assert.equal(record.phase, point === "beforeSync" ? "intended" : point === "afterSync" ? "synced" : "resultReady");
      f.sheet.name = "Renamed";
      // Simulate reopening with only persisted state: no live tool state is needed.
      const reopened = createMutationController({ workbookId: "opaque-book-id", store: f.store });
      assert.equal(await reopened.reconcile("call-1", f.read), point === "beforeSync" ? "unchanged" : "applied");
      f.controller.resume();
      const result = await execution;
      assert.equal(result.details.blocked, false);
      assert.deepEqual(f.saves.map((r) => r.phase), ["intended", "synced", "resultReady"]);
      f.edit();
      assert.equal(await reopened.reconcile("call-1", f.read), "conflict");
      assert.equal(f.writes(), 1);
      await assert.rejects(f.execute(), /already recorded/u);
      assert.equal(f.writes(), 1);
    });
  });
}

void test("fault injection is explicitly gated and post-commit verification never reports a failed write", async () => {
  await withHost(async (f) => {
    assert.throws(() => f.controller.configureFault("beforeSync"), /disabled/u);
    f.controller.configureFault(null);
    f.rejectVerification();
    const result = await f.execute();
    assert.equal(f.writes(), 1);
    assert.match(result.content.map((c) => c.type === "text" ? c.text : "").join(""), /Cells were written/u);
    assert.equal(await f.controller.reconcile("call-1", f.read), "applied");
  }, false);
});

void test("default controller disallows faults; post-sync persistence failure rejects without failed-write receipt", async () => {
  const controller = createMutationController({ workbookId: "book", store: { load: () => Promise.resolve(null), save: async () => {} } });
  assert.throws(() => controller.configureFault("afterSync"), /disabled/u);
  await withHost(async (f) => {
    f.store.save = (record) => {
      if (record.phase === "synced") throw new Error("commit unavailable");
      f.records.set(record.toolCallId, structuredClone(record));
      return Promise.resolve();
    };
    await assert.rejects(f.execute(), /commit unavailable/u);
    assert.equal(f.writes(), 1);
    assert.equal(await f.controller.reconcile("call-1", f.read), "applied");
  });
});
