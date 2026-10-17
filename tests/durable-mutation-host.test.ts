import assert from "node:assert/strict";
import { test } from "node:test";
import { createWriteCellsTool } from "../src/tools/write-cells.ts";
import { createModifyStructureTool } from "../src/tools/modify-structure.ts";
import { createWorkbookMutationHost } from "../src/durable/mutation-host.ts";
import { createMutationController, type MutationRecord } from "../src/durable/mutations.ts";
import { withWorkbookCoordinator } from "../src/tools/with-workbook-coordinator.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import { mutationOfficeFixture } from "./fixtures/mutation-office.ts";

async function withOffice(run: (f: ReturnType<typeof mutationOfficeFixture>) => Promise<void>): Promise<void> {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "Excel");
  const f = mutationOfficeFixture();
  Reflect.set(globalThis, "Excel", f.excel);
  try { await run(f); } finally {
    if (previous) Object.defineProperty(globalThis, "Excel", previous);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
}

function journal() {
  const records = new Map<string, MutationRecord>();
  const coordinator = createWorkbookCoordinator();
  const host = createWorkbookMutationHost({ kind: "office", getWorkbookId: () => Promise.resolve("book") });
  const store = { load: (id: string) => Promise.resolve(structuredClone(records.get(id) ?? null)),
    list: () => Promise.resolve(structuredClone([...records.values()])),
    save: (record: MutationRecord) => { records.set(record.mutationId, structuredClone(record)); return Promise.resolve(); } };
  const controller = createMutationController({ workbookId: "book", store, host, coordinator, allowFaultInjection: true });
  const tools = withWorkbookCoordinator([
    createWriteCellsTool({ appendAuditEntry: async () => {}, appendRecoverySnapshot: () => Promise.resolve(null) }),
    createModifyStructureTool(),
  ], coordinator, { getWorkbookId: () => Promise.resolve("book"), getSessionId: () => "s" }, undefined,
  { mutationGate: controller, getMutationId: (id) => id });
  const write = tools[0], structure = tools[1];
  assert.ok(write && structure);
  return { records, controller, write, structure, host };
}

void test("actual write_cells + production observer: formula proof, stable IDs, anchors, and no repeat after sync/result gap", async () => {
  await withOffice(async (f) => {
    const j = journal();
    f.onWrite(() => assert.equal(j.records.get("write")?.phase, "dispatched"));
    j.controller.configureFault("beforeResult");
    const result = j.write.execute("write", { start_cell: "A1", values: [["new", "=1+1"]], allow_overwrite: true });
    for (let i = 0; i < 100 && !j.controller.status().paused; i++) await new Promise<void>((resolve) => setImmediate(resolve));
    assert.equal(j.controller.status().paused, "beforeResult");
    assert.deepEqual(f.read(), [["new", "=1+1"], ["anchor", 29]]);
    assert.equal((await j.controller.reconcile("write")).classification, "applied");
    const record = j.records.get("write");
    assert.ok(record);
    assert.match(JSON.stringify(record.target), /sheet-1/u);
    assert.match(JSON.stringify(record.before.structure), /table-1/u);
    j.controller.resume();
    await result;
    f.edit([["new", "=1+1"], ["user changed anchor", 29]]);
    assert.equal((await j.controller.reconcile("write")).classification, "conflict");
    await assert.rejects(j.write.execute("write", { start_cell: "A1", values: [["bad"]], allow_overwrite: true }), /replay/u);
    assert.equal(f.writes(), 1);
  });
});

for (const action of ["insert_rows", "delete_rows", "insert_columns", "delete_columns"]) {
  void test(`actual modify_structure ${action}: sync applies once, interruption stays uncertain`, async () => {
    await withOffice(async (f) => {
      const j = journal();
      f.onWrite(() => assert.equal(j.records.get("structure")?.phase, "dispatched"));
      f.failAfterWrite();
      await j.structure.execute("structure", { action, position: 1, count: 1 });
      assert.equal(f.writes(), 1);
      assert.equal((await j.controller.reconcile("structure")).classification, "conflict");
      // A caught host error is still uncertain even after Durable persists its error receipt.
      await j.controller.resultCommitted("structure");
      await assert.rejects(j.structure.execute("retry", { action, position: 1 }), /pending/u);
      assert.equal(f.writes(), 1);
      const actual = f.read();
      if (action === "insert_rows") assert.deepEqual(actual, [["", ""], ["old", 17], ["anchor", 29]]);
      if (action === "delete_rows") assert.deepEqual(actual, [["anchor", 29]]);
      if (action === "insert_columns") assert.deepEqual(actual, [["", "old", 17], ["", "anchor", 29]]);
      if (action === "delete_columns") assert.deepEqual(actual, [[17], [29]]);
    });
  });
}

void test("production observer follows stable sheet identity for rename and detects name reuse", async () => {
  await withOffice(async (f) => {
    const j = journal();
    await j.structure.execute("rename", { action: "rename_sheet", sheet: "Original", new_name: "Renamed" });
    assert.equal(f.sheet.name, "Renamed");
    assert.equal((await j.controller.reconcile("rename")).classification, "applied");
    f.sheet.id = "replacement-id";
    assert.equal((await j.controller.reconcile("rename")).classification, "conflict");
  });
});

void test("changing the active sheet while intent commits blocks an unqualified write before dispatch", async () => {
  await withOffice(async (f) => {
    const activateOther = await f.excel.run((context) => {
      const other = { ...f.sheet, id: "sheet-2", name: "Other", position: 1 };
      context.workbook.worksheets.items.push(other);
      return Promise.resolve(() => { context.workbook.worksheets.getActiveWorksheet = () => other; });
    });
    const j = journal();
    j.controller.configureFault("beforeDispatch");
    const result = j.write.execute("write", { start_cell: "A1", values: [["new"]], allow_overwrite: true });
    const rejected = assert.rejects(result, /Workbook changed/u);
    for (let i = 0; i < 100 && !j.controller.status().paused; i++) await new Promise<void>((resolve) => setImmediate(resolve));
    assert.equal(j.controller.status().paused, "beforeDispatch");
    activateOther();
    const review = await j.controller.reconcile("write");
    assert.equal(review.classification, "conflict");
    assert.match(review.reason, /never entered/u);
    j.controller.resume();
    await rejected;
    assert.equal(f.writes(), 0);
    assert.equal((await j.controller.pending()).length, 1);
  });
});

void test("one-cell write on a full-sheet synthetic used range loads only the target and adjacent cells", async () => {
  await withOffice(async (f) => {
    f.setUsedRange("A1:XFD1048576");
    const loaded = new Set<string>();
    f.onRangeLoad((address, properties) => {
      if (!/values|formulas/u.test(properties)) return;
      loaded.add(address);
      // Prevent a broken observer from allocating the synthetic sheet in this test.
      if (address === "A1:XFD1048576") throw new Error("Full-sheet data load forbidden");
    });
    const j = journal();
    const params = { start_cell: "A1", values: [["replacement"]], allow_overwrite: true };
    await j.write.execute("bounded", params);
    assert.deepEqual([...loaded].sort(), ["A1:A1", "A2", "B1"]);
    assert.deepEqual(f.read(), [["replacement", 17], ["anchor", 29]]);
    assert.deepEqual(j.records.get("bounded")?.intent, params);
    assert.equal((await j.controller.reconcile("bounded")).classification, "applied");
    assert.match(j.records.get("bounded")?.before.limitation ?? "", /does not prove unrelated sheet state/u);
    await assert.rejects(j.write.execute("bounded", params), /replay/u);
    await assert.rejects(j.write.execute("retry", params), /pending/u);
    assert.equal(f.writes(), 1);
  });
});

void test("adjacent evidence preserves full multi-cell intent and respects the last Excel cell", async () => {
  await withOffice(async (f) => {
    const loaded = new Set<string>();
    f.onRangeLoad((address, properties) => { if (/values|formulas/u.test(properties)) loaded.add(address); });
    const j = journal();
    const params = { start_cell: "C3", values: [[11, 12, 13], [21, 22, 23]] };
    const plan = await j.host.prepare({ mutationId: "rectangle", toolCallId: "raw", toolName: "write_cells",
      workbookId: "book", sessionId: "s", params });
    assert.deepEqual(plan.intent, params);
    assert.deepEqual(plan.expected?.state, [{ range: "C3:E4", content: params.values }]);
    assert.deepEqual([...loaded].sort(), ["B3", "B4", "C2", "C3:E4", "C5", "E2", "E5", "F3", "F4"]);
    loaded.clear();
    await j.host.prepare({ mutationId: "edge", toolCallId: "raw", toolName: "write_cells",
      workbookId: "book", sessionId: "s", params: { start_cell: "XFD1048576", values: [[31]] } });
    assert.deepEqual([...loaded].sort(), ["XFC1048576", "XFD1048575", "XFD1048576:XFD1048576"]);
  });
});

void test("structural evidence retains intended row/column regions and IDs without copying sheet contents", async () => {
  await withOffice(async (f) => {
    f.setUsedRange("A1:XFD1048576");
    const dataLoads: string[] = [];
    f.onRangeLoad((address, properties) => {
      if (/values|formulas/u.test(properties)) {
        dataLoads.push(address);
        throw new Error("Structural journal must not copy sheet data");
      }
    });
    const j = journal();
    for (const action of ["insert_rows", "delete_rows", "insert_columns", "delete_columns", "duplicate_sheet", "delete_sheet"]) {
      const params = { action, sheet: "Original", position: 3, count: 2 };
      const plan = await j.host.prepare({ mutationId: action, toolCallId: "raw", toolName: "modify_structure",
        workbookId: "book", sessionId: "s", params });
      assert.deepEqual(plan.intent, params);
      assert.deepEqual(plan.before.state, { affectedRanges: action.endsWith("rows") ? ["3:4"] : action.endsWith("columns") ? ["C:D"] : [] });
      assert.match(JSON.stringify(plan.before.structure), /sheet-1/u);
      assert.equal(plan.before.complete, false);
      assert.equal(plan.expected, null);
    }
    assert.deepEqual(dataLoads, []);
  });
});
