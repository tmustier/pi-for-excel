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

void test("singleton writes reconcile canonical host addresses after sync, including anchored and quoted references", async () => {
  for (const start_cell of ["A1", "'O''Brien! Data'!$A$1"]) {
    await withOffice(async (f) => {
      f.sheet.name = "O'Brien! Data";
      const j = journal();
      j.controller.configureFault("afterDispatch");
      const params = { start_cell, values: [[53]], allow_overwrite: true };
      const execution = j.write.execute("single", params);
      try {
        for (let i = 0; i < 100 && !j.controller.status().paused; i++) await new Promise<void>((resolve) => setImmediate(resolve));
        assert.equal(j.controller.status().paused, "afterDispatch");
        assert.deepEqual(j.records.get("single")?.before.state, [{ range: "A1", content: [["old"]] }]);
        assert.equal((await j.controller.reconcile("single")).classification, "applied");
        assert.deepEqual(f.read(), [[53, 17], ["anchor", 29]]);
      } finally {
        j.controller.resume();
        await execution;
      }
      await assert.rejects(j.write.execute("single", params), /replay/u);
      f.edit([[99, 17], ["anchor", 29]]);
      const review = await j.controller.reconcile("single");
      assert.equal(review.classification, "conflict");
      await j.controller.resolve("single", review.observationToken);
      assert.deepEqual(f.read(), [[99, 17], ["anchor", 29]]);
      assert.equal(f.writes(), 1);
    });
  }
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
      const review = await j.controller.reconcile("structure");
      assert.equal(review.observation.complete, false);
      const edited = structuredClone(actual);
      const firstRow = edited[0];
      assert.ok(firstRow);
      firstRow[0] = "=1+1";
      f.edit(edited);
      const changed = await j.controller.reconcile("structure");
      assert.equal(changed.classification, "conflict");
      assert.notEqual(changed.observationToken, review.observationToken);
      await assert.rejects(j.controller.resolve("structure", review.observationToken), /changed/u);
      assert.deepEqual(f.read(), edited);
      await j.controller.resolve("structure", changed.observationToken);
      assert.deepEqual(f.read(), edited);
      await assert.rejects(j.structure.execute("structure", { action, position: 1 }), /replay/u);
      f.onWrite(() => assert.equal(j.records.get("new-task")?.phase, "dispatched"));
      await j.write.execute("new-task", { start_cell: "B2", values: [[71]], allow_overwrite: true });
      assert.equal(f.read()[0]?.[0], "=1+1");
      assert.equal(f.read()[1]?.[1], 71);
      assert.equal(f.writes(), 2);
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
    const rejected = assert.rejects(result, /not executed/u);
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

void test("structural evidence samples boundary values and formulas for every row/column action", async () => {
  await withOffice(async (f) => {
    const grid = Array.from({ length: 6 }, (_, row) => Array.from({ length: 6 }, (_, col) => `${row + 1},${col + 1}`));
    const secondRow = grid[1];
    assert.ok(secondRow);
    secondRow[1] = "=1+1";
    f.edit(grid);
    f.setUsedRange("B2:F6");
    const j = journal();
    for (const action of ["insert_rows", "delete_rows", "insert_columns", "delete_columns"]) {
      const params = { action, sheet: "Original", position: 3, count: 2 };
      const plan = await j.host.prepare({ mutationId: action, toolCallId: "raw", toolName: "modify_structure",
        workbookId: "book", sessionId: "s", params });
      const rowAction = action.endsWith("rows");
      const coordinates = rowAction
        ? ["B2", "D2", "F2", "B3", "D3", "F3", "B4", "D4", "F4", "B5", "D5", "F5"]
        : ["B2", "B4", "B6", "C2", "C4", "C6", "D2", "D4", "D6", "E2", "E4", "E6"];
      const expectedCells = coordinates.map((range) => {
        const value = range === "B2" ? 2 : `${range.slice(1)},${range.charCodeAt(0) - 64}`;
        return { range, values: [[value]], formulas: [[range === "B2" ? "=1+1" : ""]] };
      });
      assert.deepEqual(plan.before.state, {
        affectedRanges: rowAction ? ["3:4"] : ["C:D"],
        neighborhood: {
          samplingPolicy: "agent-default: band before/first/last/after x used-axis first/middle/last; empty axis starts at A/1; at most 12 cells per band",
          usedExtent: { rowIndex: 1, columnIndex: 1, rowCount: 5, columnCount: 5 },
          cells: expectedCells,
        },
      });
      assert.deepEqual(plan.intent, params);
      assert.match(JSON.stringify(plan.before.structure), /sheet-1/u);
      assert.match(JSON.stringify(plan.before.structure), /table-1/u);
      assert.equal(plan.before.complete, false);
      assert.equal(plan.expected, null);
      assert.equal(f.writes(), 0);
    }
  });
});

void test("structural neighborhood stays bounded on a full-sheet used range without truncating intent", async () => {
  await withOffice(async (f) => {
    f.setUsedRange("A1:XFD1048576");
    const dataLoads: string[] = [];
    f.onRangeLoad((address, properties) => {
      if (/values|formulas/u.test(properties)) {
        dataLoads.push(address);
        // Throw before the fixture could allocate any enormous range data.
        assert.match(address, /^[A-Z]+\d+$/u, "Only individual sample cells may load contents");
      }
    });
    const j = journal();
    for (const action of ["insert_rows", "delete_rows", "insert_columns", "delete_columns"]) {
      dataLoads.length = 0;
      const params = { action, sheet: "Original", position: 3, count: 1000 };
      const plan = await j.host.prepare({ mutationId: action, toolCallId: "raw", toolName: "modify_structure",
        workbookId: "book", sessionId: "s", params });
      assert.deepEqual(plan.intent, params);
      assert.equal(dataLoads.length, 12);
      assert.equal(new Set(dataLoads).size, 12);
      assert.ok(plan.before.state);
      assert.match(JSON.stringify(plan.before.state), action.endsWith("rows") ? /"3:1002"/u : /"C:ALN"/u);
      assert.equal(plan.before.complete, false);
      assert.equal(plan.expected, null);
    }
    assert.equal(f.writes(), 0);
  });
});

void test("structural samples respect physical sheet edges and tool count normalization", async () => {
  await withOffice(async (f) => {
    f.setUsedRange("A1:XFD1048576");
    const loaded: string[] = [];
    f.onRangeLoad((address, properties) => { if (/values|formulas/u.test(properties)) loaded.push(address); });
    const j = journal();
    for (const [action, position, expected] of [
      ["insert_rows", 1, ["A1", "LCB1", "XFD1", "A2", "LCB2", "XFD2"]],
      ["delete_rows", 1048576, ["A1048575", "LCB1048575", "XFD1048575", "A1048576", "LCB1048576", "XFD1048576"]],
      ["insert_columns", 1, ["A1", "A524288", "A1048576", "B1", "B524288", "B1048576"]],
      ["delete_columns", 16384, ["XFC1", "XFC524288", "XFC1048576", "XFD1", "XFD524288", "XFD1048576"]],
    ] as const) {
      for (const count of [undefined, 1.9, 0, -2]) {
        loaded.length = 0;
        const params = { action, position, ...(count === undefined ? {} : { count }) };
        const plan = await j.host.prepare({ mutationId: action, toolCallId: "raw", toolName: "modify_structure",
          workbookId: "book", sessionId: "s", params });
        assert.deepEqual(loaded, expected);
        assert.deepEqual(plan.intent, params);
        assert.equal(plan.before.complete, false);
        assert.ok(plan.before.state);
      }
    }
  });
});

void test("structural prepared evidence detects neighboring user edits but matching samples never prove unchanged", async () => {
  await withOffice(async (f) => {
    const j = journal();
    j.controller.configureFault("beforeDispatch");
    const execution = j.structure.execute("prepared", { action: "insert_rows", position: 2 });
    const rejected = assert.rejects(execution, /not executed/u);
    try {
      for (let i = 0; i < 100 && !j.controller.status().paused; i++) await new Promise<void>((resolve) => setImmediate(resolve));
      assert.equal(j.controller.status().paused, "beforeDispatch");
      const review = await j.controller.reconcile("prepared");
      assert.equal(review.classification, "conflict");
      assert.equal(review.observation.complete, false);
      f.edit([["user edit", 17], ["anchor", 29]]);
      const changed = await j.controller.reconcile("prepared");
      assert.equal(changed.classification, "conflict");
      assert.notEqual(changed.observationToken, review.observationToken);
    } finally {
      j.controller.resume();
      await rejected;
    }
    assert.deepEqual(f.read(), [["user edit", 17], ["anchor", 29]]);
    assert.equal(f.writes(), 0);
    assert.equal((await j.controller.pending()).length, 1);
  });
});
