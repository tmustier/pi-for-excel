import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentToolResult } from "@earendil-works/pi-agent-core";

import { WorkbookChangeAuditLog } from "../src/audit/workbook-change-audit.ts";
import { createReadRangeTool } from "../src/tools/read-range.ts";
import { createWorkbookHistoryTool } from "../src/tools/workbook-history.ts";
import { createWriteCellsTool } from "../src/tools/write-cells.ts";
import { composeCoreToolsForHost } from "../src/tools/host-selection.ts";
import type { CoreToolName } from "../src/tools/names.ts";
import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import { createInMemorySettingsStore } from "./fixtures/recovery-log.ts";

interface StoredCell {
  value: DynamicValue;
  formula: string;
}

interface ParsedRange {
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
}

function parseRange(address: string): ParsedRange {
  const cellPart = address.includes("!") ? address.slice(address.indexOf("!") + 1) : address;
  const [startText = "", endText = startText] = cellPart.split(":");
  const parseCell = (cell: string): { col: number; row: number } => {
    const match = /^([A-Z]+)(\d+)$/iu.exec(cell);
    if (!match?.[1] || !match[2]) throw new Error(`Unsupported test address: ${address}`);
    let col = 0;
    for (const letter of match[1].toUpperCase()) col = col * 26 + letter.charCodeAt(0) - 64;
    return { col: col - 1, row: Number(match[2]) - 1 };
  };
  const start = parseCell(startText);
  const end = parseCell(endText);
  return { startCol: start.col, startRow: start.row, endCol: end.col, endRow: end.row };
}

class RangeStoreRange {
  readonly address: string;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly numberFormat: DynamicValue[][];
  private readonly store: RangeStore;
  private readonly parsed: ParsedRange;

  constructor(store: RangeStore, address: string) {
    this.store = store;
    this.address = address.includes("!") ? address.slice(address.indexOf("!") + 1) : address;
    this.parsed = parseRange(this.address);
    this.rowCount = this.parsed.endRow - this.parsed.startRow + 1;
    this.columnCount = this.parsed.endCol - this.parsed.startCol + 1;
    this.numberFormat = Array.from({ length: this.rowCount }, () =>
      Array.from({ length: this.columnCount }, () => "General"));
  }

  load(_properties?: string | string[]): void {}

  get values(): DynamicValue[][] {
    return this.store.readGrid(this.parsed, "value");
  }

  set values(values: DynamicValue[][]) {
    this.store.writeGrid(this.parsed, values);
  }

  get formulas(): DynamicValue[][] {
    return this.store.readGrid(this.parsed, "formula");
  }
}

class RangeStoreSheet {
  readonly name = "Sheet1";
  readonly comments = { items: [] as DynamicValue[], load: (_properties?: string) => undefined };
  private readonly store: RangeStore;

  constructor(store: RangeStore) {
    this.store = store;
  }

  load(_properties?: string | string[]): void {}

  getRange(address: string): RangeStoreRange {
    return new RangeStoreRange(this.store, address);
  }
}

class RangeStore {
  private readonly cells = new Map<string, StoredCell>();
  readonly sheet = new RangeStoreSheet(this);
  readonly context = {
    workbook: {
      worksheets: {
        getActiveWorksheet: (): RangeStoreSheet => this.sheet,
        getItem: (name: string): RangeStoreSheet => {
          if (name !== this.sheet.name) throw new Error(`Worksheet not found: ${name}`);
          return this.sheet;
        },
      },
    },
    sync: (): Promise<void> => Promise.resolve(),
  };

  private key(row: number, col: number): string {
    return `${row}:${col}`;
  }

  readGrid(parsed: ParsedRange, field: keyof StoredCell): DynamicValue[][] {
    const rows: DynamicValue[][] = [];
    for (let row = parsed.startRow; row <= parsed.endRow; row += 1) {
      const values: DynamicValue[] = [];
      for (let col = parsed.startCol; col <= parsed.endCol; col += 1) {
        const cell = this.cells.get(this.key(row, col));
        values.push(cell ? cell[field] : field === "formula" ? "" : null);
      }
      rows.push(values);
    }
    return rows;
  }

  writeGrid(parsed: ParsedRange, values: DynamicValue[][]): void {
    for (let rowOffset = 0; rowOffset <= parsed.endRow - parsed.startRow; rowOffset += 1) {
      for (let colOffset = 0; colOffset <= parsed.endCol - parsed.startCol; colOffset += 1) {
        const value = values[rowOffset]?.[colOffset] ?? null;
        const formula = typeof value === "string" && value.startsWith("=") ? value : "";
        this.cells.set(this.key(parsed.startRow + rowOffset, parsed.startCol + colOffset), {
          // This narrow double does not calculate formulas; Excel/WPS acceptance must verify values.
          value: formula ? null : value,
          formula,
        });
      }
    }
  }

  applySnapshot(address: string, values: DynamicValue[][]): {
    values: DynamicValue[][];
    formulas: DynamicValue[][];
  } {
    const range = new RangeStoreRange(this, address);
    const before = { values: range.values, formulas: range.formulas };
    range.values = values;
    return before;
  }
}

async function withRangeStore<T>(store: RangeStore, action: () => Promise<T>): Promise<T> {
  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <TResult>(callback: (context: DynamicValue) => Promise<TResult>): Promise<TResult> =>
      callback(store.context),
  });

  try {
    return await action();
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
}

function firstText<T>(result: AgentToolResult<T>): string {
  const block = result.content[0];
  if (!block || block.type !== "text") throw new Error("Expected text tool result.");
  return block.text;
}

void test("user can write, inspect, reject overwrite, and restore through workbook tools", async () => {
  const store = new RangeStore();
  const settings = createInMemorySettingsStore();
  let workbookId = "url_sha256:contract-book";
  const workbookContext = (): Promise<WorkbookContext> => Promise.resolve({
    workbookId,
    workbookName: "Contract.xlsx",
    source: "document.url",
  });
  let id = 0;
  const recovery = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settings),
    getWorkbookContext: workbookContext,
    createId: () => `snapshot-${id += 1}`,
    now: () => 1_700_000_000_000 + id,
    applySnapshot: (address, values) => Promise.resolve(store.applySnapshot(address, values)),
  });
  const audit = new WorkbookChangeAuditLog({
    getSettingsStore: () => Promise.resolve({ ...settings, delete: () => Promise.resolve() }),
    getWorkbookContext: workbookContext,
    createId: () => `audit-${id += 1}`,
    now: () => 1_700_000_000_000 + id,
  });
  const write = createWriteCellsTool({
    appendAuditEntry: (entry) => audit.append(entry),
    appendRecoverySnapshot: (args) => recovery.append(args),
  });
  const read = createReadRangeTool();
  const history = createWorkbookHistoryTool({
    getRecoveryLog: () => recovery,
    appendAuditEntry: (entry) => audit.append(entry),
  });

  await withRangeStore(store, async () => {
    const initialWrite = await write.execute("write-blank", {
      start_cell: "Sheet1!A1",
      values: [["Name", "Total"], ["Ada", 7]],
    });
    assert.match(firstText(initialWrite), /Written to/u);
    assert.deepEqual(initialWrite.details.kind, "write_cells");
    assert.equal(initialWrite.details.blocked, false);
    assert.equal(initialWrite.details.address, "Sheet1!A1:B2");
    assert.equal(initialWrite.details.changes?.changedCount, 4);
    assert.equal(initialWrite.details.recovery?.status, "checkpoint_created");

    const listed = await history.execute("history-list", { action: "list" });
    assert.equal(listed.details.count, 1);
    assert.equal(listed.details.snapshots?.[0]?.address, "Sheet1!A1:B2");
    assert.equal(listed.details.snapshots?.[0]?.workbookId, workbookId);
    assert.equal(listed.details.snapshots?.[0]?.workbookLabel, "Contract.xlsx");

    const readInitial = await read.execute("read-initial", { range: "Sheet1!A1:B2", mode: "csv" });
    assert.match(firstText(readInitial), /Name,Total\nAda,7/u);

    const blocked = await write.execute("write-blocked", {
      start_cell: "Sheet1!A1",
      values: [["Changed"]],
    });
    assert.equal(blocked.details.blocked, true);
    assert.match(firstText(blocked), /Write blocked/u);
    const readBlocked = await read.execute("read-blocked", { range: "Sheet1!A1", mode: "csv" });
    assert.match(firstText(readBlocked), /Name/u);

    const overwritten = await write.execute("write-overwrite", {
      start_cell: "Sheet1!A1",
      values: [["Changed"]],
      allow_overwrite: true,
    });
    assert.equal(overwritten.details.blocked, false);
    const overwriteSnapshotId = overwritten.details.recovery?.snapshotId;
    assert.ok(overwriteSnapshotId);

    const restored = await history.execute("history-restore", {
      action: "restore",
      snapshot_id: overwriteSnapshotId,
    });
    assert.match(firstText(restored), /Restored backup/u);
    assert.equal(restored.details.address, "Sheet1!A1:A1");
    assert.equal(restored.details.restoredSnapshotId, overwriteSnapshotId);
    const readRestored = await read.execute("read-restored", { range: "Sheet1!A1", mode: "csv" });
    assert.match(firstText(readRestored), /Name/u);

    const formulaWrite = await write.execute("write-formula", {
      start_cell: "Sheet1!C1",
      values: [["=A2+B2"]],
    });
    assert.equal(formulaWrite.details.blocked, false);
    const formulaRead = await read.execute("read-formula", { range: "Sheet1!C1" });
    assert.match(firstText(formulaRead), /C1: =A2\+B2/u);

    const replaceFormula = await write.execute("replace-formula", {
      start_cell: "Sheet1!C1",
      values: [[99]],
      allow_overwrite: true,
    });
    const formulaSnapshotId = replaceFormula.details.recovery?.snapshotId;
    assert.ok(formulaSnapshotId);
    await history.execute("restore-formula", { action: "restore", snapshot_id: formulaSnapshotId });
    const formulaRestored = await read.execute("read-formula-restored", { range: "Sheet1!C1" });
    assert.match(firstText(formulaRestored), /C1: =A2\+B2/u);

    const snapshotBeforeWrongWorkbook = overwritten.details.recovery?.snapshotId;
    assert.ok(snapshotBeforeWrongWorkbook);
    workbookId = "url_sha256:other-book";
    const wrongWorkbookRestore = await history.execute("wrong-workbook", {
      action: "restore",
      snapshot_id: snapshotBeforeWrongWorkbook,
    });
    assert.match(firstText(wrongWorkbookRestore), /different workbook/u);
    assert.match(wrongWorkbookRestore.details.error ?? "", /different workbook/u);
    const unchanged = await read.execute("read-after-refusal", { range: "Sheet1!A1", mode: "csv" });
    assert.match(firstText(unchanged), /Name/u);
  });

  const auditEntries = await audit.list();
  assert.equal(auditEntries.some((entry) => entry.toolCallId === "write-blank" && entry.workbookId === "url_sha256:contract-book"), true);
  assert.equal(auditEntries.some((entry) => entry.toolCallId === "write-blocked" && entry.blocked), true);
});

void test("unsupported workbook capability returns a clear tool result instead of throwing", async () => {
  const tools = composeCoreToolsForHost((name: CoreToolName) => {
    if (name === "workbook_history") return createWorkbookHistoryTool();
    return {
      name,
      label: name,
      description: name,
      parameters: createWorkbookHistoryTool().parameters,
      execute: () => Promise.resolve({ content: [], details: undefined }),
    };
  }, "wps");
  const history = tools.find((tool) => tool.name === "workbook_history");
  assert.ok(history);
  const result = await history.execute("unsupported-history", { action: "list" });
  assert.match(firstText(result), /not yet supported on WPS Spreadsheets/u);
  assert.deepEqual(result.details, {
    code: "unsupported_host_tool",
    hostKind: "wps",
    toolName: "workbook_history",
  });
});
