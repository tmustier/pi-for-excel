import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentTool } from "@earendil-works/pi-agent-core";

import { parseCell } from "../src/excel/helpers.ts";
import { applyToolOutputTruncation } from "../src/tools/output-truncation.ts";
import { createReadRangeTool } from "../src/tools/read-range.ts";

type Cell = { value: string | number | boolean; formula?: string; r1c1?: string; text?: string; format?: string };

interface SheetFixture {
  name: string;
  /** Sparse cells by A1 address. */
  cells?: Record<string, Cell>;
  /** Dense generator for big sheets: value at (rowIndex, columnIndex), 0-based. */
  dense?: { rows: number; columns: number; value: (row: number, column: number) => string | number };
  protected?: boolean;
  merged?: string;
  hiddenRows?: number[];
}

/** Excel on the web refuses a response this big; one load of more cells throws. */
const MAX_CELLS_PER_LOAD = 50_000;
const SHEET_ROWS = 1_048_576;
const SHEET_COLUMNS = 16_384;

function parseAddress(address: string): { rowIndex: number; columnIndex: number; rowCount: number; columnCount: number } {
  const clean = address.replace(/\$/gu, "").toUpperCase();
  const [start = "", end = start] = clean.split(":");
  if (/^[A-Z]+$/u.test(start)) {
    const first = parseCell(`${start}1`).col;
    const last = parseCell(`${end}1`).col;
    return { rowIndex: 0, columnIndex: first, rowCount: SHEET_ROWS, columnCount: last - first + 1 };
  }
  if (/^\d+$/u.test(start)) {
    return { rowIndex: Number(start) - 1, columnIndex: 0, rowCount: Number(end) - Number(start) + 1, columnCount: SHEET_COLUMNS };
  }
  const a = parseCell(start);
  const b = parseCell(end);
  return { rowIndex: a.row - 1, columnIndex: a.col, rowCount: b.row - a.row + 1, columnCount: b.col - a.col + 1 };
}

function makeSheet(fixture: SheetFixture, fail: (error: Error) => void) {
  const cells = new Map(Object.entries(fixture.cells ?? {}).map(([address, cell]) => {
    const { rowIndex, columnIndex } = parseAddress(address);
    return [`${rowIndex},${columnIndex}`, cell] as const;
  }));
  const cellAt = (row: number, column: number): Cell | undefined => {
    const sparse = cells.get(`${row},${column}`);
    if (sparse) return sparse;
    if (fixture.dense && row < fixture.dense.rows && column < fixture.dense.columns) {
      return { value: fixture.dense.value(row, column) };
    }
    return undefined;
  };
  let usedRows = fixture.dense?.rows ?? 0;
  let usedColumns = fixture.dense?.columns ?? 0;
  for (const key of cells.keys()) {
    const [row = 0, column = 0] = key.split(",").map(Number);
    usedRows = Math.max(usedRows, row + 1);
    usedColumns = Math.max(usedColumns, column + 1);
  }
  const hidden = new Set(fixture.hiddenRows ?? []);
  const nullObject = { isNullObject: true, load: () => undefined };

  const rangeAt = (rowIndex: number, columnIndex: number, rowCount: number, columnCount: number) => {
    const grid = <T>(pick: (cell: Cell | undefined) => T): T[][] => {
      if (rowCount * columnCount > MAX_CELLS_PER_LOAD) {
        throw new Error("The response payload size has exceeded the limit.");
      }
      return Array.from({ length: rowCount }, (_, r) =>
        Array.from({ length: columnCount }, (_, c) => pick(cellAt(rowIndex + r, columnIndex + c))));
    };
    const end = `${String.fromCharCode(65 + columnIndex + columnCount - 1)}${rowIndex + rowCount}`;
    const address = `${fixture.name}!${String.fromCharCode(65 + columnIndex)}${rowIndex + 1}:${end}`;
    return {
      isNullObject: false,
      rowIndex,
      columnIndex,
      rowCount,
      columnCount,
      address,
      load(properties: string) {
        if (/values|text|formulas/u.test(properties) && rowCount * columnCount > MAX_CELLS_PER_LOAD) {
          fail(new Error("The response payload size has exceeded the limit."));
        }
      },
      get values() { return grid((cell) => cell?.value ?? ""); },
      get formulas() { return grid((cell) => cell?.formula ?? cell?.value ?? ""); },
      get formulasR1C1() { return grid((cell) => cell?.r1c1 ?? cell?.formula ?? cell?.value ?? ""); },
      get numberFormat() { return grid((cell) => cell?.format ?? "General"); },
      get text() { return grid((cell) => cell?.text ?? (cell === undefined ? "" : String(cell.value))); },
      format: { protection: { locked: true, load: () => undefined } },
      getMergedAreasOrNullObject: () => fixture.merged
        ? { isNullObject: false, address: `${fixture.name}!${fixture.merged}`, areaCount: 1, load: () => undefined }
        : nullObject,
      getSpecialCellsOrNullObject: (kind: string) => {
        if (kind !== "Visible") return nullObject;
        const areas: string[] = [];
        for (let r = rowIndex; r < rowIndex + rowCount; r += 1) {
          if (!hidden.has(r + 1)) areas.push(`${fixture.name}!A${r + 1}:${end.replace(/\d+$/u, String(r + 1))}`);
        }
        return { isNullObject: areas.length === 0, address: areas.join(","), load: () => undefined };
      },
    };
  };

  return {
    name: fixture.name,
    load: () => undefined,
    isNullObject: false,
    protection: { protected: fixture.protected ?? false, load: () => undefined },
    tables: { items: [], load: () => undefined },
    autoFilter: { enabled: false, isDataFiltered: false, load: () => undefined, getRangeOrNullObject: () => nullObject },
    comments: { items: [], load: () => undefined },
    getRange: (address: string) => {
      const block = parseAddress(address);
      return {
        ...rangeAt(block.rowIndex, block.columnIndex, Math.min(block.rowCount, 5), Math.min(block.columnCount, 5)),
        rowCount: block.rowCount,
        columnCount: block.columnCount,
        getSpillingToRangeOrNullObject: () => nullObject,
      };
    },
    getRangeByIndexes: rangeAt,
    getUsedRangeOrNullObject: () => usedRows === 0 ? nullObject : rangeAt(0, 0, usedRows, usedColumns),
  };
}

async function readRange(
  fixtures: SheetFixture[],
  params: { range: string; mode?: "compact" | "csv" | "detailed" },
  options: { activeSheet?: string; calculationMode?: string; tool?: AgentTool } = {},
): Promise<string> {
  let pendingError: Error | undefined;
  const sheets = fixtures.map((fixture) => makeSheet(fixture, (error) => { pendingError = error; }));
  const context = {
    workbook: {
      application: { calculationMode: options.calculationMode ?? "Automatic", load: () => undefined },
      worksheets: {
        items: sheets,
        load: () => undefined,
        getActiveWorksheet: () => sheets.find((sheet) => sheet.name === options.activeSheet) ?? sheets[0],
        getItemOrNullObject: (name: string) =>
          sheets.find((sheet) => sheet.name.toLowerCase() === name.toLowerCase()) ?? { isNullObject: true, load: () => undefined },
      },
    },
    sync: () => {
      const error = pendingError;
      pendingError = undefined;
      return error ? Promise.reject(error) : Promise.resolve();
    },
  };

  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <T>(callback: (host: unknown) => Promise<T>): Promise<T> => callback(context),
  });
  try {
    const tool = options.tool ?? createReadRangeTool();
    const result = await tool.execute("read", params);
    const block = result.content[0];
    if (!block || block.type !== "text") throw new Error("Expected text output.");
    return block.text;
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
}

const bigSheet: SheetFixture = {
  name: "Big",
  dense: { rows: 200_001, columns: 5, value: (row, column) => (row === 0 ? `H${column}` : row * 10 + column) },
};

void test("a whole-column read of a 200k-row sheet returns the first page and names the next one", async () => {
  const text = await readRange([{ name: "Small", cells: { A1: { value: "x" } } }, bigSheet], { range: "Big!A:A" });

  assert.doesNotMatch(text, /payload/u);
  assert.match(text, /^\*\*Big!A1:A500\*\* \(500×1\)/u);
  assert.match(text, /clipped to the data: A1:A200001/u);
  assert.match(text, /Next page: read_range "Big!A501:A1000"/u);
  assert.match(text, /\| 500 \| 4990 \|/u);
  assert.doesNotMatch(text, /\| 501 \|/u);
});

void test("a large rectangular read pages by rows and keeps every column", async () => {
  const text = await readRange([bigSheet], { range: "Big!A1:E50001" });

  assert.match(text, /^\*\*Big!A1:E500\*\* \(500×5\)/u);
  assert.match(text, /Next page: read_range "Big!A501:E1000"/u);
});

void test("cells read as displayed, keep the stored value when the display hides it, and stay inside the grid", async () => {
  const text = await readRange([{
    name: "Data",
    cells: {
      A1: { value: 45292, text: "01/01/2024", format: "dd/mm/yyyy" },
      B1: { value: 0.125, text: "12.5%", format: "0.0%" },
      C1: { value: 0 },
      D1: { value: "a|b" },
      E1: { value: "line1\nline2" },
      F1: { value: 98765.4321, text: "£98,765.43", format: "£#,##0.00" },
      G1: { value: 1e16 },
      H1: { value: 0.000001234 },
    },
  }], { range: "Data!A1:H1" });

  const row = text.split("\n").find((line) => line.startsWith("| 1 |"));
  assert.equal(
    row,
    "| 1 | 01/01/2024 [45292] | 12.5% | 0 | a\\|b | line1\\nline2 | £98,765.43 [98765.4321] | 10000000000000000 | 0.000001234 |",
  );
  assert.match(text, /\[brackets\] hold the stored value/u);
});

void test("an address without a sheet says which sheet was read, and an unknown sheet lists the real ones", async () => {
  const fixtures: SheetFixture[] = [
    { name: "Inputs", cells: { A1: { value: "in" } } },
    { name: "Model", cells: { A1: { value: "model" } } },
  ];

  const unqualified = await readRange(fixtures, { range: "A1" }, { activeSheet: "Model" });
  assert.match(unqualified, /^\*\*Model!A1\*\*/u);
  assert.match(unqualified, /Sheet: "Model" \(the active sheet; the address named no sheet\)/u);
  assert.match(unqualified, /\| 1 \| model \|/u);

  const qualified = await readRange(fixtures, { range: "Inputs!A1" }, { activeSheet: "Model" });
  assert.doesNotMatch(qualified, /active sheet/u);

  const missing = await readRange(fixtures, { range: "Input!A1" });
  assert.match(missing, /Sheet "Input" not found\. Sheets in this workbook: "Inputs", "Model"\./u);
});

void test("sheet state, formulas and errors come before the table and survive output truncation", async () => {
  const cells: Record<string, Cell> = { A1: { value: "Header" } };
  for (let row = 2; row <= 400; row += 1) {
    cells[`B${row}`] = { value: row * 2, formula: `=A${row}*2`, r1c1: "=RC[-1]*2" };
  }
  cells.C400 = { value: "#DIV/0!", formula: "=1/0" };

  // Truncation at 200 lines, the floor for small-window models.
  const tool = applyToolOutputTruncation([createReadRangeTool()], { limits: { maxLines: 200, maxBytes: 8 * 1024 } })[0];
  assert.ok(tool);
  const text = await readRange([{
    name: "Calc",
    cells,
    protected: true,
    merged: "A1:B1",
    hiddenRows: [3, 4, 5, 9],
  }], { range: "Calc!A1:C400" }, { calculationMode: "Manual", tool });

  assert.match(text, /\[Output truncated/u);
  assert.match(text, /Calculation mode is Manual/u);
  assert.match(text, /Sheet is protected: every cell here is locked/u);
  assert.match(text, /Hidden rows \(filtered or hidden\): 3-5, 9/u);
  assert.match(text, /Merged: A1:B1/u);
  // A filled block is one entry, not 399.
  assert.match(text, /\*\*Formulas:\*\* B2:B400: =A2\*2 \(filled, 399 cells\), C400: =1\/0/u);
  assert.match(text, /\*\*Errors:\*\* C400=#DIV\/0!/u);
  assert.ok(text.indexOf("**Errors:**") < text.indexOf("| --- |"), "summaries precede the table");
});
