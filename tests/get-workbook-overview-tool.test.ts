import assert from "node:assert/strict";
import { test } from "node:test";

import { createGetWorkbookOverviewTool } from "../src/tools/get-workbook-overview.ts";

interface SheetFixture {
  name: string;
  /** Used range by zero-based indexes, with its rows of values. */
  used: { rowIndex: number; columnIndex: number; rows: Array<Array<string | number>> };
}

const nullObject = { isNullObject: true, load: () => undefined };

function makeSheet(fixture: SheetFixture, position: number) {
  const { rowIndex, columnIndex, rows } = fixture.used;
  const columnCount = Math.max(...rows.map((row) => row.length));
  return {
    name: fixture.name,
    position,
    visibility: "Visible",
    getUsedRangeOrNullObject: () => ({
      isNullObject: false,
      rowIndex,
      columnIndex,
      rowCount: rows.length,
      columnCount,
      address: `${fixture.name}!X`,
      load: () => undefined,
    }),
    getRangeByIndexes: (row: number, column: number, rowCount: number, count: number) => ({
      load: () => undefined,
      get values() {
        return rows.slice(row - rowIndex, row - rowIndex + rowCount)
          .map((values) => values.slice(column - columnIndex, column - columnIndex + count));
      },
    }),
    // Row 1 of a sheet whose data starts lower down is empty.
    getRange: () => ({ getUsedRangeOrNullObject: () => nullObject }),
    tables: { items: [], load: () => undefined },
    charts: { count: 0, load: () => undefined },
    pivotTables: { getCount: () => ({ value: 0 }) },
    shapes: { items: [], load: () => undefined },
  };
}

async function overview(fixtures: SheetFixture[]): Promise<string> {
  const sheets = fixtures.map(makeSheet);
  const context = {
    workbook: {
      name: "Book.xlsx",
      load: () => undefined,
      application: { calculationMode: "Automatic", load: () => undefined },
      names: { items: [], load: () => undefined },
      worksheets: {
        items: sheets.map((sheet) => ({ ...sheet, protection: { protected: false, load: () => undefined } })),
        load: () => undefined,
      },
    },
    sync: () => Promise.resolve(),
  };
  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", { run: <T>(callback: (host: unknown) => Promise<T>): Promise<T> => callback(context) });
  try {
    const result = await createGetWorkbookOverviewTool().execute("overview", {});
    const block = result.content[0];
    if (!block || block.type !== "text") throw new Error("Expected text output.");
    return block.text;
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
}

void test("headers come from the first row of the used range, and name that row when it is not row 1", async () => {
  const text = await overview([
    { name: "Top", used: { rowIndex: 0, columnIndex: 0, rows: [["ID", "Name"], [1, "a"]] } },
    { name: "Offset", used: { rowIndex: 4, columnIndex: 2, rows: [["Region", "Q1", "Q2"], ["North", 10, 12]] } },
  ]);

  assert.match(text, /\*\*Top\*\* — 2 rows × 2 cols\n {3}Headers: ID, Name\n/u);
  assert.match(text, /\*\*Offset\*\* — 2 rows × 3 cols\n {3}Headers \(row 5\): Region, Q1, Q2/u);
});
