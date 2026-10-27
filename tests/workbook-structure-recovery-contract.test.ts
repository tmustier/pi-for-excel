import assert from "node:assert/strict";
import { test } from "node:test";

import { createModifyStructureTool } from "../src/tools/modify-structure.ts";
import { createWorkbookHistoryTool } from "../src/tools/workbook-history.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import type { RecoveryStructureValueRangeState } from "../src/workbook/recovery-states.ts";
import {
  createInMemorySettingsStore,
  type InMemorySettingsStore,
  RECOVERY_SETTING_KEY,
} from "./fixtures/recovery-log.ts";

const workbook: WorkbookContext = {
  workbookId: "url_sha256:structure-safety",
  workbookName: "Structure safety.xlsx",
  source: "document.url",
};

function firstText(result: { content: Array<{ type: string; text?: string }> }): string {
  const block = result.content[0];
  if (!block || block.type !== "text" || typeof block.text !== "string") {
    throw new Error("Expected a text tool result.");
  }
  return block.text;
}

function createHistory(settings: InMemorySettingsStore): ReturnType<typeof createWorkbookHistoryTool> {
  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settings),
    getWorkbookContext: () => Promise.resolve(workbook),
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });
  return createWorkbookHistoryTool({
    getRecoveryLog: () => log,
    appendAuditEntry: () => Promise.resolve(),
  });
}

async function storeDeletedRow(
  settings: InMemorySettingsStore,
  id: string,
  dataRange: RecoveryStructureValueRangeState,
): Promise<void> {
  await settings.set(RECOVERY_SETTING_KEY, {
    version: 1,
    snapshots: [{
      id,
      at: 1,
      toolName: "modify_structure",
      toolCallId: "delete-row",
      address: "Data!5:5",
      workbookId: workbook.workbookId,
      snapshotKind: "modify_structure_state",
      modifyStructureState: {
        kind: "rows_present",
        sheetId: "sheet-data",
        sheetName: "Data",
        position: 5,
        count: 1,
        dataRange,
      },
    }],
  });
}

async function withExcel<TContext, TResult>(
  context: TContext,
  action: () => Promise<TResult>,
): Promise<TResult> {
  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: (callback: (host: TContext) => Promise<TResult>): Promise<TResult> => callback(context),
  });

  try {
    return await action();
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
}

void test("an inconsistent persisted row backup never reaches the workbook host", async () => {
  const settings = createInMemorySettingsStore();
  await storeDeletedRow(settings, "rows-with-invalid-address-shape", {
    address: "A5:B6",
    rowCount: 1,
    columnCount: 2,
    values: [[10, 20]],
    formulas: [["", ""]],
  });

  let workbookTouched = false;
  const context = {
    workbook: {
      worksheets: {
        getItemOrNullObject: () => {
          workbookTouched = true;
          throw new Error("Invalid backup reached the workbook host.");
        },
      },
    },
  };

  const result = await withExcel(context, () => createHistory(settings).execute("restore-invalid-rows", {
    action: "restore",
    snapshot_id: "rows-with-invalid-address-shape",
  }));

  assert.match(firstText(result), /^Error:/u);
  assert.equal(workbookTouched, false);
});

void test("a persisted row backup for a different row never reaches the workbook host", async () => {
  const settings = createInMemorySettingsStore();
  await storeDeletedRow(settings, "rows-with-displaced-data", {
    address: "A6:B6",
    rowCount: 1,
    columnCount: 2,
    values: [[10, 20]],
    formulas: [["", ""]],
  });

  let workbookTouched = false;
  const context = {
    workbook: {
      worksheets: {
        getItemOrNullObject: () => {
          workbookTouched = true;
          throw new Error("Invalid backup reached the workbook host.");
        },
      },
    },
  };

  const result = await withExcel(context, () => createHistory(settings).execute("restore-displaced-rows", {
    action: "restore",
    snapshot_id: "rows-with-displaced-data",
  }));

  assert.match(firstText(result), /^Error:/u);
  assert.equal(workbookTouched, false);
});

void test("duplicating a sheet to an existing name leaves the workbook unchanged", async () => {
  let copiesCreated = 0;
  const context = {
    workbook: {
      worksheets: {
        getItem: () => ({
          copy: () => {
            copiesCreated += 1;
          },
        }),
        getItemOrNullObject: () => ({
          isNullObject: false,
          load: () => undefined,
        }),
      },
    },
    sync: () => Promise.resolve(),
  };

  const result = await withExcel(context, () => createModifyStructureTool().execute("duplicate-existing-name", {
    action: "duplicate_sheet",
    sheet: "Source",
    new_name: "Existing",
  }));

  assert.equal(result.isError, true);
  assert.match(firstText(result), /^Error \(duplicate_sheet\) — the workbook was not changed\./u);
  assert.equal(copiesCreated, 0);
});

void test("a completed sheet duplicate stays successful when backup inspection fails", async () => {
  let copiesCreated = 0;
  const context = {
    workbook: {
      worksheets: {
        getItem: () => ({
          copy: () => {
            copiesCreated += 1;
            return {
              id: "copied-sheet",
              name: "Source (2)",
              load: () => undefined,
              getUsedRangeOrNullObject: () => {
                throw new Error("Used-range inspection unavailable");
              },
            };
          },
        }),
      },
    },
    sync: () => Promise.resolve(),
  };

  const result = await withExcel(context, () => createModifyStructureTool().execute("duplicate-without-inspection", {
    action: "duplicate_sheet",
    sheet: "Source",
  }));

  assert.match(firstText(result), /^Duplicated "Source" as "Source \(2\)"\./u);
  assert.match(firstText(result), /No backup was saved.*Used-range inspection unavailable/u);
  assert.equal(copiesCreated, 1);
});

void test("restoring a deleted row reinserts its captured values", async () => {
  const settings = createInMemorySettingsStore();
  await storeDeletedRow(settings, "deleted-row", {
    address: "A5:B5",
    rowCount: 1,
    columnCount: 2,
    values: [[10, 20]],
    formulas: [["", ""]],
  });

  let insertedRows = 0;
  let restoredValues: Array<Array<string | number>> = [];
  const rowRange = {
    insert: (direction: string) => {
      assert.equal(direction, "Down");
      insertedRows += 1;
    },
  };
  const dataRange = {
    rowCount: 1,
    columnCount: 2,
    load: () => undefined,
    set values(value: Array<Array<string | number>>) {
      restoredValues = value;
    },
  };
  const sheet = {
    id: "sheet-data",
    name: "Data",
    visibility: "Visible",
    position: 0,
    isNullObject: false,
    load: () => undefined,
    getRange: (address: string) => address === "5:5" ? rowRange : dataRange,
  };
  const context = {
    workbook: {
      worksheets: {
        getItemOrNullObject: () => sheet,
      },
    },
    sync: () => Promise.resolve(),
  };

  const result = await withExcel(context, () => createHistory(settings).execute("restore-deleted-row", {
    action: "restore",
    snapshot_id: "deleted-row",
  }));

  assert.match(firstText(result), /^✅ Restored backup/u);
  assert.equal(insertedRows, 1);
  assert.deepEqual(restoredValues, [[10, 20]]);
});

/**
 * A two-column sheet for delete_rows. Rows are hidden by hand or by a filter. As in
 * Excel on the web, deleting deletes hand-hidden rows but keeps filtered-out ones.
 */
function rowSheetHost(rows: Array<{ values: string[]; hidden?: "manual" | "filter" }>): {
  context: unknown;
  rows: Array<{ values: string[]; hidden?: "manual" | "filter" }>;
} {
  const parse = (address: string): [number, number] => {
    const [start = "1", end = start] = address.split(":").map((part) => part.replace(/[A-Z]/gu, ""));
    return [Number(start), Number(end)];
  };
  const rowAt = (index: number) => rows[index - 1] ?? { values: ["", ""] };
  const rowsIn = (first: number, last: number) => Array.from({ length: last - first + 1 }, (_, i) => rowAt(first + i));
  const filtered = rows.some((row) => row.hidden === "filter");
  const rangeAt = (address: string) => {
    const [first, last] = parse(address);
    return {
      load: (): void => {},
      get rowHidden(): boolean | null {
        const states = new Set(rowsIn(first, last).map((row) => row.hidden !== undefined));
        return states.size === 1 ? [...states][0] ?? null : null;
      },
      get values(): string[][] { return rowsIn(first, last).map((row) => row.values); },
      getTables: () => ({ load: (): void => {}, items: [{ name: "Sales", autoFilter: { isDataFiltered: filtered } }] }),
      getUsedRangeOrNullObject: () => {
        const used = rowsIn(first, last).filter((row) => row.values.some((value) => value !== ""));
        return {
          load: (): void => {},
          isNullObject: used.length === 0,
          address: `Data!A${first}:B${last}`,
          rowCount: last - first + 1,
          columnCount: 2,
          values: rowsIn(first, last).map((row) => row.values),
          formulas: rowsIn(first, last).map((row) => row.values),
        };
      },
      delete: (): void => {
        const kept = rows.slice(first - 1, last).filter((row) => row.hidden === "filter");
        rows.splice(first - 1, last - first + 1, ...kept);
      },
    };
  };
  const sheet = {
    id: "sheet-data",
    name: "Data",
    load: (): void => {},
    getRange: rangeAt,
    getUsedRangeOrNullObject: () => ({ load: (): void => {}, isNullObject: false, columnIndex: 0, columnCount: 2 }),
    autoFilter: { load: (): void => {}, isDataFiltered: false },
  };
  return {
    rows,
    context: {
      workbook: { worksheets: { getItem: () => sheet, getActiveWorksheet: () => sheet } },
      sync: (): Promise<void> => Promise.resolve(),
    },
  };
}

function regionRows(hidden: "manual" | "filter"): Array<{ values: string[]; hidden?: "manual" | "filter" }> {
  return [
    { values: ["Region", "Code"] },
    { values: ["North", "a"] },
    { values: ["South", "b"], hidden },
    { values: ["East", "c"] },
  ];
}

void test("delete_rows over a hand-hidden row says the hidden row was deleted, with its data", async () => {
  const host = rowSheetHost(regionRows("manual"));
  const result = await withExcel(host.context, () => createModifyStructureTool().execute("delete-hidden", {
    action: "delete_rows", sheet: "Data", position: 2, count: 2,
  }));

  assert.equal(result.isError, undefined, firstText(result));
  assert.match(firstText(result), /^Deleted 2 row\(s\)/u);
  assert.match(firstText(result), /⚠️ 1 deleted row\(s\) were hidden: row 3\. The rows were in table \*\*Sales\*\*\. Deleted hidden data: row 3 \("South", "b"\)/u);
  assert.equal(result.details.hiddenRowsDeleted, 1);
  assert.deepEqual(host.rows.map((row) => row.values[0]), ["Region", "East"]);
});

void test("delete_rows in a filtered range says Excel kept the filtered-out row and deleted fewer rows", async () => {
  const host = rowSheetHost(regionRows("filter"));
  const result = await withExcel(host.context, () => createModifyStructureTool().execute("delete-filtered", {
    action: "delete_rows", sheet: "Data", position: 2, count: 2,
  }));

  assert.match(firstText(result), /^Deleted 1 of 2 row\(s\) starting at row 2/u);
  assert.match(firstText(result), /Excel kept 1 hidden row\(s\) \(a filter is on\).*originally row 3\) now start at row 2/u);
  assert.doesNotMatch(firstText(result), /Deleted hidden data/u);
  assert.equal(result.details.hiddenRowsDeleted, 0);
  // A backup of both rows would bring "South" back twice.
  assert.equal(result.details.recovery?.status, "not_available");
  assert.deepEqual(host.rows.map((row) => row.values[0]), ["Region", "South", "East"]);
});

void test("delete_rows over visible rows only adds no hidden-row warning", async () => {
  const host = rowSheetHost(regionRows("manual"));
  const result = await withExcel(host.context, () => createModifyStructureTool().execute("delete-visible", {
    action: "delete_rows", sheet: "Data", position: 2,
  }));
  assert.doesNotMatch(firstText(result), /hidden/u);
  assert.equal(result.details.hiddenRowsDeleted, undefined);
});
