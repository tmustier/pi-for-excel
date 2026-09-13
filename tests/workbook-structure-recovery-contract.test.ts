import assert from "node:assert/strict";
import { test } from "node:test";

import { createModifyStructureTool } from "../src/tools/modify-structure.ts";
import { createWorkbookHistoryTool } from "../src/tools/workbook-history.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import {
  createInMemorySettingsStore,
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

void test("restoring an inconsistent row backup leaves workbook structure unchanged", async () => {
  const settings = createInMemorySettingsStore();
  await settings.set(RECOVERY_SETTING_KEY, {
    version: 1,
    snapshots: [{
      id: "rows-with-invalid-address-shape",
      at: 1,
      toolName: "modify_structure",
      toolCallId: "delete-rows",
      address: "Data!5:5",
      workbookId: workbook.workbookId,
      snapshotKind: "modify_structure_state",
      modifyStructureState: {
        kind: "rows_present",
        sheetId: "sheet-data",
        sheetName: "Data",
        position: 5,
        count: 1,
        dataRange: {
          address: "A5:B6",
          rowCount: 1,
          columnCount: 2,
          values: [[10, 20]],
          formulas: [["", ""]],
        },
      },
    }],
  });

  let insertedRows = 0;
  const sheet = {
    id: "sheet-data",
    name: "Data",
    visibility: "Visible",
    position: 0,
    isNullObject: false,
    load: (_properties: string | string[]) => undefined,
    getRange: (address: string) => ({
      rowCount: address === "A5:B6" ? 2 : 1,
      columnCount: address === "A5:B6" ? 2 : 16_384,
      load: (_properties: string | string[]) => undefined,
      insert: (direction: string) => {
        assert.equal(direction, "Down");
        insertedRows += 1;
      },
    }),
  };
  const context = {
    workbook: {
      worksheets: {
        getItemOrNullObject: (reference: string) => {
          assert.equal(reference, "sheet-data");
          return sheet;
        },
      },
    },
    sync: () => Promise.resolve(),
  };

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settings),
    getWorkbookContext: () => Promise.resolve(workbook),
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });
  const history = createWorkbookHistoryTool({
    getRecoveryLog: () => log,
    appendAuditEntry: () => Promise.resolve(),
  });

  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <TResult>(callback: (host: unknown) => Promise<TResult>): Promise<TResult> => callback(context),
  });

  try {
    const result = await history.execute("restore-invalid-rows", {
      action: "restore",
      snapshot_id: "rows-with-invalid-address-shape",
    });

    assert.match(firstText(result), /^Error:/u);
    assert.equal(insertedRows, 0);
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
});

void test("restoring a row backup that targets a different row leaves workbook structure unchanged", async () => {
  const settings = createInMemorySettingsStore();
  await settings.set(RECOVERY_SETTING_KEY, {
    version: 1,
    snapshots: [{
      id: "rows-with-displaced-data",
      at: 1,
      toolName: "modify_structure",
      toolCallId: "delete-rows",
      address: "Data!5:5",
      workbookId: workbook.workbookId,
      snapshotKind: "modify_structure_state",
      modifyStructureState: {
        kind: "rows_present",
        sheetId: "sheet-data",
        sheetName: "Data",
        position: 5,
        count: 1,
        dataRange: {
          address: "A6:B6",
          rowCount: 1,
          columnCount: 2,
          values: [[10, 20]],
          formulas: [["", ""]],
        },
      },
    }],
  });

  let insertedRows = 0;
  const sheet = {
    id: "sheet-data",
    name: "Data",
    visibility: "Visible",
    position: 0,
    isNullObject: false,
    load: (_properties: string | string[]) => undefined,
    getRange: (_address: string) => ({
      rowCount: 1,
      columnCount: 2,
      load: (_properties: string | string[]) => undefined,
      insert: () => {
        insertedRows += 1;
      },
    }),
  };
  const context = {
    workbook: {
      worksheets: {
        getItemOrNullObject: (_reference: string) => sheet,
      },
    },
    sync: () => Promise.resolve(),
  };
  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settings),
    getWorkbookContext: () => Promise.resolve(workbook),
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });
  const history = createWorkbookHistoryTool({
    getRecoveryLog: () => log,
    appendAuditEntry: () => Promise.resolve(),
  });

  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <TResult>(callback: (host: unknown) => Promise<TResult>): Promise<TResult> => callback(context),
  });

  try {
    const result = await history.execute("restore-displaced-rows", {
      action: "restore",
      snapshot_id: "rows-with-displaced-data",
    });

    assert.match(firstText(result), /^Error:/u);
    assert.equal(insertedRows, 0);
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
});

void test("duplicating a sheet to an existing name leaves the workbook unchanged", async () => {
  let copiesCreated = 0;
  const existing = {
    isNullObject: false,
    load: (_properties: string | string[]) => undefined,
  };
  const source = {
    copy: (position: string) => {
      assert.equal(position, "End");
      copiesCreated += 1;
      return {
        id: "copied-sheet",
        get name() {
          return "Source (2)";
        },
        set name(_value: string) {
          throw new Error("A worksheet named Existing already exists.");
        },
        load: (_properties: string | string[]) => undefined,
      };
    },
  };
  const context = {
    workbook: {
      worksheets: {
        getItem: (name: string) => {
          assert.equal(name, "Source");
          return source;
        },
        getItemOrNullObject: (name: string) => {
          assert.equal(name, "Existing");
          return existing;
        },
      },
    },
    sync: () => Promise.resolve(),
  };

  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <TResult>(callback: (host: unknown) => Promise<TResult>): Promise<TResult> => callback(context),
  });

  try {
    const result = await createModifyStructureTool().execute("duplicate-existing-name", {
      action: "duplicate_sheet",
      sheet: "Source",
      new_name: "Existing",
    });

    assert.match(firstText(result), /^Error:/u);
    assert.equal(copiesCreated, 0);
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
});

void test("a completed sheet duplicate is reported as success when backup inspection fails", async () => {
  let copiesCreated = 0;
  const copy = {
    id: "copied-sheet",
    name: "Source (2)",
    load: (_properties: string | string[]) => undefined,
    getUsedRangeOrNullObject: () => {
      throw new Error("Used-range inspection unavailable");
    },
  };
  const source = {
    copy: () => {
      copiesCreated += 1;
      return copy;
    },
  };
  const context = {
    workbook: {
      worksheets: {
        getItem: (_name: string) => source,
      },
    },
    sync: () => Promise.resolve(),
  };

  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <TResult>(callback: (host: unknown) => Promise<TResult>): Promise<TResult> => callback(context),
  });

  try {
    const result = await createModifyStructureTool().execute("duplicate-without-inspection", {
      action: "duplicate_sheet",
      sheet: "Source",
    });

    assert.match(firstText(result), /^Duplicated "Source" as "Source \(2\)"\./u);
    assert.match(firstText(result), /Backup not created/u);
    assert.equal(copiesCreated, 1);
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
});

void test("restoring a deleted row reinserts its captured values", async () => {
  const settings = createInMemorySettingsStore();
  await settings.set(RECOVERY_SETTING_KEY, {
    version: 1,
    snapshots: [{
      id: "deleted-row",
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
        dataRange: {
          address: "A5:B5",
          rowCount: 1,
          columnCount: 2,
          values: [[10, 20]],
          formulas: [["", ""]],
        },
      },
    }],
  });

  let insertedRows = 0;
  let restoredValues: unknown[][] = [];
  const rowRange = {
    insert: (direction: string) => {
      assert.equal(direction, "Down");
      insertedRows += 1;
    },
  };
  const dataRange = {
    rowCount: 1,
    columnCount: 2,
    load: (_properties: string | string[]) => undefined,
    set values(value: unknown[][]) {
      restoredValues = value;
    },
  };
  const sheet = {
    id: "sheet-data",
    name: "Data",
    visibility: "Visible",
    position: 0,
    isNullObject: false,
    load: (_properties: string | string[]) => undefined,
    getRange: (address: string) => address === "5:5" ? rowRange : dataRange,
  };
  const context = {
    workbook: {
      worksheets: {
        getItemOrNullObject: (_reference: string) => sheet,
      },
    },
    sync: () => Promise.resolve(),
  };
  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settings),
    getWorkbookContext: () => Promise.resolve(workbook),
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });
  const history = createWorkbookHistoryTool({
    getRecoveryLog: () => log,
    appendAuditEntry: () => Promise.resolve(),
  });

  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <TResult>(callback: (host: unknown) => Promise<TResult>): Promise<TResult> => callback(context),
  });

  try {
    const result = await history.execute("restore-deleted-row", {
      action: "restore",
      snapshot_id: "deleted-row",
    });

    assert.match(firstText(result), /^✅ Restored backup/u);
    assert.equal(insertedRows, 1);
    assert.deepEqual(restoredValues, [[10, 20]]);
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
});
