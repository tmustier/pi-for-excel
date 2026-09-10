import assert from "node:assert/strict";
import { test } from "node:test";

import { MAX_RECOVERY_CELLS, WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import { captureValueDataRange } from "../src/workbook/recovery/structure-state.ts";
import { type RecoveryModifyStructureState } from "../src/workbook/recovery-states.ts";
import {
  createInMemorySettingsStore,
  findSnapshotById,
  withoutUndefined,
} from "./fixtures/recovery-log.ts";

void test("captureValueDataRange short-circuits oversized captures before loading cell grids", async () => {
  const loadCalls: Array<string | string[]> = [];

  const usedRange = {
    isNullObject: false,
    address: "Sheet1!A1:CV201",
    rowCount: 201,
    columnCount: 100,
    values: [] as DynamicValue[][],
    formulas: [] as DynamicValue[][],
    load: (propertyNames: string | string[]): void => {
      loadCalls.push(propertyNames);
    },
  };

  const targetRange = {
    getUsedRangeOrNullObject: (_valuesOnly?: boolean): typeof usedRange => usedRange,
  };

  const context = {
    sync: (): Promise<DynamicValue> => Promise.resolve(),
  };

  const capture = await captureValueDataRange(context, targetRange, MAX_RECOVERY_CELLS);

  assert.equal(capture.status, "too_large");
  assert.ok(capture.cellCount > MAX_RECOVERY_CELLS);
  assert.deepEqual(loadCalls, [["isNullObject", "address", "rowCount", "columnCount"]]);
});

void test("captureValueDataRange captures in-range value/formula payloads", async () => {
  const usedRange = {
    isNullObject: false,
    address: "Sheet1!B2:C3",
    rowCount: 2,
    columnCount: 2,
    values: [[1, 2], [3, 4]],
    formulas: [["", ""], ["", ""]],
    load: (_propertyNames: string | string[]): void => {},
  };

  const targetRange = {
    getUsedRangeOrNullObject: (_valuesOnly?: boolean): typeof usedRange => usedRange,
  };

  const context = {
    sync: (): Promise<DynamicValue> => Promise.resolve(),
  };

  const capture = await captureValueDataRange(context, targetRange, MAX_RECOVERY_CELLS);

  assert.equal(capture.status, "captured");
  assert.deepEqual(capture.dataRange, {
    address: "B2:C3",
    rowCount: 2,
    columnCount: 2,
    values: [[1, 2], [3, 4]],
    formulas: [["", ""], ["", ""]],
  });
});

void test("restore round-trips extended modify-structure kinds", async () => {
  const scenarios: ReadonlyArray<{
    name: string;
    address: string;
    changedCount: number;
    targetState: RecoveryModifyStructureState;
    currentState: RecoveryModifyStructureState;
  }> = [
    {
      name: "sheet_absent",
      address: "Draft",
      changedCount: 1,
      targetState: {
        kind: "sheet_absent",
        sheetId: "sheet-draft",
        sheetName: "Draft",
      },
      currentState: {
        kind: "sheet_present",
        sheetId: "sheet-draft",
        sheetName: "Draft",
        position: 1,
        visibility: "Visible",
      },
    },
    {
      name: "sheet_present",
      address: "Backlog",
      changedCount: 1,
      targetState: {
        kind: "sheet_present",
        sheetId: "sheet-backlog",
        sheetName: "Backlog",
        position: 3,
        visibility: "Hidden",
      },
      currentState: {
        kind: "sheet_absent",
        sheetId: "sheet-backlog",
        sheetName: "Backlog",
      },
    },
    {
      name: "sheet_present_with_data",
      address: "Archive",
      changedCount: 1,
      targetState: {
        kind: "sheet_present",
        sheetId: "sheet-archive",
        sheetName: "Archive",
        position: 4,
        visibility: "Visible",
        dataRange: {
          address: "A1:B2",
          rowCount: 2,
          columnCount: 2,
          values: [["Year", "Value"], [2024, 100]],
          formulas: [["", ""], ["", ""]],
        },
      },
      currentState: {
        kind: "sheet_absent",
        sheetId: "sheet-archive",
        sheetName: "Archive",
        allowDataDelete: true,
      },
    },
    {
      name: "rows_absent",
      address: "Data!8:9",
      changedCount: 2,
      targetState: {
        kind: "rows_absent",
        sheetId: "sheet-grid",
        sheetName: "Data",
        position: 8,
        count: 2,
        allowDataDelete: true,
      },
      currentState: {
        kind: "rows_present",
        sheetId: "sheet-grid",
        sheetName: "Data",
        position: 8,
        count: 2,
        dataRange: {
          address: "B8:C9",
          rowCount: 2,
          columnCount: 2,
          values: [[10, 20], [30, 40]],
          formulas: [["", ""], ["", ""]],
        },
      },
    },
    {
      name: "rows_present",
      address: "Data!15:16",
      changedCount: 2,
      targetState: {
        kind: "rows_present",
        sheetId: "sheet-grid",
        sheetName: "Data",
        position: 15,
        count: 2,
        dataRange: {
          address: "A15:B16",
          rowCount: 2,
          columnCount: 2,
          values: [["A", "B"], ["C", "D"]],
          formulas: [["", ""], ["", ""]],
        },
      },
      currentState: {
        kind: "rows_absent",
        sheetId: "sheet-grid",
        sheetName: "Data",
        position: 15,
        count: 2,
        allowDataDelete: true,
      },
    },
    {
      name: "columns_absent",
      address: "Data!C:D",
      changedCount: 2,
      targetState: {
        kind: "columns_absent",
        sheetId: "sheet-grid",
        sheetName: "Data",
        position: 3,
        count: 2,
        allowDataDelete: true,
      },
      currentState: {
        kind: "columns_present",
        sheetId: "sheet-grid",
        sheetName: "Data",
        position: 3,
        count: 2,
        dataRange: {
          address: "C2:D4",
          rowCount: 3,
          columnCount: 2,
          values: [["Q1", "Q2"], ["Q3", "Q4"], ["Q5", "Q6"]],
          formulas: [["", ""], ["", ""], ["", ""]],
        },
      },
    },
    {
      name: "columns_present",
      address: "Data!F:G",
      changedCount: 2,
      targetState: {
        kind: "columns_present",
        sheetId: "sheet-grid",
        sheetName: "Data",
        position: 6,
        count: 2,
        dataRange: {
          address: "F1:G2",
          rowCount: 2,
          columnCount: 2,
          values: [[1, 2], [3, 4]],
          formulas: [["", ""], ["", ""]],
        },
      },
      currentState: {
        kind: "columns_absent",
        sheetId: "sheet-grid",
        sheetName: "Data",
        position: 6,
        count: 2,
        allowDataDelete: true,
      },
    },
  ];

  for (let index = 0; index < scenarios.length; index += 1) {
    const scenario = scenarios[index];
    if (!scenario) {
      throw new Error("Expected structure scenario.");
    }

    const settingsStore = createInMemorySettingsStore();

    const workbookContext: WorkbookContext = {
      workbookId: `url_sha256:workbook-structure-roundtrip-${index}`,
      workbookName: "StructureRoundtrip.xlsx",
      source: "document.url",
    };

    let idCounter = 0;
    const createId = (): string => {
      idCounter += 1;
      return `snap-structure-roundtrip-${index}-${idCounter}`;
    };

    let appliedAddress = "";
    let appliedState: RecoveryModifyStructureState | null = null;

    const log = new WorkbookRecoveryLog({
      getSettingsStore: () => Promise.resolve(settingsStore),
      getWorkbookContext: () => Promise.resolve(workbookContext),
      now: () => 1700000001960 + index,
      createId,
      applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
      applyModifyStructureSnapshot: (address, state) => {
        appliedAddress = address;
        appliedState = state;
        return Promise.resolve(scenario.currentState);
      },
    });

    const appended = await log.appendModifyStructure({
      toolName: "modify_structure",
      toolCallId: `call-structure-roundtrip-${scenario.name}`,
      address: scenario.address,
      changedCount: scenario.changedCount,
      modifyStructureState: scenario.targetState,
    });

    assert.ok(appended, `Expected appended checkpoint for ${scenario.name}.`);
    if (!appended) {
      throw new Error(`Expected appended checkpoint for ${scenario.name}.`);
    }

    const restored = await log.restore(appended.id);

    assert.equal(restored.address, scenario.address, `Expected restored address for ${scenario.name}.`);
    assert.equal(
      restored.restoredSnapshotId,
      appended.id,
      `Expected restored snapshot id for ${scenario.name}.`,
    );
    assert.equal(appliedAddress, scenario.address, `Expected apply address for ${scenario.name}.`);
    assert.deepEqual(appliedState, scenario.targetState, `Expected target state for ${scenario.name}.`);

    const snapshots = await log.listForCurrentWorkbook(10);
    const inverse = restored.inverseSnapshotId
      ? findSnapshotById(snapshots, restored.inverseSnapshotId)
      : null;

    assert.ok(inverse, `Expected inverse snapshot for ${scenario.name}.`);
    assert.equal(inverse?.snapshotKind, "modify_structure_state", `Expected structure kind for ${scenario.name}.`);
    assert.equal(inverse?.restoredFromSnapshotId, appended.id, `Expected inverse source for ${scenario.name}.`);
    assert.deepEqual(
      inverse?.modifyStructureState,
      scenario.currentState,
      `Expected inverse state for ${scenario.name}.`,
    );
  }
});
