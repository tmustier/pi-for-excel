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
} from "./recovery-log-test-helpers.test.ts";

void test("captureValueDataRange short-circuits oversized captures before loading cell grids", async () => {
  const loadCalls: Array<string | string[]> = [];

  const usedRange = {
    isNullObject: false,
    address: "Sheet1!A1:CV201",
    rowCount: 201,
    columnCount: 100,
    values: [] as unknown[][],
    formulas: [] as unknown[][],
    load: (propertyNames: string | string[]): void => {
      loadCalls.push(propertyNames);
    },
  };

  const targetRange = {
    getUsedRangeOrNullObject: (_valuesOnly?: boolean): typeof usedRange => usedRange,
  };

  const context = {
    sync: (): Promise<unknown> => Promise.resolve(),
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
    sync: (): Promise<unknown> => Promise.resolve(),
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

void test("persisted modify-structure checkpoints retain extended state kinds", async () => {
  const settingsStore = createInMemorySettingsStore();

  const getWorkbookContext = (): Promise<WorkbookContext> => Promise.resolve({
    workbookId: "url_sha256:workbook-structure-persist",
    workbookName: "Ops.xlsx",
    source: "document.url",
  });

  const states: readonly RecoveryModifyStructureState[] = [
    {
      kind: "sheet_absent",
      sheetId: "sheet-added-1",
      sheetName: "Draft",
      allowDataDelete: true,
    },
    {
      kind: "sheet_present",
      sheetId: "sheet-added-2",
      sheetName: "Roadmap",
      position: 2,
      visibility: "Visible",
      dataRange: {
        address: "A1:B2",
        rowCount: 2,
        columnCount: 2,
        values: [["Title", "Owner"], ["Roadmap", "Pi"]],
        formulas: [["", ""], ["", ""]],
      },
    },
    {
      kind: "rows_absent",
      sheetId: "sheet-grid-1",
      sheetName: "Data",
      position: 4,
      count: 2,
      allowDataDelete: true,
    },
    {
      kind: "rows_present",
      sheetId: "sheet-grid-1",
      sheetName: "Data",
      position: 4,
      count: 2,
      dataRange: {
        address: "B4:C5",
        rowCount: 2,
        columnCount: 2,
        values: [[1, 2], [3, 4]],
        formulas: [["", ""], ["", ""]],
      },
    },
    {
      kind: "columns_absent",
      sheetId: "sheet-grid-1",
      sheetName: "Data",
      position: 3,
      count: 1,
      allowDataDelete: true,
    },
    {
      kind: "columns_present",
      sheetId: "sheet-grid-1",
      sheetName: "Data",
      position: 3,
      count: 1,
      dataRange: {
        address: "C2:C4",
        rowCount: 3,
        columnCount: 1,
        values: [["Q1"], ["Q2"], ["Q3"]],
        formulas: [[""], [""], [""]],
      },
    },
  ];

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-structure-persist-${idCounter}`;
  };

  const logA = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
    now: () => 1700000000150,
    createId,
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  for (let index = 0; index < states.length; index += 1) {
    const state = states[index];
    if (!state) {
      throw new Error("Expected structure checkpoint state.");
    }

    const appended = await logA.appendModifyStructure({
      toolName: "modify_structure",
      toolCallId: `call-structure-persist-${index + 1}`,
      address: "Sheet1",
      changedCount: 1,
      modifyStructureState: state,
    });

    assert.ok(appended);
  }

  const logB = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  const entries = await logB.listForCurrentWorkbook(20);
  assert.equal(entries.length, states.length);

  for (let index = 0; index < states.length; index += 1) {
    const expectedState = states[index];
    if (!expectedState) {
      throw new Error("Expected structure checkpoint state.");
    }

    const toolCallId = `call-structure-persist-${index + 1}`;
    const entry = entries.find((snapshot) => snapshot.toolCallId === toolCallId);
    if (!entry) {
      throw new Error(`Expected checkpoint entry for ${toolCallId}.`);
    }

    assert.equal(entry.snapshotKind, "modify_structure_state");
    assert.deepEqual(withoutUndefined(entry.modifyStructureState), withoutUndefined(expectedState));
  }
});
void test("restore applies modify-structure checkpoints and creates inverse checkpoint", async () => {
  const settingsStore = createInMemorySettingsStore();

  const workbookContext: WorkbookContext = {
    workbookId: "url_sha256:workbook-structure",
    workbookName: "Structure.xlsx",
    source: "document.url",
  };

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-structure-${idCounter}`;
  };

  let appliedAddress = "";
  let appliedState: RecoveryModifyStructureState | null = null;

  const restoredState: RecoveryModifyStructureState = {
    kind: "sheet_name",
    sheetId: "sheet-id-1",
    name: "Revenue",
  };

  const currentState: RecoveryModifyStructureState = {
    kind: "sheet_name",
    sheetId: "sheet-id-1",
    name: "Revenue (draft)",
  };

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve(workbookContext),
    now: () => 1700000001900,
    createId,
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
    applyModifyStructureSnapshot: (address, state) => {
      appliedAddress = address;
      appliedState = state;
      return Promise.resolve(currentState);
    },
  });

  const appended = await log.appendModifyStructure({
    toolName: "modify_structure",
    toolCallId: "call-structure",
    address: "Revenue (draft)",
    changedCount: 1,
    modifyStructureState: restoredState,
  });

  assert.ok(appended);

  const restored = await log.restore(appended?.id ?? "");

  assert.equal(restored.address, "Revenue (draft)");
  assert.equal(restored.restoredSnapshotId, appended?.id);
  assert.equal(appliedAddress, "Revenue (draft)");
  assert.deepEqual(appliedState, restoredState);

  const snapshots = await log.listForCurrentWorkbook(10);
  const inverse = restored.inverseSnapshotId
    ? findSnapshotById(snapshots, restored.inverseSnapshotId)
    : null;

  assert.ok(inverse);
  assert.equal(inverse?.toolName, "restore_snapshot");
  assert.equal(inverse?.snapshotKind, "modify_structure_state");
  assert.equal(inverse?.restoredFromSnapshotId, appended?.id);
  assert.deepEqual(inverse?.modifyStructureState, currentState);
});
void test("restore applies row-structure checkpoints and creates inverse checkpoint", async () => {
  const settingsStore = createInMemorySettingsStore();

  const workbookContext: WorkbookContext = {
    workbookId: "url_sha256:workbook-structure-rows",
    workbookName: "StructureRows.xlsx",
    source: "document.url",
  };

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-structure-rows-${idCounter}`;
  };

  let appliedAddress = "";
  let appliedState: RecoveryModifyStructureState | null = null;

  const restoredState: RecoveryModifyStructureState = {
    kind: "rows_absent",
    sheetId: "sheet-id-rows",
    sheetName: "Data",
    position: 4,
    count: 2,
  };

  const currentState: RecoveryModifyStructureState = {
    kind: "rows_present",
    sheetId: "sheet-id-rows",
    sheetName: "Data",
    position: 4,
    count: 2,
  };

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve(workbookContext),
    now: () => 1700000001950,
    createId,
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
    applyModifyStructureSnapshot: (address, state) => {
      appliedAddress = address;
      appliedState = state;
      return Promise.resolve(currentState);
    },
  });

  const appended = await log.appendModifyStructure({
    toolName: "modify_structure",
    toolCallId: "call-structure-rows",
    address: "Data!4:5",
    changedCount: 2,
    modifyStructureState: restoredState,
  });

  assert.ok(appended);

  const restored = await log.restore(appended?.id ?? "");

  assert.equal(restored.address, "Data!4:5");
  assert.equal(restored.restoredSnapshotId, appended?.id);
  assert.equal(appliedAddress, "Data!4:5");
  assert.deepEqual(appliedState, restoredState);

  const snapshots = await log.listForCurrentWorkbook(10);
  const inverse = restored.inverseSnapshotId
    ? findSnapshotById(snapshots, restored.inverseSnapshotId)
    : null;

  assert.ok(inverse);
  assert.equal(inverse?.toolName, "restore_snapshot");
  assert.equal(inverse?.snapshotKind, "modify_structure_state");
  assert.equal(inverse?.restoredFromSnapshotId, appended?.id);
  assert.deepEqual(inverse?.modifyStructureState, currentState);
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
