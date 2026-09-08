import assert from "node:assert/strict";
import { test } from "node:test";

import {
  MAX_RECOVERY_CELLS,
  WorkbookRecoveryLog,
} from "../src/workbook/recovery-log.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import { finalizeMutationOperation } from "../src/tools/mutation/finalize.ts";
import type { RecoveryCheckpointDetails } from "../src/tools/tool-details.ts";
import { RECOVERY_SETTING_KEY, createInMemorySettingsStore } from "./recovery-log-test-helpers.test.ts";

void test("recovery log appends and reloads workbook-scoped snapshots", async () => {
  const settingsStore = createInMemorySettingsStore();

  const getWorkbookContext = (): Promise<WorkbookContext> => Promise.resolve({
    workbookId: "url_sha256:workbook-a",
    workbookName: "Ops.xlsx",
    source: "document.url",
  });

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-${idCounter}`;
  };

  const logA = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
    now: () => 1700000000000,
    createId,
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  const appended = await logA.append({
    toolName: "write_cells",
    toolCallId: "call-1",
    address: "Sheet1!A1",
    changedCount: 1,
    beforeValues: [["before"]],
    beforeFormulas: [["before"]],
  });

  assert.ok(appended);
  assert.equal(appended?.id, "snap-1");

  const entriesA = await logA.listForCurrentWorkbook();
  assert.equal(entriesA.length, 1);
  assert.equal(entriesA[0]?.workbookId, "url_sha256:workbook-a");

  const logB = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  const entriesB = await logB.listForCurrentWorkbook();
  assert.equal(entriesB.length, 1);
  assert.equal(entriesB[0]?.toolCallId, "call-1");
});
void test("append is skipped when workbook identity is unavailable", async () => {
  const settingsStore = createInMemorySettingsStore();

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId: null,
      workbookName: null,
      source: "unknown",
    }),
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  const appended = await log.append({
    toolName: "write_cells",
    toolCallId: "call-null-id",
    address: "Sheet1!A1",
    beforeValues: [["before"]],
    beforeFormulas: [["before"]],
  });

  assert.equal(appended, null);
  assert.equal((await log.list({ limit: 10 })).length, 0);
  assert.equal((await log.listForCurrentWorkbook(10)).length, 0);
});
void test("delete is scoped to the active workbook", async () => {
  const settingsStore = createInMemorySettingsStore();

  let currentWorkbookId: string | null = "url_sha256:workbook-a";
  const getWorkbookContext = (): Promise<WorkbookContext> => Promise.resolve({
    workbookId: currentWorkbookId,
    workbookName: "Workbook",
    source: currentWorkbookId ? "document.url" : "unknown",
  });

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-${idCounter}`;
  };

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
    createId,
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });

  const snapshotA = await log.append({
    toolName: "write_cells",
    toolCallId: "call-a",
    address: "Sheet1!A1",
    beforeValues: [["a"]],
    beforeFormulas: [["a"]],
  });

  currentWorkbookId = "url_sha256:workbook-b";

  const snapshotB = await log.append({
    toolName: "write_cells",
    toolCallId: "call-b",
    address: "Sheet1!A2",
    beforeValues: [["b"]],
    beforeFormulas: [["b"]],
  });

  assert.ok(snapshotA);
  assert.ok(snapshotB);

  currentWorkbookId = "url_sha256:workbook-a";

  const deletedOtherWorkbook = await log.delete(snapshotB?.id ?? "");
  assert.equal(deletedOtherWorkbook, false);

  const deletedCurrentWorkbook = await log.delete(snapshotA?.id ?? "");
  assert.equal(deletedCurrentWorkbook, true);

  currentWorkbookId = null;
  const deletedWithoutIdentity = await log.delete(snapshotB?.id ?? "");
  assert.equal(deletedWithoutIdentity, false);

  currentWorkbookId = "url_sha256:workbook-b";
  const remainingCurrent = await log.listForCurrentWorkbook(10);
  assert.equal(remainingCurrent.length, 1);
  assert.equal(remainingCurrent[0]?.id, snapshotB?.id);
});
void test("clearForCurrentWorkbook removes only matching workbook checkpoints", async () => {
  const settingsStore = createInMemorySettingsStore();

  let currentWorkbookId: string | null = "url_sha256:workbook-c";
  const getWorkbookContext = (): Promise<WorkbookContext> => Promise.resolve({
    workbookId: currentWorkbookId,
    workbookName: "Workbook",
    source: currentWorkbookId ? "document.url" : "unknown",
  });

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-${idCounter}`;
  };

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
    createId,
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });

  await log.append({
    toolName: "write_cells",
    toolCallId: "call-3",
    address: "Sheet1!A1",
    beforeValues: [["a"]],
    beforeFormulas: [["a"]],
  });

  currentWorkbookId = "url_sha256:workbook-d";

  await log.append({
    toolName: "write_cells",
    toolCallId: "call-4",
    address: "Sheet1!A2",
    beforeValues: [["b"]],
    beforeFormulas: [["b"]],
  });

  currentWorkbookId = "url_sha256:workbook-c";
  const removed = await log.clearForCurrentWorkbook();
  assert.equal(removed, 1);

  const remainingCurrent = await log.listForCurrentWorkbook(10);
  assert.equal(remainingCurrent.length, 0);

  const remainingAll = await log.list({ limit: 10 });
  assert.equal(remainingAll.length, 1);
  assert.equal(remainingAll[0]?.toolCallId, "call-4");
});
void test("clearForCurrentWorkbook is a no-op when workbook identity is unavailable", async () => {
  const settingsStore = createInMemorySettingsStore();

  await settingsStore.set(RECOVERY_SETTING_KEY, {
    version: 1,
    snapshots: [
      {
        id: "snap-a",
        at: 1700000000000,
        toolName: "write_cells",
        toolCallId: "call-a",
        address: "Sheet1!A1",
        changedCount: 1,
        cellCount: 1,
        beforeValues: [["a"]],
        beforeFormulas: [["a"]],
        workbookId: "url_sha256:a",
      },
    ],
  });

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId: null,
      workbookName: null,
      source: "unknown",
    }),
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });

  const removed = await log.clearForCurrentWorkbook();
  assert.equal(removed, 0);
  assert.equal((await log.list({ limit: 10 })).length, 1);
});
void test("append skips oversized checkpoints", async () => {
  const settingsStore = createInMemorySettingsStore();

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve({
      workbookId: "url_sha256:big-workbook",
      workbookName: "Big.xlsx",
      source: "document.url",
    }),
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });

  const rows = 201;
  const cols = 101;
  const bigValues = Array.from({ length: rows }, () => Array.from({ length: cols }, () => "v"));
  const bigFormulas = Array.from({ length: rows }, () => Array.from({ length: cols }, () => ""));

  assert.ok(rows * cols > MAX_RECOVERY_CELLS);

  const snapshot = await log.append({
    toolName: "write_cells",
    toolCallId: "call-big",
    address: "Sheet1!A1:CZ201",
    beforeValues: bigValues,
    beforeFormulas: bigFormulas,
  });

  assert.equal(snapshot, null);

  const entries = await log.listForCurrentWorkbook(10);
  assert.equal(entries.length, 0);
});

void test("appendModifyStructure stores cell counts from preserved data ranges", async () => {
  const settingsStore = createInMemorySettingsStore();

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve({
      workbookId: "url_sha256:structure-data-count",
      workbookName: "Structure.xlsx",
      source: "document.url",
    }),
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });

  const snapshot = await log.appendModifyStructure({
    toolName: "modify_structure",
    toolCallId: "call-structure-count",
    address: "Data!3:4",
    modifyStructureState: {
      kind: "rows_present",
      sheetId: "sheet-data",
      sheetName: "Data",
      position: 3,
      count: 2,
      dataRange: {
        address: "B3:C4",
        rowCount: 2,
        columnCount: 2,
        values: [[1, 2], [3, 4]],
        formulas: [["", ""], ["", ""]],
      },
    },
  });

  assert.ok(snapshot);
  assert.equal(snapshot?.cellCount, 4);
});

void test("recovery log surfaces persistence failure and does not retain an unsaved checkpoint", async () => {
  const values = new Map<string, DynamicValue>();
  let rejectWrites = true;
  const settingsStore = {
    get: <T>(key: string): Promise<T | null> => Promise.resolve((values.get(key) as T | undefined) ?? null),
    set: (key: string, value: DynamicValue): Promise<void> => {
      if (rejectWrites) return Promise.reject(new Error("quota exceeded"));
      values.set(key, value);
      return Promise.resolve();
    },
  };
  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve({
      workbookId: "url_sha256:persistence-failure",
      workbookName: "Failure.xlsx",
      source: "document.url",
    }),
  });

  await assert.rejects(log.append({
    toolName: "write_cells",
    toolCallId: "call-failed-save",
    address: "Sheet1!A1",
    beforeValues: [["before"]],
    beforeFormulas: [[""]],
  }), /quota exceeded/u);

  rejectWrites = false;
  assert.equal((await log.list()).length, 0);
});

void test("recovery log retries loading after a storage read failure", async () => {
  const settingsStore = createInMemorySettingsStore();
  await settingsStore.set(RECOVERY_SETTING_KEY, { version: 1, snapshots: [] });
  let rejectRead = true;
  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve({
      get: <T>(key: string): Promise<T | null> => rejectRead
        ? Promise.reject(new Error("read failed"))
        : settingsStore.get<T>(key),
      set: (key: string, value: DynamicValue): Promise<void> => settingsStore.set(key, value),
    }),
  });

  await assert.rejects(log.list(), /read failed/u);
  rejectRead = false;
  assert.deepEqual(await log.list(), []);
});

void test("mutation finalization warns instead of claiming failed audit and recovery persistence", async () => {
  const result: {
    content: [{ type: "text"; text: string }];
    details: { recovery?: RecoveryCheckpointDetails };
  } = {
    content: [{ type: "text", text: "Changed workbook." }],
    details: {},
  };
  const appendResultNote = (target: typeof result, note: string): void => {
    target.content[0].text += `\n\n${note}`;
  };

  const finalized = await finalizeMutationOperation({
    appendAuditEntry: () => Promise.reject(new Error("audit write failed")),
  }, {
    auditEntry: {
      toolName: "write_cells",
      toolCallId: "call-finalize-failure",
      blocked: false,
      changedCount: 1,
      changes: [],
    },
    recovery: {
      result,
      appendRecoverySnapshot: () => Promise.reject(new Error("recovery write failed")),
      appendResultNote,
      unavailableReason: "not expected",
      unavailableNote: "not expected",
    },
  });

  assert.deepEqual(finalized, { checkpointCreated: false });
  assert.match(result.content[0].text, /audit entry could not be saved/u);
  assert.match(result.content[0].text, /Recovery checkpoint could not be saved/u);
  assert.equal(result.details.recovery?.status, "not_available");
  assert.equal(result.details.recovery?.reason, "checkpoint_creation_failed");
});

void test("appendModifyStructure skips oversized preserved data ranges", async () => {
  const settingsStore = createInMemorySettingsStore();

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve({
      workbookId: "url_sha256:structure-data-big",
      workbookName: "StructureBig.xlsx",
      source: "document.url",
    }),
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });

  const snapshot = await log.appendModifyStructure({
    toolName: "modify_structure",
    toolCallId: "call-structure-big",
    address: "Data!A1:CU200",
    modifyStructureState: {
      kind: "rows_present",
      sheetId: "sheet-data",
      sheetName: "Data",
      position: 1,
      count: 201,
      dataRange: {
        address: "A1:CV201",
        rowCount: 201,
        columnCount: 100,
        values: Array.from({ length: 201 }, () => Array.from({ length: 100 }, () => "x")),
        formulas: Array.from({ length: 201 }, () => Array.from({ length: 100 }, () => "")),
      },
    },
  });

  assert.equal(snapshot, null);
});
