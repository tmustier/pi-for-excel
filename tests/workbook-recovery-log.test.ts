import assert from "node:assert/strict";
import { test } from "node:test";

import {
  MAX_RECOVERY_CELLS,
  WorkbookRecoveryLog,
  type WorkbookRecoverySnapshot,
} from "../src/workbook/recovery-log.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";

const RECOVERY_SETTING_KEY = "workbook.recovery-snapshots.v1";

interface InMemorySettingsStore {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: unknown): Promise<void>;
}

function createInMemorySettingsStore(): InMemorySettingsStore {
  const values = new Map<string, unknown>();

  return {
    get: <T>(key: string): Promise<T | null> => {
      const value = values.get(key);
      return Promise.resolve(value === undefined ? null : value as T);
    },
    set: (key: string, value: unknown): Promise<void> => {
      values.set(key, value);
      return Promise.resolve();
    },
  };
}

function findSnapshotById(snapshots: WorkbookRecoverySnapshot[], id: string): WorkbookRecoverySnapshot | null {
  for (const snapshot of snapshots) {
    if (snapshot.id === id) {
      return snapshot;
    }
  }

  return null;
}

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

void test("restore rejects legacy snapshots without workbook identity", async () => {
  const settingsStore = createInMemorySettingsStore();

  await settingsStore.set(RECOVERY_SETTING_KEY, {
    version: 1,
    snapshots: [
      {
        id: "legacy-1",
        at: 1700000000000,
        toolName: "write_cells",
        toolCallId: "call-legacy",
        address: "Sheet1!A1",
        changedCount: 1,
        cellCount: 1,
        beforeValues: [["before"]],
        beforeFormulas: [["before"]],
      },
    ],
  });

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId: "url_sha256:target",
      workbookName: "Target.xlsx",
      source: "document.url",
    }),
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  await assert.rejects(
    () => log.restore("legacy-1"),
    /missing workbook identity/i,
  );
});

void test("restore rejects when current workbook identity is unavailable", async () => {
  const settingsStore = createInMemorySettingsStore();

  await settingsStore.set(RECOVERY_SETTING_KEY, {
    version: 1,
    snapshots: [
      {
        id: "snap-1",
        at: 1700000000000,
        toolName: "write_cells",
        toolCallId: "call-1",
        address: "Sheet1!A1",
        changedCount: 1,
        cellCount: 1,
        beforeValues: [["before"]],
        beforeFormulas: [["before"]],
        workbookId: "url_sha256:origin",
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
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  await assert.rejects(
    () => log.restore("snap-1"),
    /identity is unavailable/i,
  );
});

void test("restore applies checkpoint values and creates inverse checkpoint", async () => {
  const settingsStore = createInMemorySettingsStore();

  const workbookContext: WorkbookContext = {
    workbookId: "url_sha256:workbook-b",
    workbookName: "Model.xlsx",
    source: "document.url",
  };

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-${idCounter}`;
  };

  let appliedAddress = "";
  let appliedValues: unknown[][] = [];

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve(workbookContext),
    now: () => 1700000001000,
    createId,
    applySnapshot: (address: string, values: unknown[][]) => {
      appliedAddress = address;
      appliedValues = values;
      return Promise.resolve({
        values: [[42]],
        formulas: [[42]],
      });
    },
  });

  const appended = await log.append({
    toolName: "fill_formula",
    toolCallId: "call-2",
    address: "Sheet2!B4",
    changedCount: 1,
    beforeValues: [[10]],
    beforeFormulas: [["=A1+A2"]],
  });

  assert.ok(appended);

  const restored = await log.restore(appended?.id ?? "");

  assert.equal(restored.address, "Sheet2!B4");
  assert.equal(restored.restoredSnapshotId, appended?.id);
  assert.equal(appliedAddress, "Sheet2!B4");
  assert.deepEqual(appliedValues, [["=A1+A2"]]);

  const snapshots = await log.listForCurrentWorkbook(10);
  assert.equal(snapshots.length, 2);

  const inverse = restored.inverseSnapshotId
    ? findSnapshotById(snapshots, restored.inverseSnapshotId)
    : null;

  assert.ok(inverse);
  assert.equal(inverse?.toolName, "restore_snapshot");
  assert.equal(inverse?.restoredFromSnapshotId, appended?.id);
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

void test("restore rejects checkpoints from another workbook", async () => {
  const settingsStore = createInMemorySettingsStore();

  let currentWorkbookId: string | null = "url_sha256:workbook-src";
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

  const snapshot = await log.append({
    toolName: "write_cells",
    toolCallId: "call-5",
    address: "Sheet1!B2",
    beforeValues: [["before"]],
    beforeFormulas: [["before"]],
  });

  assert.ok(snapshot);

  currentWorkbookId = "url_sha256:workbook-other";

  await assert.rejects(
    async () => log.restore(snapshot?.id ?? ""),
    /different workbook/i,
  );
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
