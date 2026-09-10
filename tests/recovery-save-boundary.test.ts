import assert from "node:assert/strict";
import { test } from "node:test";

import type { WorkbookContext } from "../src/workbook/context.ts";
import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import { WorkbookSaveBoundaryMonitor } from "../src/workbook/save-boundary-monitor.ts";
import { createInMemorySettingsStore } from "./fixtures/recovery-log.ts";

const contexts: Record<string, WorkbookContext> = {
  "book-1": {
    workbookId: "url_sha256:book-1",
    workbookName: "Book1.xlsx",
    source: "document.url",
  },
  "book-2": {
    workbookId: "url_sha256:book-2",
    workbookName: "Book2.xlsx",
    source: "document.url",
  },
};

void test("save removes earlier backups while a backup made after save survives the next poll", async () => {
  const settings = createInMemorySettingsStore();
  let isDirty = true;
  const workbookContext = contexts["book-1"];
  if (!workbookContext) throw new Error("Expected workbook context.");

  let nextId = 0;
  const recoveryLog = new WorkbookRecoveryLog({
    settings,
    getWorkbookContext: () => Promise.resolve(workbookContext),
    createId: () => `backup-${nextId += 1}`,
  });
  const monitor = new WorkbookSaveBoundaryMonitor({
    resolveWorkbookId: () => Promise.resolve(workbookContext.workbookId),
    readWorkbookDirtyState: () => Promise.resolve(isDirty),
    clearBackupsForWorkbook: (workbookId) => recoveryLog.clearForWorkbook(workbookId),
  });

  await recoveryLog.append({
    toolName: "write_cells",
    toolCallId: "before-save",
    address: "Sheet1!A1",
    beforeValues: [["old"]],
    beforeFormulas: [[""]],
  });
  await monitor.checkOnce();

  isDirty = false;
  await monitor.checkOnce();
  assert.deepEqual(await recoveryLog.listForCurrentWorkbook(), []);

  await recoveryLog.append({
    toolName: "write_cells",
    toolCallId: "after-save",
    address: "Sheet1!B1",
    beforeValues: [["new"]],
    beforeFormulas: [[""]],
  });
  await monitor.checkOnce();

  const surviving = await recoveryLog.listForCurrentWorkbook();
  assert.deepEqual(surviving.map((snapshot) => snapshot.toolCallId), ["after-save"]);
});

void test("an already-saved workbook loses existing backups only on the first poll", async () => {
  const settings = createInMemorySettingsStore();
  const workbookContext = contexts["book-1"];
  if (!workbookContext) throw new Error("Expected workbook context.");

  let nextId = 0;
  const recoveryLog = new WorkbookRecoveryLog({
    settings,
    getWorkbookContext: () => Promise.resolve(workbookContext),
    createId: () => `backup-${nextId += 1}`,
  });
  const monitor = new WorkbookSaveBoundaryMonitor({
    resolveWorkbookId: () => Promise.resolve(workbookContext.workbookId),
    readWorkbookDirtyState: () => Promise.resolve(false),
    clearBackupsForWorkbook: (workbookId) => recoveryLog.clearForWorkbook(workbookId),
  });

  await recoveryLog.append({
    toolName: "write_cells",
    toolCallId: "before-first-poll",
    address: "Sheet1!A1",
    beforeValues: [[1]],
    beforeFormulas: [[""]],
  });
  await monitor.checkOnce();
  assert.deepEqual(await recoveryLog.listForCurrentWorkbook(), []);

  await recoveryLog.append({
    toolName: "write_cells",
    toolCallId: "after-first-poll",
    address: "Sheet1!B1",
    beforeValues: [[2]],
    beforeFormulas: [[""]],
  });
  await monitor.checkOnce();
  assert.deepEqual(
    (await recoveryLog.listForCurrentWorkbook()).map((snapshot) => snapshot.toolCallId),
    ["after-first-poll"],
  );
});

void test("a save poll without workbook identity preserves backups", async () => {
  const settings = createInMemorySettingsStore();
  const knownContext = contexts["book-1"];
  if (!knownContext) throw new Error("Expected workbook context.");
  let currentContext: WorkbookContext = knownContext;

  const recoveryLog = new WorkbookRecoveryLog({
    settings,
    getWorkbookContext: () => Promise.resolve(currentContext),
    createId: () => "backup-1",
  });
  await recoveryLog.append({
    toolName: "write_cells",
    toolCallId: "identity-guard",
    address: "Sheet1!A1",
    beforeValues: [[1]],
    beforeFormulas: [[""]],
  });

  currentContext = { workbookId: null, workbookName: null, source: "unknown" };
  const monitor = new WorkbookSaveBoundaryMonitor({
    resolveWorkbookId: () => Promise.resolve(currentContext.workbookId),
    readWorkbookDirtyState: () => Promise.resolve(false),
    clearBackupsForWorkbook: (workbookId) => recoveryLog.clearForWorkbook(workbookId),
  });
  await monitor.checkOnce();

  currentContext = knownContext;
  assert.deepEqual(
    (await recoveryLog.listForCurrentWorkbook()).map((snapshot) => snapshot.toolCallId),
    ["identity-guard"],
  );
});

void test("save transitions clear backups independently for each workbook", async () => {
  const settings = createInMemorySettingsStore();
  let currentWorkbook = "book-1";
  let isDirty = true;
  let nextId = 0;
  const getWorkbookContext = (): Promise<WorkbookContext> => {
    const context = contexts[currentWorkbook];
    if (!context) throw new Error("Expected workbook context.");
    return Promise.resolve(context);
  };
  const recoveryLog = new WorkbookRecoveryLog({
    settings,
    getWorkbookContext,
    createId: () => `backup-${nextId += 1}`,
  });
  const monitor = new WorkbookSaveBoundaryMonitor({
    resolveWorkbookId: async () => (await getWorkbookContext()).workbookId,
    readWorkbookDirtyState: () => Promise.resolve(isDirty),
    clearBackupsForWorkbook: (workbookId) => recoveryLog.clearForWorkbook(workbookId),
  });

  for (const workbook of ["book-1", "book-2"]) {
    currentWorkbook = workbook;
    await recoveryLog.append({
      toolName: "write_cells",
      toolCallId: `${workbook}-backup`,
      address: "Sheet1!A1",
      beforeValues: [[workbook]],
      beforeFormulas: [[""]],
    });
    await monitor.checkOnce();
  }

  currentWorkbook = "book-1";
  isDirty = false;
  await monitor.checkOnce();
  assert.deepEqual(await recoveryLog.listForCurrentWorkbook(), []);

  currentWorkbook = "book-2";
  assert.deepEqual(
    (await recoveryLog.listForCurrentWorkbook()).map((snapshot) => snapshot.toolCallId),
    ["book-2-backup"],
  );
  await monitor.checkOnce();
  assert.deepEqual(await recoveryLog.listForCurrentWorkbook(), []);
});
