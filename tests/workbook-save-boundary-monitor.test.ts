import assert from "node:assert/strict";
import { test } from "node:test";

import type { WorkbookContext } from "../src/workbook/context.ts";
import { WorkbookSaveBoundaryMonitor } from "../src/workbook/save-boundary-monitor.ts";

void test("clears backups on first poll when workbook is already saved", async () => {
  const isDirty = false;
  let clears = 0;

  const monitor = new WorkbookSaveBoundaryMonitor({
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId: "url_sha256:book-1",
      workbookName: "Book1.xlsx",
      source: "document.url",
    }),
    readWorkbookDirtyState: () => Promise.resolve(isDirty),
    clearBackupsForCurrentWorkbook: () => {
      clears += 1;
      return Promise.resolve(1);
    },
  });

  await monitor.checkOnce();
  assert.equal(clears, 1);

  await monitor.checkOnce();
  assert.equal(clears, 1);
});

void test("clears backups when workbook dirty state transitions from dirty to saved", async () => {
  let isDirty = true;
  let clears = 0;

  const monitor = new WorkbookSaveBoundaryMonitor({
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId: "url_sha256:book-1",
      workbookName: "Book1.xlsx",
      source: "document.url",
    }),
    readWorkbookDirtyState: () => Promise.resolve(isDirty),
    clearBackupsForCurrentWorkbook: () => {
      clears += 1;
      return Promise.resolve(1);
    },
  });

  await monitor.checkOnce();
  assert.equal(clears, 0);

  isDirty = false;
  await monitor.checkOnce();
  assert.equal(clears, 1);

  await monitor.checkOnce();
  assert.equal(clears, 1);
});

void test("does not clear backups when workbook identity is unavailable", async () => {
  let clears = 0;

  const monitor = new WorkbookSaveBoundaryMonitor({
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId: null,
      workbookName: null,
      source: "unknown",
    }),
    readWorkbookDirtyState: () => Promise.resolve(false),
    clearBackupsForCurrentWorkbook: () => {
      clears += 1;
      return Promise.resolve(0);
    },
  });

  await monitor.checkOnce();
  assert.equal(clears, 0);
});

void test("tracks dirty transitions independently per workbook", async () => {
  let workbookId = "url_sha256:book-1";
  let isDirty = true;
  const clearedFor: string[] = [];

  const monitor = new WorkbookSaveBoundaryMonitor({
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId,
      workbookName: "Workbook",
      source: "document.url",
    }),
    readWorkbookDirtyState: () => Promise.resolve(isDirty),
    clearBackupsForCurrentWorkbook: () => {
      clearedFor.push(workbookId);
      return Promise.resolve(1);
    },
  });

  await monitor.checkOnce(); // book-1 dirty

  workbookId = "url_sha256:book-2";
  isDirty = true;
  await monitor.checkOnce(); // book-2 dirty

  workbookId = "url_sha256:book-1";
  isDirty = false;
  await monitor.checkOnce(); // book-1 saved => clear

  workbookId = "url_sha256:book-2";
  isDirty = false;
  await monitor.checkOnce(); // book-2 saved => clear

  assert.deepEqual(clearedFor, ["url_sha256:book-1", "url_sha256:book-2"]);
});

