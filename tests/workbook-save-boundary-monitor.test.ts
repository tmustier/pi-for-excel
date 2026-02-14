import assert from "node:assert/strict";
import { test } from "node:test";

import type { WorkspaceFileEntry } from "../src/files/types.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import { ManualFullWorkbookBackupStore } from "../src/workbook/manual-full-backup.ts";
import { WorkbookSaveBoundaryMonitor } from "../src/workbook/save-boundary-monitor.ts";

function makeManualBackupFile(args: {
  workbookId: string;
  backupId: string;
  modifiedAt: number;
}): WorkspaceFileEntry {
  return {
    path: `manual-backups/full-workbook/v1/${args.workbookId}/${args.backupId}.xlsx`,
    name: `${args.backupId}.xlsx`,
    size: 1024,
    modifiedAt: args.modifiedAt,
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    kind: "binary",
    sourceKind: "workspace",
    readOnly: false,
    workbookTag: {
      workbookId: args.workbookId,
      workbookLabel: "Workbook",
      taggedAt: args.modifiedAt,
    },
  };
}

function createManualBackupStoreForTest(files: WorkspaceFileEntry[]): {
  store: ManualFullWorkbookBackupStore;
  downloads: string[];
  deletes: string[];
} {
  const downloads: string[] = [];
  const deletes: string[] = [];

  const store = new ManualFullWorkbookBackupStore({
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId: "wb-1",
      workbookName: "Workbook.xlsx",
      source: "document.url",
    }),
    getWorkspace: () => ({
      listFiles: () => Promise.resolve(files),
      writeBase64File: () => Promise.resolve(),
      downloadFile: (path: string) => {
        downloads.push(path);
        return Promise.resolve();
      },
      deleteFile: (path: string) => {
        deletes.push(path);
        return Promise.resolve();
      },
    }),
    captureWorkbookBytes: () => Promise.resolve(new Uint8Array([1])),
    now: () => 0,
    createSuffix: () => "suffix",
  });

  return { store, downloads, deletes };
}

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

void test("manual backup restore by id searches beyond first 500 entries", async () => {
  const files: WorkspaceFileEntry[] = [];

  for (let index = 0; index < 620; index += 1) {
    const backupId = `backup-${String(index).padStart(4, "0")}`;
    files.push(makeManualBackupFile({
      workbookId: "wb-1",
      backupId,
      modifiedAt: index,
    }));
  }

  const targetBackupId = "backup-0010";
  const targetFile = files.find((file) => file.name === `${targetBackupId}.xlsx`);
  assert.ok(targetFile);

  const { store, downloads } = createManualBackupStoreForTest(files);
  const restored = await store.downloadByIdForCurrentWorkbook(targetBackupId);

  assert.ok(restored);
  assert.equal(restored?.id, targetBackupId);
  assert.deepEqual(downloads, [targetFile.path]);
});

void test("manual backup clear removes all workbook backups beyond first 500 entries", async () => {
  const files: WorkspaceFileEntry[] = [];

  for (let index = 0; index < 620; index += 1) {
    const backupId = `backup-${String(index).padStart(4, "0")}`;
    files.push(makeManualBackupFile({
      workbookId: "wb-1",
      backupId,
      modifiedAt: index,
    }));
  }

  for (let index = 0; index < 5; index += 1) {
    const backupId = `other-${String(index).padStart(2, "0")}`;
    files.push(makeManualBackupFile({
      workbookId: "wb-2",
      backupId,
      modifiedAt: index,
    }));
  }

  const { store, deletes } = createManualBackupStoreForTest(files);
  const removed = await store.clearForCurrentWorkbook();

  assert.equal(removed, 620);
  assert.equal(deletes.length, 620);
  assert.equal(deletes.some((path) => path.includes("/wb-2/")), false);
});
