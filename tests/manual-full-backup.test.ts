import assert from "node:assert/strict";
import { test } from "node:test";

import type { WorkspaceFileEntry } from "../src/files/types.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import {
  captureWorkbookCompressedBytes,
  ManualFullWorkbookBackupStore,
} from "../src/workbook/manual-full-backup.ts";

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

void test("manual backup creates a downloadable copy of the current Office workbook", async () => {
  const previousOffice = Object.getOwnPropertyDescriptor(globalThis, "Office");
  const files: WorkspaceFileEntry[] = [];
  const writtenFiles = new Map<string, string>();

  const document = {
    getFileAsync(
      fileType: string,
      options: { sliceSize?: number },
      callback?: (result: DynamicObject) => void,
    ): void {
      assert.equal(fileType, "compressed");
      assert.equal(options.sliceSize, 1_048_576);
      callback?.({
        status: "succeeded",
        value: {
          size: 2,
          sliceCount: 1,
          getSliceAsync(_index: number, sliceCallback?: (result: DynamicObject) => void): void {
            sliceCallback?.({ status: "succeeded", value: { data: [7, 9] } });
          },
          closeAsync(closeCallback?: (result: DynamicObject) => void): void {
            closeCallback?.({ status: "succeeded" });
          },
        },
      });
    },
  };

  Object.defineProperty(globalThis, "Office", {
    configurable: true,
    value: { context: { document } },
  });

  const store = new ManualFullWorkbookBackupStore({
    getWorkbookContext: () => Promise.resolve({
      workbookId: "wb-1",
      workbookName: "Workbook.xlsx",
      source: "document.url",
    }),
    getWorkspace: () => ({
      listFiles: () => Promise.resolve(files),
      writeBase64File: (path, base64, mimeType) => {
        writtenFiles.set(path, `${mimeType}:${base64}`);
        files.push(makeManualBackupFile({
          workbookId: "wb-1",
          backupId: path.slice(path.lastIndexOf("/") + 1, -5),
          modifiedAt: 1_700_000_000_000,
        }));
        return Promise.resolve();
      },
      downloadFile: () => Promise.resolve(),
      deleteFile: () => Promise.resolve(),
    }),
    captureWorkbookBytes: () => captureWorkbookCompressedBytes(),
    now: () => 1_700_000_000_000,
    createSuffix: () => "office",
  });

  try {
    const created = await store.create();
    const listed = await store.listForCurrentWorkbook();

    assert.deepEqual(created, {
      id: "2023-11-14T22-13-20-000Z_workbook.xlsx_office",
      path: "manual-backups/full-workbook/v1/wb-1/2023-11-14T22-13-20-000Z_workbook.xlsx_office.xlsx",
      createdAt: 1_700_000_000_000,
      sizeBytes: 2,
    });
    assert.equal(
      writtenFiles.get("manual-backups/full-workbook/v1/wb-1/2023-11-14T22-13-20-000Z_workbook.xlsx_office.xlsx"),
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet:Bwk=",
    );
    assert.equal(listed[0]?.id, "2023-11-14T22-13-20-000Z_workbook.xlsx_office");
  } finally {
    if (previousOffice) Object.defineProperty(globalThis, "Office", previousOffice);
    else delete (globalThis as { Office?: DynamicValue }).Office;
  }
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
  assert.equal(restored.id, targetBackupId);
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
