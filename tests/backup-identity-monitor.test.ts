import assert from "node:assert/strict";
import { test } from "node:test";

import type { DocumentInstanceIdentity } from "../src/host/document-instance.ts";
import { BackupIdentityMonitor, startBackupIdentityPolling } from "../src/workbook/backup-identity-monitor.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import { createRecoveryScopeResolver } from "../src/workbook/recovery-scope.ts";
import { createInMemorySettingsStore } from "./fixtures/recovery-log.ts";

import { deferred } from "./fixtures/workbook-deferred.ts";

const write = {
  toolName: "write_cells" as const,
  toolCallId: "call-1",
  address: "Sheet1!A1",
  beforeValues: [["before"]],
  beforeFormulas: [[""]],
};

function setup(initial: WorkbookContext, document: DocumentInstanceIdentity | null = null) {
  const settings = createInMemorySettingsStore();
  let context = initial;
  let nextId = 0;
  const getWorkbookContext = () => Promise.resolve(context);
  const getDocumentInstance = () => document;
  const createLog = () => new WorkbookRecoveryLog({
    getStore: () => Promise.resolve(settings.recovery),
    getWorkbookContext,
    getDocumentInstance,
    createId: () => `backup-${nextId += 1}`,
    now: () => 1_700_000_000_000 + nextId,
    applySnapshot: () => Promise.resolve({ values: [["after"]], formulas: [[""]] }),
  });
  const log = createLog();
  const resolver = createRecoveryScopeResolver({ getWorkbookContext, getDocumentInstance });
  const monitor = new BackupIdentityMonitor({
    resolveScope: () => resolver.resolveForRead(),
    rescopeBackups: (args) => log.rescopeWorkbook(args),
  });
  return {
    log,
    monitor,
    createLog,
    setContext: (next: WorkbookContext) => { context = next; },
  };
}

const SAVED: WorkbookContext = { workbookId: "url_sha256:book-1", workbookName: "Book1.xlsx", source: "document.url" };

void test("backups survive saves and reloads on a saved workbook (autosave hosts)", async () => {
  const { log, monitor, createLog } = setup(SAVED);

  await monitor.checkOnce();
  assert.ok(await log.append(write));
  await monitor.checkOnce();
  await monitor.checkOnce();

  assert.deepEqual((await createLog().listForCurrentWorkbook()).map((s) => s.toolCallId), ["call-1"]);
});

void test("the first save of a never-saved workbook moves its backups to the saved identity", async () => {
  let token: string | null = null;
  const document: DocumentInstanceIdentity = {
    read: () => token,
    ensure: () => Promise.resolve(token ??= "doc-token"),
  };
  const { log, monitor, setContext } = setup({ workbookId: null, workbookName: null, source: "unknown" }, document);

  const backup = await log.append(write);
  assert.equal(backup?.workbookId, "doc_instance:doc-token");
  await monitor.checkOnce();

  setContext({ workbookId: "url_sha256:first-save", workbookName: "First.xlsx", source: "document.url" });
  await monitor.checkOnce();

  const listed = await log.listForCurrentWorkbook();
  assert.deepEqual(listed.map((s) => [s.id, s.workbookId, s.workbookLabel]), [
    [backup?.id, "url_sha256:first-save", "First.xlsx"],
  ]);
  const restored = await log.restore(listed[0]?.id ?? "");
  assert.equal(restored.address, "Sheet1!A1", "still restorable after the move");
});

void test("switching between two saved workbooks moves nothing", async () => {
  const { log, monitor, setContext } = setup(SAVED);
  assert.ok(await log.append(write));
  await monitor.checkOnce();

  setContext({ workbookId: "url_sha256:book-2", workbookName: "Book2.xlsx", source: "document.url" });
  await monitor.checkOnce();
  assert.deepEqual(await log.listForCurrentWorkbook(), []);

  setContext(SAVED);
  assert.equal((await log.listForCurrentWorkbook()).length, 1);
});

void test("overlapping identity checks join one observation and one backup move", async () => {
  const observation = deferred<{ workbookId: string; workbookLabel: string; unsaved: boolean }>();
  let reads = 0;
  let moves = 0;
  const monitor = new BackupIdentityMonitor({
    resolveScope: () => {
      reads += 1;
      return reads === 1
        ? Promise.resolve({ workbookId: "doc_instance:token", workbookLabel: "Unsaved", unsaved: true })
        : observation.promise;
    },
    rescopeBackups: () => { moves += 1; return Promise.resolve(1); },
  });
  await monitor.checkOnce();
  const first = monitor.checkOnce();
  const joined = monitor.checkOnce();
  observation.resolve({ workbookId: "saved", workbookLabel: "Saved", unsaved: false });
  await Promise.all([first, joined]);
  assert.equal(reads, 2, "concurrent checks must share the host observation");
  assert.equal(moves, 1, "first-save backups must be moved once");
});

void test("stopping polling invalidates an in-flight identity observation", async () => {
  const observation = deferred<{ workbookId: string; workbookLabel: string; unsaved: boolean }>();
  let reads = 0;
  let moves = 0;
  const monitor = new BackupIdentityMonitor({
    resolveScope: () => {
      reads += 1;
      return reads === 1
        ? Promise.resolve({ workbookId: "doc_instance:token", workbookLabel: "Unsaved", unsaved: true })
        : observation.promise;
    },
    rescopeBackups: () => { moves += 1; return Promise.resolve(1); },
  });
  await monitor.checkOnce();
  const stop = startBackupIdentityPolling({ monitor, intervalMs: 60_000 });
  const pending = monitor.checkOnce();
  stop();
  observation.resolve({ workbookId: "saved", workbookLabel: "Saved", unsaved: false });
  await pending;
  assert.equal(moves, 0, "stopped polling must not move backups from its pending observation");
  await monitor.checkOnce();
  assert.equal(moves, 1, "a later active check must still observe the first save");
});

void test("stopping during an atomic backup move prevents observation publication", async () => {
  const moving = deferred<void>();
  const finishMove = deferred<number>();
  let reads = 0;
  let moves = 0;
  const monitor = new BackupIdentityMonitor({
    resolveScope: () => Promise.resolve(++reads === 1
      ? { workbookId: "doc_instance:token", workbookLabel: "Unsaved", unsaved: true }
      : { workbookId: "saved", workbookLabel: "Saved", unsaved: false }),
    rescopeBackups: () => {
      moves += 1;
      if (moves === 1) { moving.resolve(); return finishMove.promise; }
      return Promise.resolve(0);
    },
  });
  await monitor.checkOnce();
  const controller = new AbortController();
  const pending = monitor.checkOnce(controller.signal);
  await moving.promise;
  controller.abort();
  finishMove.resolve(1);
  await pending;
  await monitor.checkOnce();
  assert.equal(moves, 2, "the stopped check must not publish its observed identity after the move");
});
