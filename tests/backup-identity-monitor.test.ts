import assert from "node:assert/strict";
import { test } from "node:test";

import type { DocumentInstanceIdentity } from "../src/host/document-instance.ts";
import { BackupIdentityMonitor } from "../src/workbook/backup-identity-monitor.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import { createRecoveryScopeResolver } from "../src/workbook/recovery-scope.ts";
import { createInMemorySettingsStore } from "./fixtures/recovery-log.ts";

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
