import assert from "node:assert/strict";
import { test } from "node:test";

import type { DocumentInstanceIdentity } from "../src/host/document-instance.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import { WorkbookRecoveryLog, type WorkbookRecoverySnapshot } from "../src/workbook/recovery-log.ts";
import { createRecoveryScopeResolver } from "../src/workbook/recovery-scope.ts";
import { WorkbookSaveBoundaryMonitor } from "../src/workbook/save-boundary-monitor.ts";
import { createInMemorySettingsStore, RECOVERY_SETTING_KEY } from "./fixtures/recovery-log.ts";

const UNSAVED: WorkbookContext = { workbookId: null, workbookName: null, source: "unknown" };

/**
 * A document that can hold one token. `persists: false` models a host whose
 * settings save fails: `ensure()` then resolves null and stores nothing.
 */
function createDocument(options: { token?: string | null; persists?: boolean } = {}) {
  let token = options.token ?? null;
  const persists = options.persists ?? true;
  let ensureCalls = 0;
  const identity: DocumentInstanceIdentity = {
    read: () => token,
    ensure: () => {
      ensureCalls += 1;
      if (token) return Promise.resolve(token);
      if (!persists) return Promise.resolve(null);
      token = `token-${ensureCalls}`;
      return Promise.resolve(token);
    },
  };
  return {
    identity,
    get token() { return token; },
    get ensureCalls() { return ensureCalls; },
  };
}

function createLog(args: {
  settings: ReturnType<typeof createInMemorySettingsStore>;
  context?: () => WorkbookContext;
  document: DocumentInstanceIdentity | null;
  applied?: string[];
}): WorkbookRecoveryLog {
  let nextId = 0;
  return new WorkbookRecoveryLog({
    settings: args.settings,
    getWorkbookContext: () => Promise.resolve(args.context?.() ?? UNSAVED),
    getDocumentInstance: () => args.document,
    createId: () => `snap-${nextId += 1}`,
    now: () => 1_700_000_000_000 + nextId,
    applySnapshot: (address) => {
      args.applied?.push(address);
      return Promise.resolve({ values: [["current"]], formulas: [["current"]] });
    },
  });
}

async function persistedSnapshots(settings: ReturnType<typeof createInMemorySettingsStore>): Promise<WorkbookRecoverySnapshot[]> {
  const payload = await settings.get<{ snapshots: WorkbookRecoverySnapshot[] }>(RECOVERY_SETTING_KEY);
  return payload?.snapshots ?? [];
}

const write = { toolName: "write_cells" as const, toolCallId: "call-1", address: "Sheet1!A1", beforeValues: [["before"]], beforeFormulas: [["before"]] };

void test("a never-saved workbook gets checkpoints scoped to its document token, and can restore them", async () => {
  const settings = createInMemorySettingsStore();
  const document = createDocument();
  const applied: string[] = [];
  const log = createLog({ settings, document: document.identity, applied });

  const snapshot = await log.append(write);
  assert.ok(snapshot, "checkpoint created for the unsaved workbook");
  assert.equal(document.ensureCalls, 1);
  assert.equal(document.token, "token-1");
  assert.equal(snapshot.workbookId, "doc_instance:token-1");

  const listed = await log.listForCurrentWorkbook();
  assert.deepEqual(listed.map((item) => item.id), [snapshot.id]);

  const restored = await log.restore(snapshot.id);
  assert.deepEqual(applied, ["Sheet1!A1"]);
  assert.ok(restored.inverseSnapshotId);
  const inverse = (await log.listForCurrentWorkbook()).find((item) => item.id === restored.inverseSnapshotId);
  assert.equal(inverse?.workbookId, snapshot.workbookId, "inverse checkpoint lands in the same scope");
  assert.equal(document.ensureCalls, 1, "restore reuses the token rather than minting one");
});

void test("reading history on a pristine unsaved workbook creates no token", async () => {
  const settings = createInMemorySettingsStore();
  const document = createDocument();
  const log = createLog({ settings, document: document.identity });

  assert.deepEqual(await log.listForCurrentWorkbook(), []);
  assert.equal(await log.clearForCurrentWorkbook(), 0);
  await assert.rejects(() => log.restore("missing"), /not found/i);
  assert.equal(document.ensureCalls, 0);
  assert.equal(document.token, null);
});

void test("two unsaved workbooks cannot see, delete or restore each other's checkpoints", async () => {
  const settings = createInMemorySettingsStore();
  const bookA = createDocument({ token: "aaaaaaaa-0000-4000-8000-000000000001" });
  const bookB = createDocument({ token: "bbbbbbbb-0000-4000-8000-000000000002" });
  const logA = createLog({ settings, document: bookA.identity });
  const snapshot = await logA.append(write);
  assert.ok(snapshot);

  const logB = createLog({ settings, document: bookB.identity });
  assert.deepEqual(await logB.listForCurrentWorkbook(), []);
  assert.equal(await logB.delete(snapshot.id), false);
  assert.equal(await logB.clearForCurrentWorkbook(), 0);
  await assert.rejects(() => logB.restore(snapshot.id), /different workbook/i);

  assert.equal((await persistedSnapshots(settings)).length, 1, "book A's checkpoint is untouched");
});

void test("a saved copy that carries the same document token does not inherit unsaved checkpoints", async () => {
  const settings = createInMemorySettingsStore();
  const document = createDocument({ token: "cccccccc-0000-4000-8000-000000000003" });
  const unsavedLog = createLog({ settings, document: document.identity });
  const snapshot = await unsavedLog.append(write);
  assert.ok(snapshot);

  // Save As copies the document settings, so the copy has the same token but a canonical identity.
  const savedCopy = createLog({
    settings,
    document: document.identity,
    context: () => ({ workbookId: "url_sha256:copy", workbookName: "Copy.xlsx", source: "document.url" }),
  });
  assert.deepEqual(await savedCopy.listForCurrentWorkbook(), []);
  await assert.rejects(() => savedCopy.restore(snapshot.id), /different workbook/i);

  const fresh = await savedCopy.append({ ...write, toolCallId: "call-2" });
  assert.equal(fresh?.workbookId, "url_sha256:copy", "new checkpoints use the canonical identity");
});

void test("when the host cannot persist a token the mutation proceeds without a checkpoint", async () => {
  const settings = createInMemorySettingsStore();
  const document = createDocument({ persists: false });
  const log = createLog({ settings, document: document.identity });

  assert.equal(await log.append(write), null);
  assert.equal(document.ensureCalls, 1);
  assert.deepEqual(await persistedSnapshots(settings), []);
  assert.deepEqual(await log.listForCurrentWorkbook(), []);
});

void test("hosts without document storage keep the previous behaviour for unsaved workbooks", async () => {
  const settings = createInMemorySettingsStore();
  const saved = createLog({
    settings,
    document: null,
    context: () => ({ workbookId: "url_sha256:elsewhere", workbookName: "Elsewhere.xlsx", source: "document.url" }),
  });
  const snapshot = await saved.append(write);
  assert.ok(snapshot);

  const log = createLog({ settings, document: null });
  assert.equal(await log.append(write), null);
  assert.deepEqual(await log.listForCurrentWorkbook(), []);
  await assert.rejects(() => log.restore(snapshot.id), /identity is unavailable/i);
});

void test("saving a never-saved workbook clears its checkpoints at the save boundary", async () => {
  const settings = createInMemorySettingsStore();
  const document = createDocument();
  const log = createLog({ settings, document: document.identity });
  const resolver = createRecoveryScopeResolver({
    getWorkbookContext: () => Promise.resolve(UNSAVED),
    getDocumentInstance: () => document.identity,
  });
  let isDirty = true;
  const monitor = new WorkbookSaveBoundaryMonitor({
    resolveWorkbookId: async () => (await resolver.resolveForRead())?.workbookId ?? null,
    readWorkbookDirtyState: () => Promise.resolve(isDirty),
    clearBackupsForCurrentWorkbook: () => log.clearForCurrentWorkbook(),
  });

  await monitor.checkOnce();
  assert.equal(document.ensureCalls, 0, "polling never mints a token");

  const snapshot = await log.append(write);
  assert.ok(snapshot);
  await monitor.checkOnce();
  assert.equal((await log.listForCurrentWorkbook()).length, 1, "still dirty: checkpoint kept");

  isDirty = false;
  await monitor.checkOnce();
  assert.deepEqual(await log.listForCurrentWorkbook(), [], "saved: checkpoint cleared");
});
