import assert from "node:assert/strict";
import test from "node:test";

import { getWorkbookRules, setWorkbookRules, workbookRulesKey } from "../src/rules/store.ts";
import { createWorkbookIdentityFollower } from "../src/taskpane/workbook-identity-follower.ts";
import { getSessionWorkbookId, moveSessionToWorkbook } from "../src/workbook/session-association.ts";
import { createInMemorySettingsStore } from "./fixtures/recovery-log.ts";
import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";

void test("Save As is reported once as a move from the old identity to the new one", async () => {
  const identities = ["url_sha256:budget", "url_sha256:budget", null, "url_sha256:budget-v2", "url_sha256:budget-v2"];
  const changes: Array<{ from: string | null; to: string }> = [];
  const follower = createWorkbookIdentityFollower({
    settings: createInMemorySettingsStore(),
    resolveIdentity: () => Promise.resolve(identities.shift() ?? null),
    onIdentityChanged: (change) => {
      changes.push(change);
      return Promise.resolve();
    },
  });

  // The pane opens on "budget"; a failed read (null) in between is not a move.
  for (let i = 0; i < 5; i += 1) await follower.checkOnce();
  assert.deepEqual(changes, [{ from: "url_sha256:budget", to: "url_sha256:budget-v2" }]);
});

void test("backups follow observed Save As, rename and move; a failed move retries", async () => {
  const settings = createInMemorySettingsStore();
  let identity = "url_sha256:original";
  const log = new WorkbookRecoveryLog({
    getStore: () => Promise.resolve(settings.recovery),
    getWorkbookContext: () => Promise.resolve({ workbookId: identity, workbookName: "Book.xlsx", source: "document.url" }),
    createId: () => "same-backup", now: () => 1,
    applySnapshot: () => Promise.resolve({ values: [[2]], formulas: [[2]] }),
  });
  await log.append({ toolName: "write_cells", toolCallId: "call", address: "Sheet1!A1", beforeValues: [[1]], beforeFormulas: [[1]] });
  let fail = true;
  const follower = createWorkbookIdentityFollower({
    settings, resolveIdentity: () => Promise.resolve(identity),
    onIdentityChanged: async ({ from, to }) => {
      if (fail) { fail = false; throw new Error("Storage temporarily unavailable"); }
      if (from) await log.rescopeWorkbook({ fromWorkbookId: from, toWorkbookId: to, workbookLabel: "Book.xlsx" });
    },
  });
  await follower.checkOnce();
  for (const next of ["url_sha256:save-as", "url_sha256:renamed", "url_sha256:moved"]) {
    const old = identity;
    identity = next;
    if (next === "url_sha256:save-as") await assert.rejects(follower.checkOnce(), /temporarily unavailable/);
    await follower.checkOnce();
    assert.deepEqual((await log.listForCurrentWorkbook()).map((entry) => entry.id), ["same-backup"]);
    assert.deepEqual(await log.list({ workbookId: old }), []);
    assert.deepEqual((await log.read("same-backup"))?.beforeValues, [[1]]);
  }
});

void test("the first save of a new workbook moves its tabs from the document token to the path", async () => {
  const identities = ["doc_instance:abc", "url_sha256:book1"];
  const changes: Array<{ from: string | null; to: string }> = [];
  const follower = createWorkbookIdentityFollower({
    settings: createInMemorySettingsStore(),
    resolveIdentity: () => Promise.resolve(identities.shift() ?? null),
    onIdentityChanged: (change) => {
      changes.push(change);
      return Promise.resolve();
    },
  });

  await follower.checkOnce();
  // Two polls racing on the same change handle it once.
  await Promise.all([follower.checkOnce(), follower.checkOnce()]);
  assert.deepEqual(changes, [{ from: "doc_instance:abc", to: "url_sha256:book1" }]);
});

void test("workbook rules follow identity changes without removing old rules or overwriting destination rules", async () => {
  const settings = createInMemorySettingsStore();
  let identity = "doc_instance:budget";
  await setWorkbookRules(settings, identity, "Report costs in GBP.");
  await setWorkbookRules(settings, "url_sha256:existing", "Report costs in EUR.");
  await setWorkbookRules(settings, "url_sha256:cleared", null);
  const follower = createWorkbookIdentityFollower({
    settings,
    resolveIdentity: () => Promise.resolve(identity),
    onIdentityChanged: () => Promise.resolve(),
  });
  await follower.checkOnce();

  // First save, Save As/rename and move all change the browser storage identity.
  for (const next of ["url_sha256:saved", "url_sha256:renamed", "url_sha256:moved"]) {
    const previous = identity;
    identity = next;
    await follower.checkOnce();
    assert.equal(await getWorkbookRules(settings, identity), "Report costs in GBP.");
    assert.equal(await getWorkbookRules(settings, previous), "Report costs in GBP.");
  }

  identity = "url_sha256:existing";
  await follower.checkOnce();
  assert.equal(await getWorkbookRules(settings, identity), "Report costs in EUR.");
  identity = "url_sha256:cleared";
  await follower.checkOnce();
  assert.equal(await getWorkbookRules(settings, identity), null);
});

void test("a rules read failure is not absence and the identity move can be retried", async () => {
  for (const failedIdentity of ["url_sha256:old", "url_sha256:new"]) {
    const settings = createInMemorySettingsStore();
    await setWorkbookRules(settings, "url_sha256:old", "Keep all source data.");
    let identity = "url_sha256:old";
    let failRead = true;
    const follower = createWorkbookIdentityFollower({
      settings: {
        ...settings,
        get: <T>(key: string): Promise<T | null> => {
          if (failRead && key === workbookRulesKey(failedIdentity)) {
            failRead = false;
            return Promise.reject(new Error("storage unavailable"));
          }
          return settings.get<T>(key);
        },
      },
      resolveIdentity: () => Promise.resolve(identity),
      onIdentityChanged: () => Promise.resolve(),
    });
    await follower.checkOnce();
    identity = "url_sha256:new";
    await assert.rejects(follower.checkOnce(), /storage unavailable/);
    assert.equal(await getWorkbookRules(settings, identity), null);
    await follower.checkOnce();
    assert.equal(await getWorkbookRules(settings, identity), "Keep all source data.");
    assert.equal(await getWorkbookRules(settings, "url_sha256:old"), "Keep all source data.");
  }
});

void test("moving tabs to a new identity relinks this workbook's sessions only", async () => {
  const values = new Map<string, string>([
    ["session.workbook.v1.mine", "url_sha256:budget"],
    ["session.workbook.v1.resumed", "url_sha256:other-file"],
  ]);
  const settings = {
    get: (key: string) => Promise.resolve(values.get(key) ?? null),
    set: (key: string, value: string) => {
      values.set(key, value);
      return Promise.resolve();
    },
    delete: (key: string) => {
      values.delete(key);
      return Promise.resolve();
    },
  };

  for (const sessionId of ["mine", "resumed", "blank"]) {
    await moveSessionToWorkbook(settings, sessionId, "url_sha256:budget", "url_sha256:budget-v2");
  }
  assert.equal(await getSessionWorkbookId(settings, "mine"), "url_sha256:budget-v2");
  // A session resumed from another workbook keeps its link.
  assert.equal(await getSessionWorkbookId(settings, "resumed"), "url_sha256:other-file");
  assert.equal(await getSessionWorkbookId(settings, "blank"), "url_sha256:budget-v2");
});
