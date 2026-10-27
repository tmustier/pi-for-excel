import assert from "node:assert/strict";
import test from "node:test";

import { createWorkbookIdentityFollower } from "../src/taskpane/workbook-identity-follower.ts";
import { getSessionWorkbookId, moveSessionToWorkbook } from "../src/workbook/session-association.ts";

void test("Save As is reported once as a move from the old identity to the new one", async () => {
  const identities = ["url_sha256:budget", "url_sha256:budget", null, "url_sha256:budget-v2", "url_sha256:budget-v2"];
  const changes: Array<{ from: string | null; to: string }> = [];
  const follower = createWorkbookIdentityFollower({
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

void test("the first save of a new workbook moves its tabs from the document token to the path", async () => {
  const identities = ["doc_instance:abc", "url_sha256:book1"];
  const changes: Array<{ from: string | null; to: string }> = [];
  const follower = createWorkbookIdentityFollower({
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
