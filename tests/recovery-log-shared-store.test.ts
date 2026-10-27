import assert from "node:assert/strict";
import { test } from "node:test";

import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import { createInMemorySettingsStore, type InMemorySettingsStore } from "./fixtures/recovery-log.ts";

function workbook(id: string): () => Promise<WorkbookContext> {
  return () => Promise.resolve({ workbookId: id, workbookName: `${id}.xlsx`, source: "document.url" });
}

function createLog(store: InMemorySettingsStore, workbookId: string, clock: { now: number }, prefix: string): WorkbookRecoveryLog {
  let counter = 0;
  return new WorkbookRecoveryLog({
    getStore: () => Promise.resolve(store.recovery),
    getWorkbookContext: workbook(workbookId),
    now: () => {
      clock.now += 1;
      return clock.now;
    },
    createId: () => {
      counter += 1;
      return `${prefix}-${counter}`;
    },
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });
}

function appendWrite(log: WorkbookRecoveryLog, address: string): Promise<unknown> {
  return log.append({
    toolName: "write_cells",
    toolCallId: `call-${address}`,
    address,
    changedCount: 1,
    beforeValues: [["before"]],
    beforeFormulas: [["before"]],
  });
}

void test("two taskpanes sharing one store keep each other's backups", async () => {
  const store = createInMemorySettingsStore();
  const clock = { now: 1_700_000_000_000 };
  // Both panes open (and list) before either writes, as two taskpanes do.
  const paneA = createLog(store, "url_sha256:book", clock, "a");
  const paneB = createLog(store, "url_sha256:book", clock, "b");
  assert.equal((await paneA.listForCurrentWorkbook()).length, 0);
  assert.equal((await paneB.listForCurrentWorkbook()).length, 0);

  await appendWrite(paneA, "Sheet1!K1");
  await appendWrite(paneB, "Sheet1!K2");
  await Promise.all([appendWrite(paneA, "Sheet1!K3"), appendWrite(paneB, "Sheet1!K4")]);

  const addresses = (await paneA.listForCurrentWorkbook(50)).map((entry) => entry.address).sort();
  assert.deepEqual(addresses, ["Sheet1!K1", "Sheet1!K2", "Sheet1!K3", "Sheet1!K4"]);
  assert.equal((await paneB.listForCurrentWorkbook(50)).length, 4);
});

void test("the obsolete count setting no longer expires backups", async () => {
  const store = createInMemorySettingsStore();
  await store.set("workbook.recovery.retentionLimit.v1", 5);
  const clock = { now: 1_700_000_000_000 };
  const bookA = createLog(store, "url_sha256:a", clock, "a");
  const bookB = createLog(store, "url_sha256:b", clock, "b");

  for (let index = 1; index <= 8; index += 1) await appendWrite(bookA, `Sheet1!A${index}`);
  for (let index = 1; index <= 3; index += 1) await appendWrite(bookB, `Sheet1!B${index}`);

  const keptA = (await bookA.listForCurrentWorkbook(50)).map((entry) => entry.address);
  assert.deepEqual(keptA, ["Sheet1!A8", "Sheet1!A7", "Sheet1!A6", "Sheet1!A5", "Sheet1!A4", "Sheet1!A3", "Sheet1!A2", "Sheet1!A1"]);
  assert.equal((await bookB.listForCurrentWorkbook(50)).length, 3);
});

void test("a backup that cannot be saved is reported, not silently dropped", async () => {
  const store = createInMemorySettingsStore();
  store.recovery.append = () => Promise.reject(new DOMException("Quota exceeded", "QuotaExceededError"));
  const log = createLog(store, "url_sha256:book", { now: 1 }, "x");

  await assert.rejects(appendWrite(log, "Sheet1!A1"), /Backup could not be saved: Quota exceeded/);
});

void test("restore preview finds cells edited since Pi's change", async () => {
  const store = createInMemorySettingsStore();
  let current: unknown[][] = [[1, "=B2*2"], [3, 4]];
  const log = new WorkbookRecoveryLog({
    getStore: () => Promise.resolve(store.recovery),
    getWorkbookContext: workbook("url_sha256:book"),
    now: () => 1_700_000_000_000,
    createId: () => "snap",
    applySnapshot: () => Promise.resolve({ values: [[]], formulas: [[]] }),
    readRangeFormulas: () => Promise.resolve(current),
  });
  await log.append({
    toolName: "write_cells",
    toolCallId: "call",
    address: "Sheet1!B2:C3",
    beforeValues: [["", ""], ["", ""]],
    beforeFormulas: [["", ""], ["", ""]],
    afterFormulas: [[1, "=B2*2"], [3, 4]],
  });

  assert.equal((await log.previewRestore("snap")).laterEdits, null);

  // The user later changed C2 and cleared B3.
  current = [[1, "=B2*3"], ["", 4]];
  const preview = await log.previewRestore("snap");
  assert.equal(preview.at, 1_700_000_000_000);
  assert.deepEqual(preview.laterEdits, {
    count: 2,
    sample: [
      { address: "Sheet1!C2", content: "=B2*3" },
      { address: "Sheet1!B3", content: "(empty)" },
    ],
  });
});
