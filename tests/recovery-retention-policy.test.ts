import assert from "node:assert/strict";
import { test } from "node:test";
import { RecoveryStore, type RecoveryCursor } from "../src/workbook/recovery/log-store.ts";
import { WorkbookRecoveryLog, type WorkbookRecoverySnapshot } from "../src/workbook/recovery-log.ts";
import { createInMemorySettingsStore, RECOVERY_SETTING_KEY } from "./fixtures/recovery-log.ts";

function snapshot(id: string, at: number, workbookId = "book-a"): WorkbookRecoverySnapshot {
  return { id, at, workbookId, toolName: "write_cells", toolCallId: `call-${id}`,
    address: "Sheet1!A1:B1", changedCount: 2, cellCount: 2, snapshotKind: "range_values",
    beforeValues: [[3, "Café 中文"]], beforeFormulas: [["=1+2", "Café 中文"]],
    afterFormulas: [[5, "changed"]], turnId: "turn-a" };
}

void test("concurrent migration preserves ids and afterFormulas; migrated backups restore and invert", async () => {
  const settings = createInMemorySettingsStore();
  const original = snapshot("legacy", 1);
  await settings.set(RECOVERY_SETTING_KEY, { version: 1, snapshots: [original] });
  const paneB = new RecoveryStore(settings.backend);
  await Promise.all([settings.recovery.migrate(), paneB.migrate()]);
  assert.equal(await settings.get(RECOVERY_SETTING_KEY), null);
  assert.deepEqual((await paneB.page("book-a")).entries.map((entry) => entry.id), ["legacy"]);
  assert.deepEqual(await paneB.read("legacy"), original);
  let applied: unknown[][] = [];
  const log = new WorkbookRecoveryLog({
    getStore: () => Promise.resolve(paneB),
    getWorkbookContext: () => Promise.resolve({ workbookId: "book-a", workbookName: "A.xlsx", source: "document.url" }),
    createId: () => "inverse", now: () => 2,
    readRangeFormulas: () => Promise.resolve([[99, "changed"]]),
    applySnapshot: (_address, values) => {
      applied = values;
      return Promise.resolve({ values: [[99, "changed"]], formulas: [[99, "changed"]] });
    },
  });
  assert.equal((await log.previewRestore("legacy")).laterEdits?.count, 1);
  const restored = await log.restore("legacy");
  assert.deepEqual(applied, [["=1+2", "Café 中文"]]);
  assert.equal(restored.inverseSnapshotId, "inverse");
  assert.deepEqual((await log.read("inverse"))?.beforeValues, [[99, "changed"]]);
});

void test("failed copy rolls back marker, payloads and legacy deletion without blocking settings", async () => {
  const settings = createInMemorySettingsStore();
  const legacy = { version: 1, snapshots: [snapshot("legacy", 1)] };
  await settings.set(RECOVERY_SETTING_KEY, legacy);
  const transaction = settings.backend.indexedTransaction.bind(settings.backend);
  let fail = true;
  settings.backend.indexedTransaction = (stores, mode, operation) => transaction(stores, mode, async (tx) => {
    const result = await operation(tx);
    if (fail && stores.includes("recovery-payloads") && stores.includes("settings")) {
      fail = false;
      throw new DOMException("Disk full", "QuotaExceededError");
    }
    return result;
  });
  await assert.rejects(settings.recovery.migrate(), /Disk full/);
  assert.deepEqual(await settings.get(RECOVERY_SETTING_KEY), legacy);
  await settings.set("ordinary-setting", "still works");
  assert.equal(await settings.get("ordinary-setting"), "still works");
  await settings.recovery.migrate();
  assert.deepEqual((await settings.recovery.page()).entries.map((entry) => entry.id), ["legacy"]);
  assert.equal(await settings.get(RECOVERY_SETTING_KEY), null);
});

void test("corrupt or oversized legacy data is retained, not silently trimmed on migration", async () => {
  for (const legacy of [
    { version: 1, snapshots: [snapshot("legacy", 1)] },
    { version: 1, snapshots: [{ id: "corrupt" }] },
  ]) {
    const settings = createInMemorySettingsStore();
    await settings.set(RECOVERY_SETTING_KEY, legacy);
    const tiny = new RecoveryStore(settings.backend, 1);
    await assert.rejects(tiny.migrate(), /original data is unchanged/);
    assert.deepEqual(await settings.get(RECOVERY_SETTING_KEY), legacy);
    const fallback = await tiny.page("book-a");
    assert.equal(fallback.entries.length, legacy.snapshots[0]?.id === "legacy" ? 1 : 0);
    assert.ok(fallback.entries.every((entry) => !("beforeValues" in entry)), "fallback pages also expose metadata only");
  }
});

void test("byte ceiling is inclusive; crossing it evicts globally oldest, with no per-workbook floor", async () => {
  const sample = createInMemorySettingsStore();
  await sample.recovery.append(snapshot("a", 1));
  await sample.recovery.append(snapshot("b", 2, "book-b"));
  const budget = (await sample.recovery.page()).entries.reduce((total, entry) => total + entry.sizeBytes, 0);
  const settings = createInMemorySettingsStore();
  const store = new RecoveryStore(settings.backend, budget);
  await store.append(snapshot("a", 1));
  await store.append(snapshot("b", 2, "book-b"));
  assert.deepEqual((await store.page()).entries.map((entry) => entry.id), ["b", "a"], "exactly at the ceiling keeps both");
  const newest = snapshot("c", 3, "book-b");
  delete newest.turnId; // Smaller than a; evicting a alone is enough to fit.
  await store.append(newest);
  assert.deepEqual((await store.page()).entries.map((entry) => entry.id), ["c", "b"]);
  assert.equal(await store.read("a"), null, "the only backup of book-a is not protected");
  await assert.rejects(new RecoveryStore(settings.backend, 1).append(snapshot("huge", 4)), /Too large to back up/);
  assert.deepEqual((await store.page()).entries.map((entry) => entry.id), ["c", "b"], "refusal must not evict anything");
  assert.equal(await store.delete("b", "book-a"), false);
  assert.equal(await store.clear("book-a"), 0);
  assert.equal(await store.clear("book-b"), 2);
  await store.append(snapshot("d", 4));
  assert.deepEqual((await store.page()).entries.map((entry) => entry.id), ["d"], "deletes release the byte budget");
});

void test("keyset pages stay bounded across timestamp ties, turn filtering, appends and deletes", async () => {
  const { recovery: store } = createInMemorySettingsStore();
  for (let i = 0; i < 125; i += 1) {
    await store.append({ ...snapshot(`backup-${String(i).padStart(3, "0")}`, 1), turnId: i % 2 ? "odd" : "even" });
  }
  const first = await store.page("book-a");
  assert.equal(first.entries.length, 50);
  assert.equal(first.entries[0]?.id, "backup-124");
  assert.equal("beforeValues" in (first.entries[0] ?? {}), false);
  assert.ok(first.nextCursor);
  await store.delete(first.nextCursor.id, "book-a");
  await store.append(snapshot("newer", 2));
  const rest: string[] = [];
  let cursor: RecoveryCursor | undefined = first.nextCursor;
  do {
    const page = await store.page("book-a", { cursor });
    rest.push(...page.entries.map((entry) => entry.id));
    cursor = page.nextCursor;
  } while (cursor);
  assert.deepEqual(rest, Array.from({ length: 75 }, (_, i) => `backup-${String(74 - i).padStart(3, "0")}`));
  const odd = await store.page("book-a", { turnId: "odd" });
  assert.equal(odd.entries.length, 50);
  assert.ok(odd.entries.every((entry) => entry.turnId === "odd"));
  const oldest = await store.page("book-a", { oldestFirst: true, limit: 2 });
  assert.deepEqual(oldest.entries.map((entry) => entry.id), ["backup-000", "backup-001"]);
});

void test("identity moves merge metadata without copying payloads and preserve restore guards", async () => {
  const { recovery: store } = createInMemorySettingsStore();
  await store.append(snapshot("a", 1));
  await store.append(snapshot("b", 2, "book-b"));
  assert.equal(await store.move("book-a", "book-b", "Renamed.xlsx"), 1);
  assert.equal(await store.move("book-a", "book-b", "Renamed.xlsx"), 0);
  assert.deepEqual((await store.page("book-a")).entries, []);
  assert.deepEqual((await store.page("book-b")).entries.map((entry) => entry.id), ["b", "a"]);
  assert.equal((await store.read("a", "book-b"))?.workbookLabel, "Renamed.xlsx");
  assert.deepEqual((await store.read("a", "book-b"))?.afterFormulas, [[5, "changed"]]);
  await assert.rejects(store.read("a", "book-a"), /different workbook/);
});
