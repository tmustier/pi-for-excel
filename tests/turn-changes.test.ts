import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import { Value } from "typebox/value";
import { readTurnBackups, summarizeTurnChanges, undoTurnChanges } from "../src/messages/turn-changes.ts";
import { createConvertToLlm } from "../src/messages/convert-to-llm.ts";
import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import { createInMemorySettingsStore } from "./fixtures/recovery-log.ts";

void test("finished, stopped and failed receipts undo overlapping changes newest first and keep inverse backups", async () => {
  for (const outcome of ["finished", "stopped", "failed"] as const) {
    let value = 91;
    let serial = 0;
    const settings = createInMemorySettingsStore();
    const log = new WorkbookRecoveryLog({
      getStore: () => Promise.resolve(settings.recovery),
      getWorkbookContext: () => Promise.resolve({ workbookId: "url_sha256:turn-test", workbookName: "Turn.xlsx", source: "document.url" }),
      createId: () => `backup-${++serial}`,
      now: () => serial,
      readRangeFormulas: () => Promise.resolve([[value]]),
      applySnapshot: (_address, values) => {
        assert.ok(Value.Check(Type.Array(Type.Array(Type.Number())), values));
        const before = value;
        value = values[0]?.[0] ?? value;
        return Promise.resolve({ values: [[before]], formulas: [[before]] });
      },
    });
    const older = await log.append({ toolName: "write_cells", toolCallId: "first", address: "Data!B2", changedCount: 1,
      beforeValues: [[7]], beforeFormulas: [[7]], afterFormulas: [[23]] });
    const newer = await log.append({ toolName: "write_cells", toolCallId: "second", address: "Data!B2", changedCount: 1,
      beforeValues: [[23]], beforeFormulas: [[23]], afterFormulas: [[91]] });
    assert.ok(older && newer);
    const summary = summarizeTurnChanges({ turnId: "turn-1", timestamp: 10, outcome,
      toolCallIds: new Set(["first", "second"]), audit: [], backups: [newer, older] });
    assert.ok(summary);
    assert.equal(summary.outcome, outcome);
    assert.deepEqual(createConvertToLlm()([summary]), [], "UI receipts never alter model context");
    value = 999;
    assert.equal(await undoTurnChanges(summary, async (id) => {
      const preview = await log.previewRestore(id);
      return preview.laterEdits ? false : (await log.restore(id), true);
    }), 0, "later edits can cancel without changing the workbook");
    assert.equal(value, 999);
    value = 91;
    assert.equal(await undoTurnChanges(summary, async (id) => {
      assert.equal((await log.previewRestore(id)).laterEdits, null);
      const restored = await log.restore(id);
      assert.ok(restored.inverseSnapshotId, "undo itself is backed up");
      return true;
    }), 2);
    assert.equal(value, 7, "oldest pre-turn value is restored, not the intermediate value");
    assert.equal(await undoTurnChanges(summary, () => Promise.reject(new Error("must not undo twice"))), 0);
    const inverse = (await log.listForCurrentWorkbook(1))[0];
    assert.ok(inverse);
    await log.restore(inverse.id);
    assert.equal(value, 23, "ordinary /revert can undo a restore");
  }
});

void test("turn receipts and undo include every page, exclude other turns, and preserve partial progress", async () => {
  const settings = createInMemorySettingsStore();
  let serial = 0;
  const log = new WorkbookRecoveryLog({
    getStore: () => Promise.resolve(settings.recovery),
    getWorkbookContext: () => Promise.resolve({ workbookId: "url_sha256:paging", workbookName: "Paging.xlsx", source: "document.url" }),
    createId: () => `backup-${String(++serial).padStart(3, "0")}`,
    now: () => serial,
  });
  const calls = new Set<string>();
  for (let i = 0; i < 57; i++) {
    const call = `call-${i}`;
    calls.add(call);
    await log.append({ turnId: "large-turn", toolName: "write_cells", toolCallId: call,
      address: `Data!A${i + 1}`, changedCount: 1, beforeValues: [[i]], beforeFormulas: [[i]], afterFormulas: [[i + 1]] });
  }
  await log.append({ turnId: "other-turn", toolName: "write_cells", toolCallId: "other",
    address: "Data!B1", changedCount: 1, beforeValues: [[0]], beforeFormulas: [[0]], afterFormulas: [[1]] });
  const backups = await readTurnBackups(log, "large-turn");
  assert.equal(backups.length, 57);
  const summary = summarizeTurnChanges({ turnId: "large-turn", timestamp: 100, outcome: "failed",
    toolCallIds: calls, audit: [], backups });
  assert.ok(summary);
  assert.equal(summary.changes.length, 57, "partial failed writes remain reviewable without an audit success");
  const restored: string[] = [];
  assert.equal(await undoTurnChanges(summary, (id) => {
    if (restored.length === 3) return Promise.resolve(false);
    restored.push(id);
    return Promise.resolve(true);
  }), 3);
  assert.equal(await undoTurnChanges(summary, (id) => {
    restored.push(id);
    return Promise.resolve(true);
  }), 54);
  assert.deepEqual(restored, Array.from({ length: 57 }, (_, i) => `backup-${String(57 - i).padStart(3, "0")}`));
});
