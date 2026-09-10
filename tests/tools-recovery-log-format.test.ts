import assert from "node:assert/strict";
import { test } from "node:test";

import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import type { RecoveryFormatRangeState } from "../src/workbook/recovery-states.ts";
import { RECOVERY_SETTING_KEY } from "./fixtures/recovery-log.ts";

function recoveryLog(formatRangeState: DynamicValue, apply: (state: RecoveryFormatRangeState) => void): WorkbookRecoveryLog {
  const payload = {
    version: 1,
    snapshots: [{
      id: "persisted-format",
      at: 1,
      toolName: "format_cells",
      toolCallId: "original-call",
      address: "Sheet1!A1:B2",
      changedCount: 4,
      cellCount: 4,
      beforeValues: [],
      beforeFormulas: [],
      snapshotKind: "format_cells_state",
      formatRangeState,
      workbookId: "book-format",
    }],
  };
  return new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve({
      get: <T>(key: string): Promise<T | null> => Promise.resolve(key === RECOVERY_SETTING_KEY ? payload as T : null),
      set: () => Promise.resolve(),
    }),
    getWorkbookContext: () => Promise.resolve({
      workbookId: "book-format",
      workbookName: "Formatting.xlsx",
      source: "document.url",
    }),
    applySnapshot: () => Promise.resolve({ values: [], formulas: [] }),
    applyFormatCellsSnapshot: (_address, state) => {
      apply(state);
      return Promise.resolve(state);
    },
  });
}

void test("restore accepts a persisted format grid whose dimensions match its range", async () => {
  let restored: RecoveryFormatRangeState | undefined;
  const state = {
    selection: { numberFormat: true },
    areas: [{
      address: "Sheet1!A1:B2",
      rowCount: 2,
      columnCount: 2,
      numberFormat: [["0.00", "General"], ["General", "0.00"]],
    }],
    cellCount: 4,
  };

  const result = await recoveryLog(state, (value) => { restored = value; }).restore("persisted-format");
  assert.equal(result.address, "Sheet1!A1:B2");
  assert.deepEqual(restored?.areas[0]?.numberFormat, [["0.00", "General"], ["General", "0.00"]]);
});

void test("restore rejects a persisted format grid whose dimensions do not match its range", async () => {
  const malformed = {
    selection: { numberFormat: true },
    areas: [{
      address: "Sheet1!A1:B2",
      rowCount: 2,
      columnCount: 2,
      numberFormat: [["0.00", "General"]],
    }],
    cellCount: 4,
  };

  await assert.rejects(
    () => recoveryLog(malformed, () => { throw new Error("malformed state reached workbook host"); }).restore("persisted-format"),
    /Snapshot not found/u,
  );
});
