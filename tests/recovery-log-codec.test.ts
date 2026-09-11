import assert from "node:assert/strict";
import { test } from "node:test";

import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import { RECOVERY_SETTING_KEY, createInMemorySettingsStore } from "./fixtures/recovery-log.ts";

const workbook: WorkbookContext = {
  workbookId: "url_sha256:codec",
  workbookName: "Codec.xlsx",
  source: "document.url",
};

async function loadFromPersisted(snapshots: unknown[]): Promise<{
  log: WorkbookRecoveryLog;
  ids: string[];
  persisted: () => Promise<unknown>;
}> {
  const settingsStore = createInMemorySettingsStore();
  await settingsStore.set(RECOVERY_SETTING_KEY, { version: 1, snapshots });

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve(workbook),
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });
  const loaded = await log.list({ limit: 50 });
  return {
    log,
    ids: loaded.map((snapshot) => snapshot.id),
    persisted: () => settingsStore.get(RECOVERY_SETTING_KEY),
  };
}

const base = {
  toolName: "write_cells",
  toolCallId: "call-1",
  address: "Sheet1!A1:B2",
  workbookId: workbook.workbookId,
};

const dataBar = {
  axisFormat: "Automatic",
  barDirection: "Context",
  showDataBarOnly: false,
  lowerBoundRule: { type: "LowestValue" },
  upperBoundRule: { type: "HighestValue", formula: "100" },
  positiveFillColor: "#63C384",
  positiveGradientFill: true,
  negativeFillColor: "#D13438",
  negativeMatchPositiveFillColor: false,
  negativeMatchPositiveBorderColor: false,
};

const iconSet = {
  style: "ThreeTrafficLights1",
  reverseIconOrder: false,
  showIconOnly: false,
  criteria: [
    { type: "Percent", operator: "GreaterThanOrEqual", formula: "0" },
    { type: "Percent", operator: "GreaterThanOrEqual", formula: "33", customIcon: { set: "ThreeTrafficLights1", index: 1 } },
  ],
};

void test("a range snapshot written before snapshotKind existed loads with defaults filled in", async () => {
  const before = Date.now();
  const { log, ids } = await loadFromPersisted([
    { ...base, beforeValues: [[1, 2], [3, 4]], beforeFormulas: [["", ""], ["", "=A1"]] },
  ]);

  assert.equal(ids.length, 1);
  const [snapshot] = await log.list({ limit: 1 });
  assert.ok(snapshot);
  assert.equal(snapshot.snapshotKind, "range_values");
  assert.equal(snapshot.cellCount, 4);
  assert.equal(snapshot.changedCount, 4);
  assert.ok(snapshot.id.length > 0);
  assert.ok(snapshot.at >= before);
});

void test("each snapshot kind is kept only with its own well-formed state", async () => {
  const validArea = { address: "Sheet1!A1:B2", rowCount: 2, columnCount: 2, numberFormat: [["0", "0"], ["0", "0"]], columnWidths: [10, 10] };
  const { ids } = await loadFromPersisted([
    { ...base, id: "range-ok", at: 10, beforeValues: [[1]], beforeFormulas: [[""]] },
    { ...base, id: "range-no-grid", at: 9, snapshotKind: "range_values" },
    { ...base, id: "format-ok", at: 8, snapshotKind: "format_cells_state", formatRangeState: { selection: { numberFormat: true }, areas: [validArea], cellCount: 4 } },
    { ...base, id: "format-grid-mismatch", at: 7, snapshotKind: "format_cells_state", formatRangeState: { selection: {}, areas: [{ ...validArea, numberFormat: [["0", "0"]] }], cellCount: 4 } },
    { ...base, id: "format-widths-mismatch", at: 6, snapshotKind: "format_cells_state", formatRangeState: { selection: {}, areas: [{ ...validArea, columnWidths: [10] }], cellCount: 4 } },
    { ...base, id: "structure-ok", at: 5, snapshotKind: "modify_structure_state", modifyStructureState: { kind: "rows_present", sheetId: "s1", sheetName: "Sheet1", position: 2, count: 1, dataRange: { address: "Sheet1!A2:B2", rowCount: 1, columnCount: 2, values: [[1, 2]], formulas: [["", ""]] } } },
    { ...base, id: "structure-grid-mismatch", at: 4, snapshotKind: "modify_structure_state", modifyStructureState: { kind: "rows_present", sheetId: "s1", sheetName: "Sheet1", position: 2, count: 1, dataRange: { address: "Sheet1!A2:B2", rowCount: 1, columnCount: 2, values: [[1]], formulas: [["", ""]] } } },
    { ...base, id: "structure-zero-count", at: 3, snapshotKind: "modify_structure_state", modifyStructureState: { kind: "rows_absent", sheetId: "s1", sheetName: "Sheet1", position: 2, count: 0 } },
    { ...base, id: "rules-ok", at: 2, snapshotKind: "conditional_format_rules", conditionalFormatRules: [{ type: "data_bar", dataBar }, { type: "icon_set", iconSet }, { type: "cell_value", operator: "GreaterThan", formula1: "5" }] },
    { ...base, id: "rules-bad-formula-type", at: 1.5, snapshotKind: "conditional_format_rules", conditionalFormatRules: [{ type: "data_bar", dataBar: { ...dataBar, lowerBoundRule: { type: "LowestValue", formula: 42 } } }] },
    { ...base, id: "rules-bad-operator", at: 1.4, snapshotKind: "conditional_format_rules", conditionalFormatRules: [{ type: "icon_set", iconSet: { ...iconSet, criteria: [{ type: "Percent", operator: "LessThan", formula: "0" }] } }] },
    { ...base, id: "rules-missing-required", at: 1.3, snapshotKind: "conditional_format_rules", conditionalFormatRules: [{ type: "cell_value", operator: "GreaterThan" }] },
    { ...base, id: "comment-ok", at: 1.2, snapshotKind: "comment_thread", commentThreadState: { exists: true, content: "hi", resolved: false, replies: [] } },
    { ...base, id: "chart-ok-no-grids", at: 1.1, snapshotKind: "chart_state", chartState: { kind: "chart_absent", sheetName: "Sheet1", name: "Chart 1" } },
    { ...base, id: "chart-wrong-state", at: 1.0, snapshotKind: "chart_state", commentThreadState: { exists: true, content: "hi", resolved: false, replies: [] } },
    { ...base, id: "unknown-kind", at: 0.9, snapshotKind: "pivot_state", beforeValues: [[1]], beforeFormulas: [[""]] },
    { ...base, id: "unknown-tool", at: 0.8, toolName: "pivot_tables", beforeValues: [[1]], beforeFormulas: [[""]] },
  ]);

  assert.deepEqual(ids, ["range-ok", "format-ok", "structure-ok", "rules-ok", "comment-ok", "chart-ok-no-grids"]);
});

void test("a loaded snapshot carries its state and drops properties the schema does not declare", async () => {
  const { log, persisted } = await loadFromPersisted([
    {
      ...base,
      id: "rules",
      at: 2,
      snapshotKind: "conditional_format_rules",
      conditionalFormatRules: [{ type: "data_bar", dataBar: { ...dataBar, legacyField: "x" } }],
      leftoverFromOlderBuild: { big: true },
    },
    { ...base, id: "chart", at: 1, snapshotKind: "chart_state", chartState: { kind: "chart_absent", sheetName: "Sheet1", name: "Chart 1" } },
  ]);

  const [rules, chart] = await log.list({ limit: 2 });
  assert.ok(rules && chart);
  assert.deepEqual(rules.conditionalFormatRules, [{ type: "data_bar", dataBar }]);
  assert.equal(rules.cellCount, 1);
  assert.deepEqual(chart.beforeValues, []);
  assert.equal(chart.cellCount, 1);

  await log.clearForCurrentWorkbook();
  await log.append({ toolName: "write_cells", toolCallId: "call-new", address: "Sheet1!A1", beforeValues: [[1]], beforeFormulas: [[""]] });
  assert.doesNotMatch(JSON.stringify(await persisted()), /leftoverFromOlderBuild|legacyField/);
});

void test("a payload that is not the versioned envelope loads as empty", async () => {
  const settingsStore = createInMemorySettingsStore();
  await settingsStore.set(RECOVERY_SETTING_KEY, { snapshots: [{ ...base, beforeValues: [[1]], beforeFormulas: [[""]] }] });

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve(workbook),
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });

  assert.deepEqual(await log.list({ limit: 10 }), []);
});
