import assert from "node:assert/strict";
import { test } from "node:test";

import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import {
  estimateFormatCaptureCellCount,
  firstCellAddress,
  type RecoveryConditionalFormatRule,
  type RecoveryFormatRangeState,
} from "../src/workbook/recovery-states.ts";
import {
  isRecoveryHorizontalAlignment,
  isRecoveryUnderlineStyle,
  isRecoveryVerticalAlignment,
  normalizeOptionalBoolean,
  normalizeOptionalNumber,
  normalizeOptionalString,
} from "../src/workbook/recovery/format-state-normalization.ts";
import {
  collectMergedAreaAddresses,
  dedupeRecoveryAddresses,
  validateStringGrid,
} from "../src/workbook/recovery/format-state-utils.ts";
import {
  isRecoveryConditionalDataBarState,
  isRecoveryConditionalIconSetState,
  normalizeConditionalFormatType,
} from "../src/workbook/recovery/conditional-format-normalization.ts";
import {
  createInMemorySettingsStore,
  findSnapshotById,
  withoutUndefined,
} from "./recovery-log-test-helpers.test.ts";

void test("firstCellAddress handles quoted sheet names that include !", () => {
  assert.equal(firstCellAddress("'Q1!Ops'!A1"), "A1");
  assert.equal(firstCellAddress("'Q1!Ops'!$B$2:$D$9"), "$B$2");
  assert.equal(firstCellAddress("Sheet1!C5:D7"), "C5");
});
void test("estimateFormatCaptureCellCount scales by serialized checkpoint shape", () => {
  const largeArea = [{ rowCount: 1_048_576, columnCount: 3 }];

  assert.equal(
    estimateFormatCaptureCellCount(largeArea, { columnWidth: true }),
    3,
  );

  assert.equal(
    estimateFormatCaptureCellCount(largeArea, { rowHeight: true }),
    1_048_576,
  );

  assert.equal(
    estimateFormatCaptureCellCount(largeArea, { columnWidth: true, rowHeight: true }),
    1_048_579,
  );

  assert.equal(
    estimateFormatCaptureCellCount(largeArea, { columnWidth: true, fillColor: true }),
    4,
  );

  assert.equal(
    estimateFormatCaptureCellCount(largeArea, { mergedAreas: true }),
    1_572_864,
  );

  assert.equal(
    estimateFormatCaptureCellCount(largeArea, { mergedAreas: true, rowHeight: true }),
    2_621_440,
  );
});

void test("format-state normalization guards accept only supported values", () => {
  assert.equal(isRecoveryUnderlineStyle("Single"), true);
  assert.equal(isRecoveryUnderlineStyle("DoubleAccountant"), true);
  assert.equal(isRecoveryUnderlineStyle("Wave"), false);

  assert.equal(isRecoveryHorizontalAlignment("General"), true);
  assert.equal(isRecoveryHorizontalAlignment("CenterAcrossSelection"), true);
  assert.equal(isRecoveryHorizontalAlignment("DistributedAcrossSelection"), false);

  assert.equal(isRecoveryVerticalAlignment("Top"), true);
  assert.equal(isRecoveryVerticalAlignment("Distributed"), true);
  assert.equal(isRecoveryVerticalAlignment("Middle"), false);
});

void test("format-state normalization keeps only optional scalar values", () => {
  assert.equal(normalizeOptionalString("abc"), "abc");
  assert.equal(normalizeOptionalString(123), undefined);

  assert.equal(normalizeOptionalBoolean(true), true);
  assert.equal(normalizeOptionalBoolean("true"), undefined);

  assert.equal(normalizeOptionalNumber(12.5), 12.5);
  assert.equal(normalizeOptionalNumber(Number.NaN), undefined);
  assert.equal(normalizeOptionalNumber(Infinity), undefined);
});

void test("format-state utilities dedupe addresses and merged areas deterministically", () => {
  assert.deepEqual(
    dedupeRecoveryAddresses([" Sheet1!A1:A2 ", "Sheet1!A1:A2", "", "Sheet1!B1:B2"]),
    ["Sheet1!A1:A2", "Sheet1!B1:B2"],
  );

  const merged = collectMergedAreaAddresses({
    selection: { mergedAreas: true },
    areas: [
      { address: "Sheet1!A1:B2", rowCount: 2, columnCount: 2, mergedAreas: ["Sheet1!A1:B1"] },
      { address: "Sheet1!C1:D2", rowCount: 2, columnCount: 2, mergedAreas: ["Sheet1!A1:B1", "Sheet1!C1:D1"] },
    ],
    cellCount: 4,
  });

  assert.deepEqual(merged, ["Sheet1!A1:B1", "Sheet1!C1:D1"]);
});

void test("format-state utilities validate string grid shape", () => {
  assert.deepEqual(
    validateStringGrid([["0.00", "General"], ["General", "0.00"]], 2, 2),
    [["0.00", "General"], ["General", "0.00"]],
  );

  assert.equal(validateStringGrid([["0.00"]], 2, 1), null);
  assert.equal(validateStringGrid([["0.00", 42]], 1, 2), null);
});

void test("conditional-format normalization maps supported Excel types", () => {
  assert.equal(normalizeConditionalFormatType("Custom"), "custom");
  assert.equal(normalizeConditionalFormatType("colorScale"), "color_scale");
  assert.equal(normalizeConditionalFormatType("IconSet"), "icon_set");
  assert.equal(normalizeConditionalFormatType("NotSupported"), null);
});

void test("conditional-format guards validate data-bar and icon-set rule state", () => {
  assert.equal(
    isRecoveryConditionalDataBarState({
      axisColor: "#000000",
      axisFormat: "Automatic",
      barDirection: "Context",
      showDataBarOnly: false,
      lowerBoundRule: { type: "LowestValue" },
      upperBoundRule: { type: "HighestValue", formula: "100" },
      positiveFillColor: "#63C384",
      positiveBorderColor: "#2E8540",
      positiveGradientFill: true,
      negativeFillColor: "#D13438",
      negativeBorderColor: "#A4262C",
      negativeMatchPositiveFillColor: false,
      negativeMatchPositiveBorderColor: false,
    }),
    true,
  );

  assert.equal(
    isRecoveryConditionalDataBarState({
      axisFormat: "Automatic",
      barDirection: "Context",
      showDataBarOnly: false,
      lowerBoundRule: { type: "LowestValue", formula: 42 },
      upperBoundRule: { type: "HighestValue" },
      positiveFillColor: "#63C384",
      positiveGradientFill: true,
      negativeFillColor: "#D13438",
      negativeMatchPositiveFillColor: false,
      negativeMatchPositiveBorderColor: false,
    }),
    false,
  );

  assert.equal(
    isRecoveryConditionalIconSetState({
      style: "ThreeTrafficLights1",
      reverseIconOrder: false,
      showIconOnly: false,
      criteria: [
        { type: "Percent", operator: "GreaterThanOrEqual", formula: "0" },
        {
          type: "Percent",
          operator: "GreaterThanOrEqual",
          formula: "33",
          customIcon: { set: "ThreeTrafficLights1", index: 1 },
        },
      ],
    }),
    true,
  );

  assert.equal(
    isRecoveryConditionalIconSetState({
      style: "ThreeTrafficLights1",
      reverseIconOrder: false,
      showIconOnly: false,
      criteria: [
        { type: "Percent", operator: "LessThan", formula: "0" },
      ],
    }),
    false,
  );
});

void test("persisted format checkpoints retain dimension state", async () => {
  const settingsStore = createInMemorySettingsStore();

  const getWorkbookContext = (): Promise<WorkbookContext> => Promise.resolve({
    workbookId: "url_sha256:workbook-format-persist",
    workbookName: "Ops.xlsx",
    source: "document.url",
  });

  const formatState: RecoveryFormatRangeState = {
    selection: {
      columnWidth: true,
      rowHeight: true,
      mergedAreas: true,
    },
    areas: [
      {
        address: "Sheet1!A1:B2",
        rowCount: 2,
        columnCount: 2,
        columnWidths: [64, 80],
        rowHeights: [18, 22],
        mergedAreas: ["Sheet1!A1:B1"],
      },
    ],
    cellCount: 4,
  };

  const logA = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
    now: () => 1700000000100,
    createId: () => "snap-format-persist-1",
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  const appended = await logA.appendFormatCells({
    toolName: "format_cells",
    toolCallId: "call-format-persist",
    address: "Sheet1!A1:B2",
    changedCount: 4,
    formatRangeState: formatState,
  });

  assert.ok(appended);

  const logB = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  const entries = await logB.listForCurrentWorkbook(10);
  assert.equal(entries.length, 1);
  assert.equal(entries[0]?.snapshotKind, "format_cells_state");
  assert.deepEqual(withoutUndefined(entries[0]?.formatRangeState), withoutUndefined(formatState));
});
void test("persisted conditional-format checkpoints retain extended rule types", async () => {
  const settingsStore = createInMemorySettingsStore();

  const getWorkbookContext = (): Promise<WorkbookContext> => Promise.resolve({
    workbookId: "url_sha256:workbook-conditional-format-persist",
    workbookName: "Ops.xlsx",
    source: "document.url",
  });

  const rules = [
    {
      type: "custom",
      formula: "=A1>10",
      fillColor: "#FF0000",
      appliesToAddress: "Sheet1!A1:A2",
    },
    {
      type: "cell_value",
      operator: "GreaterThan",
      formula1: "10",
      fillColor: "#0000FF",
      appliesToAddress: "Sheet1!B1:B2",
    },
    {
      type: "text_comparison",
      textOperator: "Contains",
      text: "urgent",
      fillColor: "#FFE599",
      appliesToAddress: "Sheet1!C1:C2",
    },
    {
      type: "top_bottom",
      topBottomType: "TopItems",
      rank: 3,
      fillColor: "#E2EFDA",
      appliesToAddress: "Sheet1!D1:D10",
    },
    {
      type: "preset_criteria",
      presetCriterion: "DuplicateValues",
      fillColor: "#FCE4D6",
      appliesToAddress: "Sheet1!E1:E10",
    },
    {
      type: "data_bar",
      stopIfTrue: true,
      appliesToAddress: "Sheet1!F1:F10",
      dataBar: {
        axisColor: "#000000",
        axisFormat: "Automatic",
        barDirection: "Context",
        showDataBarOnly: false,
        lowerBoundRule: { type: "LowestValue" },
        upperBoundRule: { type: "HighestValue" },
        positiveFillColor: "#63C384",
        positiveBorderColor: "#2E8540",
        positiveGradientFill: true,
        negativeFillColor: "#D13438",
        negativeBorderColor: "#A4262C",
        negativeMatchPositiveFillColor: false,
        negativeMatchPositiveBorderColor: false,
      },
    },
    {
      type: "color_scale",
      stopIfTrue: false,
      appliesToAddress: "Sheet1!G1:G10",
      colorScale: {
        minimum: { type: "LowestValue", color: "#F8696B" },
        midpoint: { type: "Percentile", formula: "50", color: "#FFEB84" },
        maximum: { type: "HighestValue", color: "#63BE7B" },
      },
    },
    {
      type: "icon_set",
      stopIfTrue: true,
      appliesToAddress: "Sheet1!H1:H10",
      iconSet: {
        style: "ThreeTrafficLights1",
        reverseIconOrder: false,
        showIconOnly: false,
        criteria: [
          { type: "Percent", operator: "GreaterThanOrEqual", formula: "0" },
          { type: "Percent", operator: "GreaterThanOrEqual", formula: "33" },
          { type: "Percent", operator: "GreaterThanOrEqual", formula: "67" },
        ],
      },
    },
  ] as const;

  const logA = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
    now: () => 1700000000150,
    createId: () => "snap-conditional-format-persist-1",
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  const appended = await logA.appendConditionalFormat({
    toolName: "conditional_format",
    toolCallId: "call-conditional-format-persist",
    address: "Sheet1!A1:E10",
    changedCount: 50,
    cellCount: 50,
    conditionalFormatRules: [...rules],
  });

  assert.ok(appended);

  const logB = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  const entries = await logB.listForCurrentWorkbook(10);
  assert.equal(entries.length, 1);
  assert.equal(entries[0]?.snapshotKind, "conditional_format_rules");
  assert.deepEqual(withoutUndefined(entries[0]?.conditionalFormatRules), withoutUndefined(rules));
});
void test("restore applies format-cells checkpoints and creates inverse checkpoint", async () => {
  const settingsStore = createInMemorySettingsStore();

  const workbookContext: WorkbookContext = {
    workbookId: "url_sha256:workbook-format",
    workbookName: "Formatting.xlsx",
    source: "document.url",
  };

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-format-${idCounter}`;
  };

  let appliedAddress = "";
  let appliedState: RecoveryFormatRangeState | null = null;

  const restoredTargetState: RecoveryFormatRangeState = {
    selection: {
      numberFormat: true,
      fillColor: true,
      bold: true,
      columnWidth: true,
      rowHeight: true,
      mergedAreas: true,
      borderTop: true,
    },
    areas: [
      {
        address: "Sheet1!A1:B1",
        rowCount: 1,
        columnCount: 2,
        numberFormat: [["0.00", "0.00"]],
        fillColor: "#FFFF00",
        bold: true,
        columnWidths: [64, 80],
        rowHeights: [24],
        mergedAreas: ["Sheet1!A1:B1"],
        borderTop: {
          style: "Continuous",
          weight: "Thin",
          color: "#000000",
        },
      },
    ],
    cellCount: 2,
  };

  const currentFormatState: RecoveryFormatRangeState = {
    selection: {
      numberFormat: true,
      fillColor: true,
      bold: true,
      columnWidth: true,
      rowHeight: true,
      mergedAreas: true,
      borderTop: true,
    },
    areas: [
      {
        address: "Sheet1!A1:B1",
        rowCount: 1,
        columnCount: 2,
        numberFormat: [["General", "General"]],
        fillColor: "#FFFFFF",
        bold: false,
        columnWidths: [72, 72],
        rowHeights: [20],
        mergedAreas: [],
        borderTop: {
          style: "None",
        },
      },
    ],
    cellCount: 2,
  };

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve(workbookContext),
    now: () => 1700000001800,
    createId,
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
    applyFormatCellsSnapshot: (address, state) => {
      appliedAddress = address;
      appliedState = state;
      return Promise.resolve(currentFormatState);
    },
  });

  const appended = await log.appendFormatCells({
    toolName: "format_cells",
    toolCallId: "call-format",
    address: "Sheet1!A1:B1",
    changedCount: 2,
    formatRangeState: restoredTargetState,
  });

  assert.ok(appended);

  const restored = await log.restore(appended?.id ?? "");

  assert.equal(restored.address, "Sheet1!A1:B1");
  assert.equal(restored.restoredSnapshotId, appended?.id);
  assert.equal(appliedAddress, "Sheet1!A1:B1");
  assert.deepEqual(withoutUndefined(appliedState), withoutUndefined(restoredTargetState));

  const snapshots = await log.listForCurrentWorkbook(10);
  const inverse = restored.inverseSnapshotId
    ? findSnapshotById(snapshots, restored.inverseSnapshotId)
    : null;

  assert.ok(inverse);
  assert.equal(inverse?.toolName, "restore_snapshot");
  assert.equal(inverse?.snapshotKind, "format_cells_state");
  assert.equal(inverse?.restoredFromSnapshotId, appended?.id);
  assert.deepEqual(withoutUndefined(inverse?.formatRangeState), withoutUndefined(currentFormatState));
});
void test("restore applies conditional-format checkpoints and creates inverse checkpoint", async () => {
  const settingsStore = createInMemorySettingsStore();

  const workbookContext: WorkbookContext = {
    workbookId: "url_sha256:workbook-cf",
    workbookName: "Formatting.xlsx",
    source: "document.url",
  };

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-cf-${idCounter}`;
  };

  let appliedAddress = "";
  const appliedRules: unknown[] = [];

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve(workbookContext),
    now: () => 1700000002000,
    createId,
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
    applyConditionalFormatSnapshot: (address, rules) => {
      appliedAddress = address;
      appliedRules.push(...rules);
      return Promise.resolve({
        supported: true,
        rules: [{
          type: "custom",
          formula: "=A1>0",
          fillColor: "#00FF00",
          appliesToAddress: "Sheet1!A1:A2",
        }],
      });
    },
  });

  const appended = await log.appendConditionalFormat({
    toolName: "conditional_format",
    toolCallId: "call-cf",
    address: "Sheet1!A1:B2",
    changedCount: 4,
    cellCount: 4,
    conditionalFormatRules: [
      {
        type: "custom",
        formula: "=A1>10",
        fillColor: "#FF0000",
        appliesToAddress: "Sheet1!A1:A2",
      },
      {
        type: "cell_value",
        operator: "GreaterThan",
        formula1: "10",
        fillColor: "#0000FF",
        appliesToAddress: "Sheet1!B1:B2",
      },
      {
        type: "text_comparison",
        textOperator: "Contains",
        text: "urgent",
        fillColor: "#FFE599",
        appliesToAddress: "Sheet1!C1:C2",
      },
      {
        type: "top_bottom",
        topBottomType: "TopItems",
        rank: 3,
        fillColor: "#E2EFDA",
        appliesToAddress: "Sheet1!D1:D10",
      },
      {
        type: "preset_criteria",
        presetCriterion: "DuplicateValues",
        fillColor: "#FCE4D6",
        appliesToAddress: "Sheet1!E1:E10",
      },
      {
        type: "data_bar",
        stopIfTrue: true,
        appliesToAddress: "Sheet1!F1:F10",
        dataBar: {
          axisColor: "#000000",
          axisFormat: "Automatic",
          barDirection: "Context",
          showDataBarOnly: false,
          lowerBoundRule: { type: "LowestValue" },
          upperBoundRule: { type: "HighestValue" },
          positiveFillColor: "#63C384",
          positiveBorderColor: "#2E8540",
          positiveGradientFill: true,
          negativeFillColor: "#D13438",
          negativeBorderColor: "#A4262C",
          negativeMatchPositiveFillColor: false,
          negativeMatchPositiveBorderColor: false,
        },
      },
      {
        type: "color_scale",
        stopIfTrue: false,
        appliesToAddress: "Sheet1!G1:G10",
        colorScale: {
          minimum: { type: "LowestValue", color: "#F8696B" },
          midpoint: { type: "Percentile", formula: "50", color: "#FFEB84" },
          maximum: { type: "HighestValue", color: "#63BE7B" },
        },
      },
      {
        type: "icon_set",
        stopIfTrue: true,
        appliesToAddress: "Sheet1!H1:H10",
        iconSet: {
          style: "ThreeTrafficLights1",
          reverseIconOrder: false,
          showIconOnly: false,
          criteria: [
            { type: "Percent", operator: "GreaterThanOrEqual", formula: "0" },
            { type: "Percent", operator: "GreaterThanOrEqual", formula: "33" },
            { type: "Percent", operator: "GreaterThanOrEqual", formula: "67" },
          ],
        },
      },
    ],
  });

  assert.ok(appended);

  const restored = await log.restore(appended?.id ?? "");

  assert.equal(restored.address, "Sheet1!A1:B2");
  assert.equal(restored.restoredSnapshotId, appended?.id);
  assert.equal(appliedAddress, "Sheet1!A1:B2");
  assert.equal(appliedRules.length, 8);
  assert.deepEqual(
    withoutUndefined(appliedRules),
    withoutUndefined([
      {
        type: "custom",
        formula: "=A1>10",
        fillColor: "#FF0000",
        appliesToAddress: "Sheet1!A1:A2",
      },
      {
        type: "cell_value",
        operator: "GreaterThan",
        formula1: "10",
        fillColor: "#0000FF",
        appliesToAddress: "Sheet1!B1:B2",
      },
      {
        type: "text_comparison",
        textOperator: "Contains",
        text: "urgent",
        fillColor: "#FFE599",
        appliesToAddress: "Sheet1!C1:C2",
      },
      {
        type: "top_bottom",
        topBottomType: "TopItems",
        rank: 3,
        fillColor: "#E2EFDA",
        appliesToAddress: "Sheet1!D1:D10",
      },
      {
        type: "preset_criteria",
        presetCriterion: "DuplicateValues",
        fillColor: "#FCE4D6",
        appliesToAddress: "Sheet1!E1:E10",
      },
      {
        type: "data_bar",
        stopIfTrue: true,
        appliesToAddress: "Sheet1!F1:F10",
        dataBar: {
          axisColor: "#000000",
          axisFormat: "Automatic",
          barDirection: "Context",
          showDataBarOnly: false,
          lowerBoundRule: { type: "LowestValue" },
          upperBoundRule: { type: "HighestValue" },
          positiveFillColor: "#63C384",
          positiveBorderColor: "#2E8540",
          positiveGradientFill: true,
          negativeFillColor: "#D13438",
          negativeBorderColor: "#A4262C",
          negativeMatchPositiveFillColor: false,
          negativeMatchPositiveBorderColor: false,
        },
      },
      {
        type: "color_scale",
        stopIfTrue: false,
        appliesToAddress: "Sheet1!G1:G10",
        colorScale: {
          minimum: { type: "LowestValue", color: "#F8696B" },
          midpoint: { type: "Percentile", formula: "50", color: "#FFEB84" },
          maximum: { type: "HighestValue", color: "#63BE7B" },
        },
      },
      {
        type: "icon_set",
        stopIfTrue: true,
        appliesToAddress: "Sheet1!H1:H10",
        iconSet: {
          style: "ThreeTrafficLights1",
          reverseIconOrder: false,
          showIconOnly: false,
          criteria: [
            { type: "Percent", operator: "GreaterThanOrEqual", formula: "0" },
            { type: "Percent", operator: "GreaterThanOrEqual", formula: "33" },
            { type: "Percent", operator: "GreaterThanOrEqual", formula: "67" },
          ],
        },
      },
    ]),
  );

  const snapshots = await log.listForCurrentWorkbook(10);
  const inverse = restored.inverseSnapshotId
    ? findSnapshotById(snapshots, restored.inverseSnapshotId)
    : null;

  assert.ok(inverse);
  assert.equal(inverse?.toolName, "restore_snapshot");
  assert.equal(inverse?.snapshotKind, "conditional_format_rules");
  assert.equal(inverse?.restoredFromSnapshotId, appended?.id);
  assert.deepEqual(
    (inverse?.conditionalFormatRules ?? []).map((rule) => ({
      type: rule.type,
      formula: rule.formula,
      fillColor: rule.fillColor,
      appliesToAddress: rule.appliesToAddress,
    })),
    [{
      type: "custom",
      formula: "=A1>0",
      fillColor: "#00FF00",
      appliesToAddress: "Sheet1!A1:A2",
    }],
  );
});
void test("restore round-trips conditional-format rules in target and inverse snapshots", async () => {
  const settingsStore = createInMemorySettingsStore();

  const workbookContext: WorkbookContext = {
    workbookId: "url_sha256:workbook-cf-roundtrip",
    workbookName: "ConditionalRoundtrip.xlsx",
    source: "document.url",
  };

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-cf-roundtrip-${idCounter}`;
  };

  const targetRules: RecoveryConditionalFormatRule[] = [
    {
      type: "custom",
      formula: "=A1>5",
      fillColor: "#F4CCCC",
      appliesToAddress: "Sheet1!A1:A3",
    },
    {
      type: "data_bar",
      appliesToAddress: "Sheet1!B1:B10",
      dataBar: {
        axisColor: "#000000",
        axisFormat: "Automatic",
        barDirection: "Context",
        showDataBarOnly: false,
        lowerBoundRule: { type: "LowestValue" },
        upperBoundRule: { type: "HighestValue" },
      },
    },
    {
      type: "icon_set",
      appliesToAddress: "Sheet1!C1:C6",
      iconSet: {
        style: "ThreeSymbols",
        reverseIconOrder: false,
        showIconOnly: false,
        criteria: [
          { type: "Percent", operator: "GreaterThanOrEqual", formula: "0" },
          { type: "Percent", operator: "GreaterThanOrEqual", formula: "33" },
          { type: "Percent", operator: "GreaterThanOrEqual", formula: "67" },
        ],
      },
    },
  ];

  const currentRules: RecoveryConditionalFormatRule[] = [
    {
      type: "text_comparison",
      textOperator: "Contains",
      text: "priority",
      fillColor: "#FFF2CC",
      appliesToAddress: "Sheet1!D1:D10",
    },
    {
      type: "color_scale",
      appliesToAddress: "Sheet1!E1:E10",
      colorScale: {
        minimum: { type: "LowestValue", color: "#F8696B" },
        midpoint: { type: "Percentile", formula: "50", color: "#FFEB84" },
        maximum: { type: "HighestValue", color: "#63BE7B" },
      },
    },
  ];

  let appliedAddress = "";
  let appliedRules: RecoveryConditionalFormatRule[] = [];

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve(workbookContext),
    now: () => 1700000002500,
    createId,
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
    applyConditionalFormatSnapshot: (address, rules) => {
      appliedAddress = address;
      appliedRules = [...rules];
      return Promise.resolve({
        supported: true,
        rules: currentRules,
      });
    },
  });

  const appended = await log.appendConditionalFormat({
    toolName: "conditional_format",
    toolCallId: "call-cf-roundtrip",
    address: "Sheet1!A1:E10",
    changedCount: 10,
    cellCount: 10,
    conditionalFormatRules: targetRules,
  });

  assert.ok(appended);
  if (!appended) {
    throw new Error("Expected appended conditional format checkpoint.");
  }

  const restored = await log.restore(appended.id);

  assert.equal(restored.address, "Sheet1!A1:E10");
  assert.equal(restored.restoredSnapshotId, appended.id);
  assert.equal(appliedAddress, "Sheet1!A1:E10");
  assert.deepEqual(withoutUndefined(appliedRules), withoutUndefined(targetRules));

  const snapshots = await log.listForCurrentWorkbook(10);
  const inverse = restored.inverseSnapshotId
    ? findSnapshotById(snapshots, restored.inverseSnapshotId)
    : null;

  assert.ok(inverse);
  assert.equal(inverse?.snapshotKind, "conditional_format_rules");
  assert.equal(inverse?.restoredFromSnapshotId, appended.id);
  assert.deepEqual(withoutUndefined(inverse?.conditionalFormatRules), withoutUndefined(currentRules));
});
