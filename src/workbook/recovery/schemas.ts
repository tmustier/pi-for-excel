/**
 * Recovery state schemas.
 *
 * Each persisted recovery state is a TypeBox schema; the exported types are
 * `Static<>` of the schemas, so the code that captures a state and the codec
 * that reads it back share one shape. Cross-field invariants (grid dimensions
 * against declared counts) are refinements on the schema, so a persisted entry
 * that breaks them fails the same check as one with a wrong field type.
 */

import { type Static, Type } from "typebox";

import { StringEnum } from "../../tools/string-enum.js";

// ---------------------------------------------------------------------------
// Conditional formats
// ---------------------------------------------------------------------------

export const RecoveryConditionalCellValueOperatorSchema = StringEnum([
  "Between",
  "NotBetween",
  "EqualTo",
  "NotEqualTo",
  "GreaterThan",
  "LessThan",
  "GreaterThanOrEqual",
  "LessThanOrEqual",
]);

export const RecoveryConditionalTextOperatorSchema = StringEnum([
  "Contains",
  "NotContains",
  "BeginsWith",
  "EndsWith",
]);

export const RecoveryConditionalTopBottomCriterionTypeSchema = StringEnum([
  "TopItems",
  "TopPercent",
  "BottomItems",
  "BottomPercent",
]);

export const RecoveryConditionalPresetCriterionSchema = StringEnum([
  "Blanks",
  "NonBlanks",
  "Errors",
  "NonErrors",
  "Yesterday",
  "Today",
  "Tomorrow",
  "LastSevenDays",
  "LastWeek",
  "ThisWeek",
  "NextWeek",
  "LastMonth",
  "ThisMonth",
  "NextMonth",
  "AboveAverage",
  "BelowAverage",
  "EqualOrAboveAverage",
  "EqualOrBelowAverage",
  "OneStdDevAboveAverage",
  "OneStdDevBelowAverage",
  "TwoStdDevAboveAverage",
  "TwoStdDevBelowAverage",
  "ThreeStdDevAboveAverage",
  "ThreeStdDevBelowAverage",
  "UniqueValues",
  "DuplicateValues",
]);

export const RecoveryConditionalDataBarAxisFormatSchema = StringEnum(["Automatic", "None", "CellMidPoint"]);

export const RecoveryConditionalDataBarDirectionSchema = StringEnum(["Context", "LeftToRight", "RightToLeft"]);

export const RecoveryConditionalDataBarRuleTypeSchema = StringEnum([
  "Automatic",
  "LowestValue",
  "HighestValue",
  "Number",
  "Percent",
  "Formula",
  "Percentile",
]);

export const RecoveryConditionalColorCriterionTypeSchema = StringEnum([
  "LowestValue",
  "HighestValue",
  "Number",
  "Percent",
  "Formula",
  "Percentile",
]);

export const RecoveryConditionalIconCriterionTypeSchema = StringEnum(["Number", "Percent", "Formula", "Percentile"]);

export const RecoveryConditionalIconCriterionOperatorSchema = StringEnum(["GreaterThan", "GreaterThanOrEqual"]);

export const RecoveryConditionalIconSetSchema = StringEnum([
  "ThreeArrows",
  "ThreeArrowsGray",
  "ThreeFlags",
  "ThreeTrafficLights1",
  "ThreeTrafficLights2",
  "ThreeSigns",
  "ThreeSymbols",
  "ThreeSymbols2",
  "FourArrows",
  "FourArrowsGray",
  "FourRedToBlack",
  "FourRating",
  "FourTrafficLights",
  "FiveArrows",
  "FiveArrowsGray",
  "FiveRating",
  "FiveQuarters",
  "ThreeStars",
  "ThreeTriangles",
  "FiveBoxes",
]);

export const RecoveryConditionalDataBarRuleSchema = Type.Object({
  type: RecoveryConditionalDataBarRuleTypeSchema,
  formula: Type.Optional(Type.String()),
});

export const RecoveryConditionalDataBarStateSchema = Type.Object({
  axisColor: Type.Optional(Type.String()),
  axisFormat: RecoveryConditionalDataBarAxisFormatSchema,
  barDirection: RecoveryConditionalDataBarDirectionSchema,
  showDataBarOnly: Type.Boolean(),
  lowerBoundRule: RecoveryConditionalDataBarRuleSchema,
  upperBoundRule: RecoveryConditionalDataBarRuleSchema,
  positiveFillColor: Type.String(),
  positiveBorderColor: Type.Optional(Type.String()),
  positiveGradientFill: Type.Boolean(),
  negativeFillColor: Type.String(),
  negativeBorderColor: Type.Optional(Type.String()),
  negativeMatchPositiveFillColor: Type.Boolean(),
  negativeMatchPositiveBorderColor: Type.Boolean(),
});

export const RecoveryConditionalColorScaleCriterionSchema = Type.Object({
  type: RecoveryConditionalColorCriterionTypeSchema,
  formula: Type.Optional(Type.String()),
  color: Type.Optional(Type.String()),
});

export const RecoveryConditionalColorScaleStateSchema = Type.Object({
  minimum: RecoveryConditionalColorScaleCriterionSchema,
  midpoint: Type.Optional(RecoveryConditionalColorScaleCriterionSchema),
  maximum: RecoveryConditionalColorScaleCriterionSchema,
});

export const RecoveryConditionalIconSchema = Type.Object({
  set: RecoveryConditionalIconSetSchema,
  index: Type.Number(),
});

export const RecoveryConditionalIconCriterionSchema = Type.Object({
  type: RecoveryConditionalIconCriterionTypeSchema,
  operator: RecoveryConditionalIconCriterionOperatorSchema,
  formula: Type.String(),
  customIcon: Type.Optional(RecoveryConditionalIconSchema),
});

export const RecoveryConditionalIconSetStateSchema = Type.Object({
  style: RecoveryConditionalIconSetSchema,
  reverseIconOrder: Type.Boolean(),
  showIconOnly: Type.Boolean(),
  criteria: Type.Array(RecoveryConditionalIconCriterionSchema, { minItems: 1 }),
});

/** Formatting and placement shared by every rule type. */
const conditionalFormatRuleBase = {
  stopIfTrue: Type.Optional(Type.Boolean()),
  fillColor: Type.Optional(Type.String()),
  fontColor: Type.Optional(Type.String()),
  bold: Type.Optional(Type.Boolean()),
  italic: Type.Optional(Type.Boolean()),
  underline: Type.Optional(Type.Boolean()),
  appliesToAddress: Type.Optional(Type.String()),
};

export const RecoveryConditionalFormatRuleSchema = Type.Union([
  Type.Object({ ...conditionalFormatRuleBase, type: Type.Literal("custom"), formula: Type.String() }),
  Type.Object({
    ...conditionalFormatRuleBase,
    type: Type.Literal("cell_value"),
    operator: RecoveryConditionalCellValueOperatorSchema,
    formula1: Type.String(),
    formula2: Type.Optional(Type.String()),
  }),
  Type.Object({
    ...conditionalFormatRuleBase,
    type: Type.Literal("text_comparison"),
    textOperator: RecoveryConditionalTextOperatorSchema,
    text: Type.String(),
  }),
  Type.Object({
    ...conditionalFormatRuleBase,
    type: Type.Literal("top_bottom"),
    topBottomType: RecoveryConditionalTopBottomCriterionTypeSchema,
    rank: Type.Number(),
  }),
  Type.Object({
    ...conditionalFormatRuleBase,
    type: Type.Literal("preset_criteria"),
    presetCriterion: RecoveryConditionalPresetCriterionSchema,
  }),
  Type.Object({ ...conditionalFormatRuleBase, type: Type.Literal("data_bar"), dataBar: RecoveryConditionalDataBarStateSchema }),
  Type.Object({
    ...conditionalFormatRuleBase,
    type: Type.Literal("color_scale"),
    colorScale: RecoveryConditionalColorScaleStateSchema,
  }),
  Type.Object({ ...conditionalFormatRuleBase, type: Type.Literal("icon_set"), iconSet: RecoveryConditionalIconSetStateSchema }),
]);

export const RecoveryConditionalFormatRuleTypeSchema = StringEnum([
  "custom",
  "cell_value",
  "text_comparison",
  "top_bottom",
  "preset_criteria",
  "data_bar",
  "color_scale",
  "icon_set",
]);

// ---------------------------------------------------------------------------
// Comments and charts
// ---------------------------------------------------------------------------

export const RecoveryCommentThreadStateSchema = Type.Object({
  exists: Type.Boolean(),
  content: Type.String(),
  resolved: Type.Boolean(),
  replies: Type.Array(Type.String()),
});

export const RecoveryChartTitleStateSchema = Type.Object({ text: Type.String(), visible: Type.Boolean() });

export const RecoveryChartLegendStateSchema = Type.Object({ position: Type.String(), visible: Type.Boolean() });

export const RecoveryChartPositionStateSchema = Type.Object({
  top: Type.Number(),
  left: Type.Number(),
  width: Type.Number(),
  height: Type.Number(),
});

export const RecoveryChartAbsentStateSchema = Type.Object({
  kind: Type.Literal("chart_absent"),
  sheetName: Type.String(),
  name: Type.String(),
  /**
   * Stable Office.js chart id captured at creation. Restore prefers this over
   * the mutable name so rename/name-reuse cannot delete an unrelated chart.
   */
  chartId: Type.Optional(Type.String()),
});

export const RecoveryChartPresentStateSchema = Type.Object({
  kind: Type.Literal("chart_present"),
  sheetName: Type.String(),
  name: Type.String(),
  chartType: Type.String(),
  title: RecoveryChartTitleStateSchema,
  legend: RecoveryChartLegendStateSchema,
  xAxisTitle: Type.Optional(RecoveryChartTitleStateSchema),
  yAxisTitle: Type.Optional(RecoveryChartTitleStateSchema),
  position: RecoveryChartPositionStateSchema,
});

export const RecoveryChartStateSchema = Type.Union([RecoveryChartAbsentStateSchema, RecoveryChartPresentStateSchema]);

// ---------------------------------------------------------------------------
// Structure
// ---------------------------------------------------------------------------

export const RecoverySheetVisibilitySchema = StringEnum(["Visible", "Hidden", "VeryHidden"]);

/**
 * Cell grids are checked with a row walk rather than a per-cell schema. The
 * cells are opaque host values, and the generic checker costs about 20 ms per
 * 20k cells, which the boot-time load of up to 120 snapshots cannot afford;
 * the row walk costs 0.03 ms.
 */
function isGrid(value: unknown): value is unknown[][] {
  return Array.isArray(value) && value.every((row) => Array.isArray(row));
}

function isStringGrid(value: unknown): value is string[][] {
  return Array.isArray(value) &&
    value.every((row) => Array.isArray(row) && row.every((cell) => typeof cell === "string"));
}

/** Cell grid as Office.js returns it: rows of cell values, no typing per cell. */
export const RecoveryGridSchema = Type.Refine(Type.Unsafe<unknown[][]>({}), isGrid, () => "cell grid must be rows");

const StringGridSchema = Type.Refine(Type.Unsafe<string[][]>({}), isStringGrid, () => "grid must be rows of strings");

const positiveInteger = Type.Integer({ minimum: 1 });

function gridMatches(grid: readonly unknown[][], rowCount: number, columnCount: number): boolean {
  return grid.length === rowCount && grid.every((row) => row.length === columnCount);
}

export const RecoveryStructureValueRangeStateSchema = Type.Refine(
  Type.Object({
    address: Type.String(),
    rowCount: positiveInteger,
    columnCount: positiveInteger,
    values: RecoveryGridSchema,
    formulas: RecoveryGridSchema,
  }),
  (state) =>
    gridMatches(state.values, state.rowCount, state.columnCount) &&
    gridMatches(state.formulas, state.rowCount, state.columnCount),
  () => "values and formulas grids must be rowCount x columnCount",
);

const sheetIdentity = { sheetId: Type.String(), sheetName: Type.String() };
const rowOrColumnRun = { ...sheetIdentity, position: positiveInteger, count: positiveInteger };

export const RecoveryModifyStructureStateSchema = Type.Union([
  Type.Object({ kind: Type.Literal("sheet_name"), sheetId: Type.String(), name: Type.String() }),
  Type.Object({ kind: Type.Literal("sheet_visibility"), sheetId: Type.String(), visibility: RecoverySheetVisibilitySchema }),
  Type.Object({ kind: Type.Literal("sheet_absent"), ...sheetIdentity, allowDataDelete: Type.Optional(Type.Boolean()) }),
  Type.Object({
    kind: Type.Literal("sheet_present"),
    ...sheetIdentity,
    position: Type.Integer({ minimum: 0 }),
    visibility: RecoverySheetVisibilitySchema,
    dataRange: Type.Optional(RecoveryStructureValueRangeStateSchema),
  }),
  Type.Object({ kind: Type.Literal("rows_absent"), ...rowOrColumnRun, allowDataDelete: Type.Optional(Type.Boolean()) }),
  Type.Object({
    kind: Type.Literal("rows_present"),
    ...rowOrColumnRun,
    dataRange: Type.Optional(RecoveryStructureValueRangeStateSchema),
  }),
  Type.Object({ kind: Type.Literal("columns_absent"), ...rowOrColumnRun, allowDataDelete: Type.Optional(Type.Boolean()) }),
  Type.Object({
    kind: Type.Literal("columns_present"),
    ...rowOrColumnRun,
    dataRange: Type.Optional(RecoveryStructureValueRangeStateSchema),
  }),
]);

// ---------------------------------------------------------------------------
// Cell formatting
// ---------------------------------------------------------------------------

export const RecoveryFormatSelectionSchema = Type.Object({
  numberFormat: Type.Optional(Type.Boolean()),
  fillColor: Type.Optional(Type.Boolean()),
  fontColor: Type.Optional(Type.Boolean()),
  bold: Type.Optional(Type.Boolean()),
  italic: Type.Optional(Type.Boolean()),
  underlineStyle: Type.Optional(Type.Boolean()),
  fontName: Type.Optional(Type.Boolean()),
  fontSize: Type.Optional(Type.Boolean()),
  horizontalAlignment: Type.Optional(Type.Boolean()),
  verticalAlignment: Type.Optional(Type.Boolean()),
  wrapText: Type.Optional(Type.Boolean()),
  columnWidth: Type.Optional(Type.Boolean()),
  rowHeight: Type.Optional(Type.Boolean()),
  mergedAreas: Type.Optional(Type.Boolean()),
  borderTop: Type.Optional(Type.Boolean()),
  borderBottom: Type.Optional(Type.Boolean()),
  borderLeft: Type.Optional(Type.Boolean()),
  borderRight: Type.Optional(Type.Boolean()),
  borderInsideHorizontal: Type.Optional(Type.Boolean()),
  borderInsideVertical: Type.Optional(Type.Boolean()),
});

export const RecoveryFormatBorderStateSchema = Type.Object({
  style: Type.String(),
  weight: Type.Optional(Type.String()),
  color: Type.Optional(Type.String()),
});

export const RecoveryFormatAreaStateSchema = Type.Refine(
  Type.Object({
    address: Type.String(),
    rowCount: Type.Number(),
    columnCount: Type.Number(),
    numberFormat: Type.Optional(StringGridSchema),
    fillColor: Type.Optional(Type.String()),
    fontColor: Type.Optional(Type.String()),
    bold: Type.Optional(Type.Boolean()),
    italic: Type.Optional(Type.Boolean()),
    underlineStyle: Type.Optional(Type.String()),
    fontName: Type.Optional(Type.String()),
    fontSize: Type.Optional(Type.Number()),
    horizontalAlignment: Type.Optional(Type.String()),
    verticalAlignment: Type.Optional(Type.String()),
    wrapText: Type.Optional(Type.Boolean()),
    columnWidths: Type.Optional(Type.Array(Type.Number())),
    rowHeights: Type.Optional(Type.Array(Type.Number())),
    mergedAreas: Type.Optional(Type.Array(Type.String())),
    borderTop: Type.Optional(RecoveryFormatBorderStateSchema),
    borderBottom: Type.Optional(RecoveryFormatBorderStateSchema),
    borderLeft: Type.Optional(RecoveryFormatBorderStateSchema),
    borderRight: Type.Optional(RecoveryFormatBorderStateSchema),
    borderInsideHorizontal: Type.Optional(RecoveryFormatBorderStateSchema),
    borderInsideVertical: Type.Optional(RecoveryFormatBorderStateSchema),
  }),
  (area) =>
    (area.numberFormat === undefined || gridMatches(area.numberFormat, area.rowCount, area.columnCount)) &&
    (area.columnWidths === undefined || area.columnWidths.length === area.columnCount) &&
    (area.rowHeights === undefined || area.rowHeights.length === area.rowCount),
  () => "numberFormat, columnWidths and rowHeights must match rowCount x columnCount",
);

export const RecoveryFormatRangeStateSchema = Type.Object({
  selection: RecoveryFormatSelectionSchema,
  areas: Type.Array(RecoveryFormatAreaStateSchema),
  cellCount: Type.Number(),
});

// ---------------------------------------------------------------------------
// Static types
// ---------------------------------------------------------------------------

export type RecoveryConditionalCellValueOperator = Static<typeof RecoveryConditionalCellValueOperatorSchema>;
export type RecoveryConditionalTextOperator = Static<typeof RecoveryConditionalTextOperatorSchema>;
export type RecoveryConditionalTopBottomCriterionType = Static<typeof RecoveryConditionalTopBottomCriterionTypeSchema>;
export type RecoveryConditionalPresetCriterion = Static<typeof RecoveryConditionalPresetCriterionSchema>;
export type RecoveryConditionalDataBarAxisFormat = Static<typeof RecoveryConditionalDataBarAxisFormatSchema>;
export type RecoveryConditionalDataBarDirection = Static<typeof RecoveryConditionalDataBarDirectionSchema>;
export type RecoveryConditionalDataBarRuleType = Static<typeof RecoveryConditionalDataBarRuleTypeSchema>;
export type RecoveryConditionalDataBarRule = Static<typeof RecoveryConditionalDataBarRuleSchema>;
export type RecoveryConditionalDataBarState = Static<typeof RecoveryConditionalDataBarStateSchema>;
export type RecoveryConditionalColorCriterionType = Static<typeof RecoveryConditionalColorCriterionTypeSchema>;
export type RecoveryConditionalColorScaleCriterion = Static<typeof RecoveryConditionalColorScaleCriterionSchema>;
export type RecoveryConditionalColorScaleState = Static<typeof RecoveryConditionalColorScaleStateSchema>;
export type RecoveryConditionalIconCriterionType = Static<typeof RecoveryConditionalIconCriterionTypeSchema>;
export type RecoveryConditionalIconCriterionOperator = Static<typeof RecoveryConditionalIconCriterionOperatorSchema>;
export type RecoveryConditionalIconSet = Static<typeof RecoveryConditionalIconSetSchema>;
export type RecoveryConditionalIcon = Static<typeof RecoveryConditionalIconSchema>;
export type RecoveryConditionalIconCriterion = Static<typeof RecoveryConditionalIconCriterionSchema>;
export type RecoveryConditionalIconSetState = Static<typeof RecoveryConditionalIconSetStateSchema>;
export type RecoveryConditionalFormatRuleType = Static<typeof RecoveryConditionalFormatRuleTypeSchema>;
export type RecoveryConditionalFormatRule = Static<typeof RecoveryConditionalFormatRuleSchema>;
export type RecoveryConditionalFormatRuleOfType<T extends RecoveryConditionalFormatRuleType> = Extract<
  RecoveryConditionalFormatRule,
  { type: T }
>;

export type RecoveryCommentThreadState = Static<typeof RecoveryCommentThreadStateSchema>;
export type RecoveryChartTitleState = Static<typeof RecoveryChartTitleStateSchema>;
export type RecoveryChartLegendState = Static<typeof RecoveryChartLegendStateSchema>;
export type RecoveryChartPositionState = Static<typeof RecoveryChartPositionStateSchema>;
export type RecoveryChartAbsentState = Static<typeof RecoveryChartAbsentStateSchema>;
export type RecoveryChartPresentState = Static<typeof RecoveryChartPresentStateSchema>;
export type RecoveryChartState = Static<typeof RecoveryChartStateSchema>;

export type RecoverySheetVisibility = Static<typeof RecoverySheetVisibilitySchema>;
export type RecoveryStructureValueRangeState = Static<typeof RecoveryStructureValueRangeStateSchema>;
export type RecoveryModifyStructureState = Static<typeof RecoveryModifyStructureStateSchema>;
export type RecoverySheetNameState = Extract<RecoveryModifyStructureState, { kind: "sheet_name" }>;
export type RecoverySheetVisibilityState = Extract<RecoveryModifyStructureState, { kind: "sheet_visibility" }>;
export type RecoverySheetAbsentState = Extract<RecoveryModifyStructureState, { kind: "sheet_absent" }>;
export type RecoverySheetPresentState = Extract<RecoveryModifyStructureState, { kind: "sheet_present" }>;
export type RecoveryRowsAbsentState = Extract<RecoveryModifyStructureState, { kind: "rows_absent" }>;
export type RecoveryRowsPresentState = Extract<RecoveryModifyStructureState, { kind: "rows_present" }>;
export type RecoveryColumnsAbsentState = Extract<RecoveryModifyStructureState, { kind: "columns_absent" }>;
export type RecoveryColumnsPresentState = Extract<RecoveryModifyStructureState, { kind: "columns_present" }>;

export type RecoveryFormatSelection = Static<typeof RecoveryFormatSelectionSchema>;
export type RecoveryFormatBorderState = Static<typeof RecoveryFormatBorderStateSchema>;
export type RecoveryFormatAreaState = Static<typeof RecoveryFormatAreaStateSchema>;
export type RecoveryFormatRangeState = Static<typeof RecoveryFormatRangeStateSchema>;
