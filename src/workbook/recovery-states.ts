import { firstCellAddress } from "./recovery/address.js";
import {
  cloneRecoveryCommentThreadState,
  cloneRecoveryConditionalFormatRules,
  cloneRecoveryFormatRangeState,
  cloneRecoveryModifyStructureState,
} from "./recovery/clone.js";
import { isRecoveryConditionalFormatRule } from "./recovery/guards.js";

export {
  firstCellAddress,
  cloneRecoveryCommentThreadState,
  cloneRecoveryConditionalFormatRules,
  cloneRecoveryFormatRangeState,
  cloneRecoveryModifyStructureState,
  isRecoveryConditionalFormatRule,
};

export { estimateFormatCaptureCellCount } from "./recovery/format-selection.js";
export { applyFormatCellsState, captureFormatCellsState } from "./recovery/format-state.js";
export type { CaptureFormatCellsStateOptions } from "./recovery/format-state.js";
export { applyModifyStructureState, captureModifyStructureState } from "./recovery/structure-state.js";
export { applyConditionalFormatState, captureConditionalFormatState } from "./recovery/conditional-format-state.js";
export { applyCommentThreadState, captureCommentThreadState } from "./recovery/comment-state.js";

export type RecoveryConditionalCellValueOperator =
  | "Between"
  | "NotBetween"
  | "EqualTo"
  | "NotEqualTo"
  | "GreaterThan"
  | "LessThan"
  | "GreaterThanOrEqual"
  | "LessThanOrEqual";

export type RecoveryConditionalTextOperator =
  | "Contains"
  | "NotContains"
  | "BeginsWith"
  | "EndsWith";

export type RecoveryConditionalTopBottomCriterionType =
  | "TopItems"
  | "TopPercent"
  | "BottomItems"
  | "BottomPercent";

export type RecoveryConditionalPresetCriterion =
  | "Blanks"
  | "NonBlanks"
  | "Errors"
  | "NonErrors"
  | "Yesterday"
  | "Today"
  | "Tomorrow"
  | "LastSevenDays"
  | "LastWeek"
  | "ThisWeek"
  | "NextWeek"
  | "LastMonth"
  | "ThisMonth"
  | "NextMonth"
  | "AboveAverage"
  | "BelowAverage"
  | "EqualOrAboveAverage"
  | "EqualOrBelowAverage"
  | "OneStdDevAboveAverage"
  | "OneStdDevBelowAverage"
  | "TwoStdDevAboveAverage"
  | "TwoStdDevBelowAverage"
  | "ThreeStdDevAboveAverage"
  | "ThreeStdDevBelowAverage"
  | "UniqueValues"
  | "DuplicateValues";

export type RecoveryConditionalDataBarAxisFormat = "Automatic" | "None" | "CellMidPoint";

export type RecoveryConditionalDataBarDirection = "Context" | "LeftToRight" | "RightToLeft";

export type RecoveryConditionalDataBarRuleType =
  | "Automatic"
  | "LowestValue"
  | "HighestValue"
  | "Number"
  | "Percent"
  | "Formula"
  | "Percentile";

export interface RecoveryConditionalDataBarRule {
  type: RecoveryConditionalDataBarRuleType;
  formula?: string;
}

export interface RecoveryConditionalDataBarState {
  axisColor?: string;
  axisFormat: RecoveryConditionalDataBarAxisFormat;
  barDirection: RecoveryConditionalDataBarDirection;
  showDataBarOnly: boolean;
  lowerBoundRule: RecoveryConditionalDataBarRule;
  upperBoundRule: RecoveryConditionalDataBarRule;
  positiveFillColor: string;
  positiveBorderColor?: string;
  positiveGradientFill: boolean;
  negativeFillColor: string;
  negativeBorderColor?: string;
  negativeMatchPositiveFillColor: boolean;
  negativeMatchPositiveBorderColor: boolean;
}

export type RecoveryConditionalColorCriterionType =
  | "LowestValue"
  | "HighestValue"
  | "Number"
  | "Percent"
  | "Formula"
  | "Percentile";

export interface RecoveryConditionalColorScaleCriterion {
  type: RecoveryConditionalColorCriterionType;
  formula?: string;
  color?: string;
}

export interface RecoveryConditionalColorScaleState {
  minimum: RecoveryConditionalColorScaleCriterion;
  midpoint?: RecoveryConditionalColorScaleCriterion;
  maximum: RecoveryConditionalColorScaleCriterion;
}

export type RecoveryConditionalIconCriterionType = "Number" | "Percent" | "Formula" | "Percentile";

export type RecoveryConditionalIconCriterionOperator = "GreaterThan" | "GreaterThanOrEqual";

export type RecoveryConditionalIconSet =
  | "ThreeArrows"
  | "ThreeArrowsGray"
  | "ThreeFlags"
  | "ThreeTrafficLights1"
  | "ThreeTrafficLights2"
  | "ThreeSigns"
  | "ThreeSymbols"
  | "ThreeSymbols2"
  | "FourArrows"
  | "FourArrowsGray"
  | "FourRedToBlack"
  | "FourRating"
  | "FourTrafficLights"
  | "FiveArrows"
  | "FiveArrowsGray"
  | "FiveRating"
  | "FiveQuarters"
  | "ThreeStars"
  | "ThreeTriangles"
  | "FiveBoxes";

export interface RecoveryConditionalIcon {
  set: RecoveryConditionalIconSet;
  index: number;
}

export interface RecoveryConditionalIconCriterion {
  type: RecoveryConditionalIconCriterionType;
  operator: RecoveryConditionalIconCriterionOperator;
  formula: string;
  customIcon?: RecoveryConditionalIcon;
}

export interface RecoveryConditionalIconSetState {
  style: RecoveryConditionalIconSet;
  reverseIconOrder: boolean;
  showIconOnly: boolean;
  criteria: RecoveryConditionalIconCriterion[];
}

export type RecoveryConditionalFormatRuleType =
  | "custom"
  | "cell_value"
  | "text_comparison"
  | "top_bottom"
  | "preset_criteria"
  | "data_bar"
  | "color_scale"
  | "icon_set";

export interface RecoveryConditionalFormatRule {
  type: RecoveryConditionalFormatRuleType;
  stopIfTrue?: boolean;
  formula?: string;
  operator?: RecoveryConditionalCellValueOperator;
  formula1?: string;
  formula2?: string;
  textOperator?: RecoveryConditionalTextOperator;
  text?: string;
  topBottomType?: RecoveryConditionalTopBottomCriterionType;
  rank?: number;
  presetCriterion?: RecoveryConditionalPresetCriterion;
  dataBar?: RecoveryConditionalDataBarState;
  colorScale?: RecoveryConditionalColorScaleState;
  iconSet?: RecoveryConditionalIconSetState;
  fillColor?: string;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  appliesToAddress?: string;
}

export interface RecoveryConditionalFormatCaptureResult {
  supported: boolean;
  rules: RecoveryConditionalFormatRule[];
  reason?: string;
}

export interface RecoveryCommentThreadState {
  exists: boolean;
  content: string;
  resolved: boolean;
  replies: string[];
}

export type RecoverySheetVisibility = "Visible" | "Hidden" | "VeryHidden";

export interface RecoverySheetNameState {
  kind: "sheet_name";
  sheetId: string;
  name: string;
}

export interface RecoverySheetVisibilityState {
  kind: "sheet_visibility";
  sheetId: string;
  visibility: RecoverySheetVisibility;
}

export interface RecoveryStructureValueRangeState {
  address: string;
  rowCount: number;
  columnCount: number;
  values: unknown[][];
  formulas: unknown[][];
}

export interface RecoverySheetAbsentState {
  kind: "sheet_absent";
  sheetId: string;
  sheetName: string;
  allowDataDelete?: boolean;
}

export interface RecoverySheetPresentState {
  kind: "sheet_present";
  sheetId: string;
  sheetName: string;
  position: number;
  visibility: RecoverySheetVisibility;
  dataRange?: RecoveryStructureValueRangeState;
}

export interface RecoveryRowsAbsentState {
  kind: "rows_absent";
  sheetId: string;
  sheetName: string;
  position: number;
  count: number;
  allowDataDelete?: boolean;
}

export interface RecoveryRowsPresentState {
  kind: "rows_present";
  sheetId: string;
  sheetName: string;
  position: number;
  count: number;
  dataRange?: RecoveryStructureValueRangeState;
}

export interface RecoveryColumnsAbsentState {
  kind: "columns_absent";
  sheetId: string;
  sheetName: string;
  position: number;
  count: number;
  allowDataDelete?: boolean;
}

export interface RecoveryColumnsPresentState {
  kind: "columns_present";
  sheetId: string;
  sheetName: string;
  position: number;
  count: number;
  dataRange?: RecoveryStructureValueRangeState;
}

export type RecoveryModifyStructureState =
  | RecoverySheetNameState
  | RecoverySheetVisibilityState
  | RecoverySheetAbsentState
  | RecoverySheetPresentState
  | RecoveryRowsAbsentState
  | RecoveryRowsPresentState
  | RecoveryColumnsAbsentState
  | RecoveryColumnsPresentState;

export interface RecoveryFormatSelection {
  numberFormat?: boolean;
  fillColor?: boolean;
  fontColor?: boolean;
  bold?: boolean;
  italic?: boolean;
  underlineStyle?: boolean;
  fontName?: boolean;
  fontSize?: boolean;
  horizontalAlignment?: boolean;
  verticalAlignment?: boolean;
  wrapText?: boolean;
  columnWidth?: boolean;
  rowHeight?: boolean;
  mergedAreas?: boolean;
  borderTop?: boolean;
  borderBottom?: boolean;
  borderLeft?: boolean;
  borderRight?: boolean;
  borderInsideHorizontal?: boolean;
  borderInsideVertical?: boolean;
}

export interface RecoveryFormatBorderState {
  style: string;
  weight?: string;
  color?: string;
}

export interface RecoveryFormatAreaState {
  address: string;
  rowCount: number;
  columnCount: number;
  numberFormat?: string[][];
  fillColor?: string;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  underlineStyle?: string;
  fontName?: string;
  fontSize?: number;
  horizontalAlignment?: string;
  verticalAlignment?: string;
  wrapText?: boolean;
  columnWidths?: number[];
  rowHeights?: number[];
  mergedAreas?: string[];
  borderTop?: RecoveryFormatBorderState;
  borderBottom?: RecoveryFormatBorderState;
  borderLeft?: RecoveryFormatBorderState;
  borderRight?: RecoveryFormatBorderState;
  borderInsideHorizontal?: RecoveryFormatBorderState;
  borderInsideVertical?: RecoveryFormatBorderState;
}

export interface RecoveryFormatRangeState {
  selection: RecoveryFormatSelection;
  areas: RecoveryFormatAreaState[];
  cellCount: number;
}

export interface RecoveryFormatCaptureResult {
  supported: boolean;
  state?: RecoveryFormatRangeState;
  reason?: string;
}

export interface RecoveryFormatAreaShape {
  rowCount: number;
  columnCount: number;
}


