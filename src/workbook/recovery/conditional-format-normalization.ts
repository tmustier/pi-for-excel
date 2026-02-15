/** Shared conditional-format normalization and schema helpers. */

import { isRecord } from "../../utils/type-guards.js";
import type {
  RecoveryConditionalCellValueOperator,
  RecoveryConditionalColorCriterionType,
  RecoveryConditionalColorScaleCriterion,
  RecoveryConditionalColorScaleState,
  RecoveryConditionalDataBarAxisFormat,
  RecoveryConditionalDataBarDirection,
  RecoveryConditionalDataBarRule,
  RecoveryConditionalDataBarRuleType,
  RecoveryConditionalDataBarState,
  RecoveryConditionalFormatRuleType,
  RecoveryConditionalIcon,
  RecoveryConditionalIconCriterion,
  RecoveryConditionalIconCriterionOperator,
  RecoveryConditionalIconCriterionType,
  RecoveryConditionalIconSet,
  RecoveryConditionalIconSetState,
  RecoveryConditionalPresetCriterion,
  RecoveryConditionalTextOperator,
  RecoveryConditionalTopBottomCriterionType,
} from "./types.js";

const SUPPORTED_CELL_VALUE_OPERATORS: readonly RecoveryConditionalCellValueOperator[] = [
  "Between",
  "NotBetween",
  "EqualTo",
  "NotEqualTo",
  "GreaterThan",
  "LessThan",
  "GreaterThanOrEqual",
  "LessThanOrEqual",
];

const SUPPORTED_TEXT_OPERATORS: readonly RecoveryConditionalTextOperator[] = [
  "Contains",
  "NotContains",
  "BeginsWith",
  "EndsWith",
];

const SUPPORTED_TOP_BOTTOM_TYPES: readonly RecoveryConditionalTopBottomCriterionType[] = [
  "TopItems",
  "TopPercent",
  "BottomItems",
  "BottomPercent",
];

const SUPPORTED_PRESET_CRITERIA: readonly RecoveryConditionalPresetCriterion[] = [
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
];

const SUPPORTED_DATA_BAR_AXIS_FORMATS: readonly RecoveryConditionalDataBarAxisFormat[] = [
  "Automatic",
  "None",
  "CellMidPoint",
];

const SUPPORTED_DATA_BAR_DIRECTIONS: readonly RecoveryConditionalDataBarDirection[] = [
  "Context",
  "LeftToRight",
  "RightToLeft",
];

const SUPPORTED_DATA_BAR_RULE_TYPES: readonly RecoveryConditionalDataBarRuleType[] = [
  "Automatic",
  "LowestValue",
  "HighestValue",
  "Number",
  "Percent",
  "Formula",
  "Percentile",
];

const SUPPORTED_COLOR_CRITERION_TYPES: readonly RecoveryConditionalColorCriterionType[] = [
  "LowestValue",
  "HighestValue",
  "Number",
  "Percent",
  "Formula",
  "Percentile",
];

const SUPPORTED_ICON_CRITERION_TYPES: readonly RecoveryConditionalIconCriterionType[] = [
  "Number",
  "Percent",
  "Formula",
  "Percentile",
];

const SUPPORTED_ICON_CRITERION_OPERATORS: readonly RecoveryConditionalIconCriterionOperator[] = [
  "GreaterThan",
  "GreaterThanOrEqual",
];

const SUPPORTED_ICON_SETS: readonly RecoveryConditionalIconSet[] = [
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
];

export function isRecoveryConditionalCellValueOperator(value: unknown): value is RecoveryConditionalCellValueOperator {
  if (typeof value !== "string") return false;

  for (const operator of SUPPORTED_CELL_VALUE_OPERATORS) {
    if (operator === value) {
      return true;
    }
  }

  return false;
}

export function isRecoveryConditionalTextOperator(value: unknown): value is RecoveryConditionalTextOperator {
  if (typeof value !== "string") return false;

  for (const operator of SUPPORTED_TEXT_OPERATORS) {
    if (operator === value) {
      return true;
    }
  }

  return false;
}

export function isRecoveryConditionalTopBottomCriterionType(value: unknown): value is RecoveryConditionalTopBottomCriterionType {
  if (typeof value !== "string") return false;

  for (const type of SUPPORTED_TOP_BOTTOM_TYPES) {
    if (type === value) {
      return true;
    }
  }

  return false;
}

export function isRecoveryConditionalPresetCriterion(value: unknown): value is RecoveryConditionalPresetCriterion {
  if (typeof value !== "string") return false;

  for (const criterion of SUPPORTED_PRESET_CRITERIA) {
    if (criterion === value) {
      return true;
    }
  }

  return false;
}

export function isRecoveryConditionalDataBarAxisFormat(value: unknown): value is RecoveryConditionalDataBarAxisFormat {
  if (typeof value !== "string") return false;

  for (const axisFormat of SUPPORTED_DATA_BAR_AXIS_FORMATS) {
    if (axisFormat === value) {
      return true;
    }
  }

  return false;
}

export function isRecoveryConditionalDataBarDirection(value: unknown): value is RecoveryConditionalDataBarDirection {
  if (typeof value !== "string") return false;

  for (const direction of SUPPORTED_DATA_BAR_DIRECTIONS) {
    if (direction === value) {
      return true;
    }
  }

  return false;
}

export function isRecoveryConditionalDataBarRuleType(value: unknown): value is RecoveryConditionalDataBarRuleType {
  if (typeof value !== "string") return false;

  for (const type of SUPPORTED_DATA_BAR_RULE_TYPES) {
    if (type === value) {
      return true;
    }
  }

  return false;
}

export function isRecoveryConditionalColorCriterionType(value: unknown): value is RecoveryConditionalColorCriterionType {
  if (typeof value !== "string") return false;

  for (const type of SUPPORTED_COLOR_CRITERION_TYPES) {
    if (type === value) {
      return true;
    }
  }

  return false;
}

export function isRecoveryConditionalIconCriterionType(value: unknown): value is RecoveryConditionalIconCriterionType {
  if (typeof value !== "string") return false;

  for (const type of SUPPORTED_ICON_CRITERION_TYPES) {
    if (type === value) {
      return true;
    }
  }

  return false;
}

export function isRecoveryConditionalIconCriterionOperator(value: unknown): value is RecoveryConditionalIconCriterionOperator {
  if (typeof value !== "string") return false;

  for (const operator of SUPPORTED_ICON_CRITERION_OPERATORS) {
    if (operator === value) {
      return true;
    }
  }

  return false;
}

export function isRecoveryConditionalIconSet(value: unknown): value is RecoveryConditionalIconSet {
  if (typeof value !== "string") return false;

  for (const style of SUPPORTED_ICON_SETS) {
    if (style === value) {
      return true;
    }
  }

  return false;
}

export function normalizeConditionalFormatType(type: unknown): RecoveryConditionalFormatRuleType | null {
  if (type === "Custom" || type === "custom") {
    return "custom";
  }

  if (type === "CellValue" || type === "cellValue") {
    return "cell_value";
  }

  if (type === "ContainsText" || type === "containsText") {
    return "text_comparison";
  }

  if (type === "TopBottom" || type === "topBottom") {
    return "top_bottom";
  }

  if (type === "PresetCriteria" || type === "presetCriteria") {
    return "preset_criteria";
  }

  if (type === "DataBar" || type === "dataBar") {
    return "data_bar";
  }

  if (type === "ColorScale" || type === "colorScale") {
    return "color_scale";
  }

  if (type === "IconSet" || type === "iconSet") {
    return "icon_set";
  }

  return null;
}

export function normalizeOptionalString(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

export function normalizeOptionalBoolean(value: unknown): boolean | undefined {
  return typeof value === "boolean" ? value : undefined;
}

export function normalizeUnderline(value: unknown): boolean | undefined {
  if (typeof value === "boolean") return value;

  if (typeof value === "string") {
    return value !== "None";
  }

  return undefined;
}

export function isRecoveryConditionalDataBarRule(value: unknown): value is RecoveryConditionalDataBarRule {
  if (!isRecord(value)) return false;
  if (!isRecoveryConditionalDataBarRuleType(value.type)) return false;
  if (value.formula !== undefined && typeof value.formula !== "string") return false;
  return true;
}

export function isRecoveryConditionalDataBarState(value: unknown): value is RecoveryConditionalDataBarState {
  if (!isRecord(value)) return false;
  if (!isRecoveryConditionalDataBarAxisFormat(value.axisFormat)) return false;
  if (!isRecoveryConditionalDataBarDirection(value.barDirection)) return false;
  if (typeof value.showDataBarOnly !== "boolean") return false;
  if (!isRecoveryConditionalDataBarRule(value.lowerBoundRule)) return false;
  if (!isRecoveryConditionalDataBarRule(value.upperBoundRule)) return false;
  if (typeof value.positiveFillColor !== "string") return false;
  if (value.positiveBorderColor !== undefined && typeof value.positiveBorderColor !== "string") return false;
  if (typeof value.positiveGradientFill !== "boolean") return false;
  if (typeof value.negativeFillColor !== "string") return false;
  if (value.negativeBorderColor !== undefined && typeof value.negativeBorderColor !== "string") return false;
  if (typeof value.negativeMatchPositiveFillColor !== "boolean") return false;
  if (typeof value.negativeMatchPositiveBorderColor !== "boolean") return false;
  if (value.axisColor !== undefined && typeof value.axisColor !== "string") return false;
  return true;
}

export function isRecoveryConditionalColorScaleCriterion(value: unknown): value is RecoveryConditionalColorScaleCriterion {
  if (!isRecord(value)) return false;
  if (!isRecoveryConditionalColorCriterionType(value.type)) return false;
  if (value.formula !== undefined && typeof value.formula !== "string") return false;
  if (value.color !== undefined && typeof value.color !== "string") return false;
  return true;
}

export function isRecoveryConditionalColorScaleState(value: unknown): value is RecoveryConditionalColorScaleState {
  if (!isRecord(value)) return false;
  if (!isRecoveryConditionalColorScaleCriterion(value.minimum)) return false;
  if (!isRecoveryConditionalColorScaleCriterion(value.maximum)) return false;
  if (value.midpoint !== undefined && !isRecoveryConditionalColorScaleCriterion(value.midpoint)) return false;
  return true;
}

export function isRecoveryConditionalIcon(value: unknown): value is RecoveryConditionalIcon {
  if (!isRecord(value)) return false;
  if (!isRecoveryConditionalIconSet(value.set)) return false;
  return typeof value.index === "number" && Number.isFinite(value.index);
}

export function isRecoveryConditionalIconCriterion(value: unknown): value is RecoveryConditionalIconCriterion {
  if (!isRecord(value)) return false;
  if (!isRecoveryConditionalIconCriterionType(value.type)) return false;
  if (!isRecoveryConditionalIconCriterionOperator(value.operator)) return false;
  if (typeof value.formula !== "string") return false;
  if (value.customIcon !== undefined && !isRecoveryConditionalIcon(value.customIcon)) return false;
  return true;
}

export function isRecoveryConditionalIconSetState(value: unknown): value is RecoveryConditionalIconSetState {
  if (!isRecord(value)) return false;
  if (!isRecoveryConditionalIconSet(value.style)) return false;
  if (typeof value.reverseIconOrder !== "boolean") return false;
  if (typeof value.showIconOnly !== "boolean") return false;
  if (!Array.isArray(value.criteria) || value.criteria.length === 0) return false;
  if (!value.criteria.every((criterion) => isRecoveryConditionalIconCriterion(criterion))) return false;
  return true;
}
