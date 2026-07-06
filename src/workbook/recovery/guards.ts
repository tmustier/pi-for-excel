function isWorkbookRecoveryGuardsPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Runtime guards for persisted recovery payloads. */

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
  RecoveryConditionalFormatRule,
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

function isRecoveryConditionalCellValueOperator(value: DynamicValue): value is RecoveryConditionalCellValueOperator {
  if (typeof value !== "string") return false;

  for (const operator of SUPPORTED_CELL_VALUE_OPERATORS) {
    if (operator === value) {
      return true;
    }
  }

  return false;
}

function isRecoveryConditionalTextOperator(value: DynamicValue): value is RecoveryConditionalTextOperator {
  if (typeof value !== "string") return false;

  for (const operator of SUPPORTED_TEXT_OPERATORS) {
    if (operator === value) {
      return true;
    }
  }

  return false;
}

function isRecoveryConditionalTopBottomCriterionType(value: DynamicValue): value is RecoveryConditionalTopBottomCriterionType {
  if (typeof value !== "string") return false;

  for (const type of SUPPORTED_TOP_BOTTOM_TYPES) {
    if (type === value) {
      return true;
    }
  }

  return false;
}

function isRecoveryConditionalPresetCriterion(value: DynamicValue): value is RecoveryConditionalPresetCriterion {
  if (typeof value !== "string") return false;

  for (const criterion of SUPPORTED_PRESET_CRITERIA) {
    if (criterion === value) {
      return true;
    }
  }

  return false;
}

function isRecoveryConditionalDataBarAxisFormat(value: DynamicValue): value is RecoveryConditionalDataBarAxisFormat {
  if (typeof value !== "string") return false;

  for (const axisFormat of SUPPORTED_DATA_BAR_AXIS_FORMATS) {
    if (axisFormat === value) {
      return true;
    }
  }

  return false;
}

function isRecoveryConditionalDataBarDirection(value: DynamicValue): value is RecoveryConditionalDataBarDirection {
  if (typeof value !== "string") return false;

  for (const direction of SUPPORTED_DATA_BAR_DIRECTIONS) {
    if (direction === value) {
      return true;
    }
  }

  return false;
}

function isRecoveryConditionalDataBarRuleType(value: DynamicValue): value is RecoveryConditionalDataBarRuleType {
  if (typeof value !== "string") return false;

  for (const type of SUPPORTED_DATA_BAR_RULE_TYPES) {
    if (type === value) {
      return true;
    }
  }

  return false;
}

function isRecoveryConditionalColorCriterionType(value: DynamicValue): value is RecoveryConditionalColorCriterionType {
  if (typeof value !== "string") return false;

  for (const type of SUPPORTED_COLOR_CRITERION_TYPES) {
    if (type === value) {
      return true;
    }
  }

  return false;
}

function isRecoveryConditionalIconCriterionType(value: DynamicValue): value is RecoveryConditionalIconCriterionType {
  if (typeof value !== "string") return false;

  for (const type of SUPPORTED_ICON_CRITERION_TYPES) {
    if (type === value) {
      return true;
    }
  }

  return false;
}

function isRecoveryConditionalIconCriterionOperator(value: DynamicValue): value is RecoveryConditionalIconCriterionOperator {
  if (typeof value !== "string") return false;

  for (const operator of SUPPORTED_ICON_CRITERION_OPERATORS) {
    if (operator === value) {
      return true;
    }
  }

  return false;
}

function isRecoveryConditionalIconSet(value: DynamicValue): value is RecoveryConditionalIconSet {
  if (typeof value !== "string") return false;

  for (const style of SUPPORTED_ICON_SETS) {
    if (style === value) {
      return true;
    }
  }

  return false;
}

function isRecoveryConditionalDataBarRule(value: DynamicValue): value is RecoveryConditionalDataBarRule {
  if (!isWorkbookRecoveryGuardsPayloadShape(value)) return false;
  if (!isRecoveryConditionalDataBarRuleType(value.type)) return false;
  if (value.formula !== undefined && typeof value.formula !== "string") return false;
  return true;
}

function isRecoveryConditionalDataBarState(value: DynamicValue): value is RecoveryConditionalDataBarState {
  if (!isWorkbookRecoveryGuardsPayloadShape(value)) return false;
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

function isRecoveryConditionalColorScaleCriterion(value: DynamicValue): value is RecoveryConditionalColorScaleCriterion {
  if (!isWorkbookRecoveryGuardsPayloadShape(value)) return false;
  if (!isRecoveryConditionalColorCriterionType(value.type)) return false;
  if (value.formula !== undefined && typeof value.formula !== "string") return false;
  if (value.color !== undefined && typeof value.color !== "string") return false;
  return true;
}

function isRecoveryConditionalColorScaleState(value: DynamicValue): value is RecoveryConditionalColorScaleState {
  if (!isWorkbookRecoveryGuardsPayloadShape(value)) return false;
  if (!isRecoveryConditionalColorScaleCriterion(value.minimum)) return false;
  if (!isRecoveryConditionalColorScaleCriterion(value.maximum)) return false;
  if (value.midpoint !== undefined && !isRecoveryConditionalColorScaleCriterion(value.midpoint)) return false;
  return true;
}

function isRecoveryConditionalIcon(value: DynamicValue): value is RecoveryConditionalIcon {
  if (!isWorkbookRecoveryGuardsPayloadShape(value)) return false;
  if (!isRecoveryConditionalIconSet(value.set)) return false;
  return typeof value.index === "number" && Number.isFinite(value.index);
}

function isRecoveryConditionalIconCriterion(value: DynamicValue): value is RecoveryConditionalIconCriterion {
  if (!isWorkbookRecoveryGuardsPayloadShape(value)) return false;
  if (!isRecoveryConditionalIconCriterionType(value.type)) return false;
  if (!isRecoveryConditionalIconCriterionOperator(value.operator)) return false;
  if (typeof value.formula !== "string") return false;
  if (value.customIcon !== undefined && !isRecoveryConditionalIcon(value.customIcon)) return false;
  return true;
}

function isRecoveryConditionalIconSetState(value: DynamicValue): value is RecoveryConditionalIconSetState {
  if (!isWorkbookRecoveryGuardsPayloadShape(value)) return false;
  if (!isRecoveryConditionalIconSet(value.style)) return false;
  if (typeof value.reverseIconOrder !== "boolean") return false;
  if (typeof value.showIconOnly !== "boolean") return false;
  if (!Array.isArray(value.criteria) || value.criteria.length === 0) return false;
  if (!value.criteria.every((criterion) => isRecoveryConditionalIconCriterion(criterion))) return false;
  return true;
}

export function isRecoveryConditionalFormatRule(value: DynamicValue): value is RecoveryConditionalFormatRule {
  if (!isWorkbookRecoveryGuardsPayloadShape(value)) return false;

  if (value.stopIfTrue !== undefined && typeof value.stopIfTrue !== "boolean") return false;
  if (value.formula !== undefined && typeof value.formula !== "string") return false;
  if (value.formula1 !== undefined && typeof value.formula1 !== "string") return false;
  if (value.formula2 !== undefined && typeof value.formula2 !== "string") return false;
  if (value.text !== undefined && typeof value.text !== "string") return false;
  if (value.rank !== undefined && (typeof value.rank !== "number" || !Number.isFinite(value.rank))) return false;
  if (value.fillColor !== undefined && typeof value.fillColor !== "string") return false;
  if (value.fontColor !== undefined && typeof value.fontColor !== "string") return false;
  if (value.bold !== undefined && typeof value.bold !== "boolean") return false;
  if (value.italic !== undefined && typeof value.italic !== "boolean") return false;
  if (value.underline !== undefined && typeof value.underline !== "boolean") return false;
  if (value.appliesToAddress !== undefined && typeof value.appliesToAddress !== "string") return false;

  const type = value.type;
  if (type === "custom") {
    return typeof value.formula === "string";
  }

  if (type === "cell_value") {
    return isRecoveryConditionalCellValueOperator(value.operator) && typeof value.formula1 === "string";
  }

  if (type === "text_comparison") {
    return isRecoveryConditionalTextOperator(value.textOperator) && typeof value.text === "string";
  }

  if (type === "top_bottom") {
    return isRecoveryConditionalTopBottomCriterionType(value.topBottomType) && typeof value.rank === "number";
  }

  if (type === "preset_criteria") {
    return isRecoveryConditionalPresetCriterion(value.presetCriterion);
  }

  if (type === "data_bar") {
    return isRecoveryConditionalDataBarState(value.dataBar);
  }

  if (type === "color_scale") {
    return isRecoveryConditionalColorScaleState(value.colorScale);
  }

  if (type === "icon_set") {
    return isRecoveryConditionalIconSetState(value.iconSet);
  }

  return false;
}
