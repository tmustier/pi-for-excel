function isRecoveryConditionalFormatNormalizationPayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Conditional-format normalization at the Office.js boundary: turns loaded
 * host values into recovery state, and recovery state back into Office.js
 * rule objects.
 */

import type { Static, TSchema } from "typebox";
import { Value } from "typebox/value";

import {
  RecoveryConditionalCellValueOperatorSchema,
  RecoveryConditionalColorCriterionTypeSchema,
  RecoveryConditionalDataBarAxisFormatSchema,
  RecoveryConditionalDataBarDirectionSchema,
  RecoveryConditionalDataBarRuleTypeSchema,
  RecoveryConditionalIconCriterionOperatorSchema,
  RecoveryConditionalIconCriterionTypeSchema,
  RecoveryConditionalIconSetSchema,
  RecoveryConditionalPresetCriterionSchema,
  RecoveryConditionalTextOperatorSchema,
  RecoveryConditionalTopBottomCriterionTypeSchema,
} from "./schemas.js";
import type {
  RecoveryConditionalColorScaleCriterion,
  RecoveryConditionalDataBarRule,
  RecoveryConditionalFormatRuleType,
  RecoveryConditionalIcon,
  RecoveryConditionalIconCriterion,
} from "./types.js";

function enumGuard<S extends TSchema>(schema: S): (value: unknown) => value is Static<S> {
  return (value: unknown): value is Static<S> => Value.Check(schema, value);
}


export const isRecoveryConditionalCellValueOperator = enumGuard(RecoveryConditionalCellValueOperatorSchema);
export const isRecoveryConditionalTextOperator = enumGuard(RecoveryConditionalTextOperatorSchema);
export const isRecoveryConditionalTopBottomCriterionType = enumGuard(RecoveryConditionalTopBottomCriterionTypeSchema);
export const isRecoveryConditionalPresetCriterion = enumGuard(RecoveryConditionalPresetCriterionSchema);
export const isRecoveryConditionalDataBarAxisFormat = enumGuard(RecoveryConditionalDataBarAxisFormatSchema);
export const isRecoveryConditionalDataBarDirection = enumGuard(RecoveryConditionalDataBarDirectionSchema);
export const isRecoveryConditionalDataBarRuleType = enumGuard(RecoveryConditionalDataBarRuleTypeSchema);
export const isRecoveryConditionalColorCriterionType = enumGuard(RecoveryConditionalColorCriterionTypeSchema);
export const isRecoveryConditionalIconCriterionType = enumGuard(RecoveryConditionalIconCriterionTypeSchema);
export const isRecoveryConditionalIconCriterionOperator = enumGuard(RecoveryConditionalIconCriterionOperatorSchema);
export const isRecoveryConditionalIconSet = enumGuard(RecoveryConditionalIconSetSchema);

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

export function normalizeConditionalFormatAddress(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

export function captureDataBarRule(value: unknown): RecoveryConditionalDataBarRule | null {
  if (!isRecoveryConditionalFormatNormalizationPayloadShape(value)) return null;

  const type = value.type;
  if (!isRecoveryConditionalDataBarRuleType(type)) {
    return null;
  }

  const formula = value.formula;
  if (formula !== undefined && typeof formula !== "string") {
    return null;
  }

  const rule: RecoveryConditionalDataBarRule = { type };
  if (typeof formula === "string") {
    rule.formula = formula;
  }

  return rule;
}

export function captureColorScaleCriterion(value: unknown): RecoveryConditionalColorScaleCriterion | null {
  if (!isRecoveryConditionalFormatNormalizationPayloadShape(value)) return null;

  const type = value.type;
  if (!isRecoveryConditionalColorCriterionType(type)) {
    return null;
  }

  const formula = value.formula;
  const color = value.color;

  if (formula !== undefined && typeof formula !== "string") {
    return null;
  }

  if (color !== undefined && typeof color !== "string") {
    return null;
  }

  const criterion: RecoveryConditionalColorScaleCriterion = { type };
  if (typeof formula === "string") {
    criterion.formula = formula;
  }
  if (typeof color === "string") {
    criterion.color = color;
  }

  return criterion;
}

export function captureConditionalIcon(value: unknown): RecoveryConditionalIcon | null {
  if (!isRecoveryConditionalFormatNormalizationPayloadShape(value)) return null;

  if (!isRecoveryConditionalIconSet(value.set)) {
    return null;
  }

  if (typeof value.index !== "number" || !Number.isFinite(value.index)) {
    return null;
  }

  return {
    set: value.set,
    index: value.index,
  };
}

export function captureIconCriterion(value: unknown): RecoveryConditionalIconCriterion | null {
  if (!isRecoveryConditionalFormatNormalizationPayloadShape(value)) return null;

  const type = value.type;
  const operator = value.operator;
  const formula = value.formula;

  if (!isRecoveryConditionalIconCriterionType(type)) {
    return null;
  }

  if (!isRecoveryConditionalIconCriterionOperator(operator)) {
    return null;
  }

  if (typeof formula !== "string") {
    return null;
  }

  let customIcon: RecoveryConditionalIcon | undefined;
  if (value.customIcon !== undefined) {
    const capturedCustomIcon = captureConditionalIcon(value.customIcon);
    if (!capturedCustomIcon) {
      return null;
    }
    customIcon = capturedCustomIcon;
  }

  const criterion: RecoveryConditionalIconCriterion = {
    type,
    operator,
    formula,
  };
  if (customIcon) {
    criterion.customIcon = customIcon;
  }

  return criterion;
}

export function toDataBarRule(rule: RecoveryConditionalDataBarRule): Excel.ConditionalDataBarRule {
  if (typeof rule.formula === "string") {
    return {
      type: rule.type,
      formula: rule.formula,
    };
  }

  return {
    type: rule.type,
  };
}

export function toColorScaleCriterion(
  criterion: RecoveryConditionalColorScaleCriterion,
): Excel.ConditionalColorScaleCriterion {
  if (typeof criterion.formula === "string" && typeof criterion.color === "string") {
    return {
      type: criterion.type,
      formula: criterion.formula,
      color: criterion.color,
    };
  }

  if (typeof criterion.formula === "string") {
    return {
      type: criterion.type,
      formula: criterion.formula,
    };
  }

  if (typeof criterion.color === "string") {
    return {
      type: criterion.type,
      color: criterion.color,
    };
  }

  return {
    type: criterion.type,
  };
}

export function toIconCriterion(criterion: RecoveryConditionalIconCriterion): Excel.ConditionalIconCriterion {
  if (criterion.customIcon) {
    return {
      type: criterion.type,
      operator: criterion.operator,
      formula: criterion.formula,
      customIcon: {
        set: criterion.customIcon.set,
        index: criterion.customIcon.index,
      },
    };
  }

  return {
    type: criterion.type,
    operator: criterion.operator,
    formula: criterion.formula,
  };
}
