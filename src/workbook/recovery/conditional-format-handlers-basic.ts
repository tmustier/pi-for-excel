function isConditionalFormatHandlersBasicPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Basic conditional-format rule handlers shared by recovery capture/apply flow. */

import {
  isRecoveryConditionalCellValueOperator,
  isRecoveryConditionalPresetCriterion,
  isRecoveryConditionalTextOperator,
  isRecoveryConditionalTopBottomCriterionType,
  normalizeOptionalBoolean,
  normalizeOptionalString,
  normalizeUnderline,
} from "./conditional-format-normalization.js";
import type { RecoveryConditionalFormatRule } from "./types.js";

export interface ConditionalFormatRuleCaptureContext {
  stopIfTrue?: boolean;
  appliesToAddress?: string;
}

interface ConditionalFormatRuleCaptureSuccess {
  supported: true;
  rule: RecoveryConditionalFormatRule;
}

interface ConditionalFormatRuleCaptureFailure {
  supported: false;
  reason: string;
}

export type ConditionalFormatRuleCaptureResult =
  | ConditionalFormatRuleCaptureSuccess
  | ConditionalFormatRuleCaptureFailure;

export interface ConditionalFormatRuleHandler {
  loadForCapture: (conditionalFormat: Excel.ConditionalFormat) => void;
  capture: (
    conditionalFormat: Excel.ConditionalFormat,
    captureContext: ConditionalFormatRuleCaptureContext,
  ) => ConditionalFormatRuleCaptureResult;
  apply: (range: Excel.Range, targetAddress: string, rule: RecoveryConditionalFormatRule) => void;
}

function captureRuleFormatting(format: Excel.ConditionalRangeFormat): {
  fillColor?: string;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
} {
  const formatting: {
    fillColor?: string;
    fontColor?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
  } = {};

  const fillColor = normalizeOptionalString(format.fill.color);
  if (fillColor !== undefined) formatting.fillColor = fillColor;
  const fontColor = normalizeOptionalString(format.font.color);
  if (fontColor !== undefined) formatting.fontColor = fontColor;
  const bold = normalizeOptionalBoolean(format.font.bold);
  if (bold !== undefined) formatting.bold = bold;
  const italic = normalizeOptionalBoolean(format.font.italic);
  if (italic !== undefined) formatting.italic = italic;
  const underline = normalizeUnderline(format.font.underline);
  if (underline !== undefined) formatting.underline = underline;

  return formatting;
}

export function attachConditionalFormatCaptureContext(
  rule: RecoveryConditionalFormatRule,
  captureContext: ConditionalFormatRuleCaptureContext,
): RecoveryConditionalFormatRule {
  if (captureContext.stopIfTrue !== undefined) {
    rule.stopIfTrue = captureContext.stopIfTrue;
  }
  if (captureContext.appliesToAddress !== undefined) {
    rule.appliesToAddress = captureContext.appliesToAddress;
  }
  return rule;
}

function applyRuleFormatting(format: Excel.ConditionalRangeFormat, rule: RecoveryConditionalFormatRule): void {
  if (rule.fillColor !== undefined) {
    format.fill.color = rule.fillColor;
  }

  if (rule.fontColor !== undefined) {
    format.font.color = rule.fontColor;
  }

  if (rule.bold !== undefined) {
    format.font.bold = rule.bold;
  }

  if (rule.italic !== undefined) {
    format.font.italic = rule.italic;
  }

  if (rule.underline !== undefined) {
    format.font.underline = rule.underline ? "Single" : "None";
  }
}

type BasicConditionalFormatRuleType = Extract<
  RecoveryConditionalFormatRule["type"],
  "custom" | "cell_value" | "text_comparison" | "top_bottom" | "preset_criteria"
>;

export const BASIC_CONDITIONAL_FORMAT_RULE_HANDLERS = {
  custom: {
    loadForCapture(conditionalFormat) {
      conditionalFormat.custom.load("rule");
      conditionalFormat.custom.format.fill.load("color");
      conditionalFormat.custom.format.font.load("bold,italic,underline,color");
    },
    capture(conditionalFormat, captureContext) {
      const formula = normalizeOptionalString(conditionalFormat.custom.rule.formula);
      if (formula === undefined) {
        return {
          supported: false,
          reason: "Conditional format checkpoint is invalid: custom rule formula is missing.",
        };
      }

      return {
        supported: true,
        rule: attachConditionalFormatCaptureContext({
          type: "custom",
          formula,
          ...captureRuleFormatting(conditionalFormat.custom.format),
        }, captureContext),
      };
    },
    apply(range, targetAddress, rule) {
      if (typeof rule.formula !== "string") {
        throw new Error("Conditional format checkpoint is invalid: custom rule formula is missing.");
      }

      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      conditionalFormat.custom.rule.formula = rule.formula;
      applyRuleFormatting(conditionalFormat.custom.format, rule);

      if (rule.stopIfTrue !== undefined) {
        conditionalFormat.stopIfTrue = rule.stopIfTrue;
      }

      conditionalFormat.setRanges(targetAddress);
    },
  },
  cell_value: {
    loadForCapture(conditionalFormat) {
      conditionalFormat.cellValue.load("rule");
      conditionalFormat.cellValue.format.fill.load("color");
      conditionalFormat.cellValue.format.font.load("bold,italic,underline,color");
    },
    capture(conditionalFormat, captureContext) {
      const ruleData = conditionalFormat.cellValue.rule;
      const operator = isConditionalFormatHandlersBasicPayloadShape(ruleData) ? ruleData.operator : undefined;

      if (!isRecoveryConditionalCellValueOperator(operator)) {
        return {
          supported: false,
          reason: "Unsupported conditional format rule operator.",
        };
      }

      const formula1 = isConditionalFormatHandlersBasicPayloadShape(ruleData) ? ruleData.formula1 : undefined;
      const formula2 = isConditionalFormatHandlersBasicPayloadShape(ruleData) ? ruleData.formula2 : undefined;

      if (typeof formula1 !== "string") {
        return {
          supported: false,
          reason: "Conditional format rule is missing formula1.",
        };
      }

      return {
        supported: true,
        rule: attachConditionalFormatCaptureContext({
          type: "cell_value",
          operator,
          formula1,
          ...(typeof formula2 === "string" ? { formula2 } : {}),
          ...captureRuleFormatting(conditionalFormat.cellValue.format),
        }, captureContext),
      };
    },
    apply(range, targetAddress, rule) {
      if (!rule.operator || typeof rule.formula1 !== "string") {
        throw new Error("Conditional format checkpoint is invalid: cell value rule is incomplete.");
      }

      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
      const cellValueRule: Excel.ConditionalCellValueRule = {
        operator: rule.operator,
        formula1: rule.formula1,
      };

      if (typeof rule.formula2 === "string") {
        cellValueRule.formula2 = rule.formula2;
      }

      conditionalFormat.cellValue.rule = cellValueRule;
      applyRuleFormatting(conditionalFormat.cellValue.format, rule);

      if (rule.stopIfTrue !== undefined) {
        conditionalFormat.stopIfTrue = rule.stopIfTrue;
      }

      conditionalFormat.setRanges(targetAddress);
    },
  },
  text_comparison: {
    loadForCapture(conditionalFormat) {
      conditionalFormat.textComparison.load("rule");
      conditionalFormat.textComparison.format.fill.load("color");
      conditionalFormat.textComparison.format.font.load("bold,italic,underline,color");
    },
    capture(conditionalFormat, captureContext) {
      const ruleData = conditionalFormat.textComparison.rule;
      const operator = isConditionalFormatHandlersBasicPayloadShape(ruleData) ? ruleData.operator : undefined;

      if (!isRecoveryConditionalTextOperator(operator)) {
        return {
          supported: false,
          reason: "Unsupported conditional format text operator.",
        };
      }

      const text = isConditionalFormatHandlersBasicPayloadShape(ruleData) ? ruleData.text : undefined;
      if (typeof text !== "string") {
        return {
          supported: false,
          reason: "Conditional format text-comparison rule is missing text.",
        };
      }

      return {
        supported: true,
        rule: attachConditionalFormatCaptureContext({
          type: "text_comparison",
          textOperator: operator,
          text,
          ...captureRuleFormatting(conditionalFormat.textComparison.format),
        }, captureContext),
      };
    },
    apply(range, targetAddress, rule) {
      if (!rule.textOperator || typeof rule.text !== "string") {
        throw new Error("Conditional format checkpoint is invalid: text-comparison rule is incomplete.");
      }

      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
      const textRule: Excel.ConditionalTextComparisonRule = {
        operator: rule.textOperator,
        text: rule.text,
      };

      conditionalFormat.textComparison.rule = textRule;
      applyRuleFormatting(conditionalFormat.textComparison.format, rule);

      if (rule.stopIfTrue !== undefined) {
        conditionalFormat.stopIfTrue = rule.stopIfTrue;
      }

      conditionalFormat.setRanges(targetAddress);
    },
  },
  top_bottom: {
    loadForCapture(conditionalFormat) {
      conditionalFormat.topBottom.load("rule");
      conditionalFormat.topBottom.format.fill.load("color");
      conditionalFormat.topBottom.format.font.load("bold,italic,underline,color");
    },
    capture(conditionalFormat, captureContext) {
      const ruleData = conditionalFormat.topBottom.rule;
      const topBottomType = isConditionalFormatHandlersBasicPayloadShape(ruleData) ? ruleData.type : undefined;
      const rank = isConditionalFormatHandlersBasicPayloadShape(ruleData) ? ruleData.rank : undefined;

      if (!isRecoveryConditionalTopBottomCriterionType(topBottomType)) {
        return {
          supported: false,
          reason: "Unsupported conditional format top/bottom criterion type.",
        };
      }

      if (typeof rank !== "number" || !Number.isFinite(rank)) {
        return {
          supported: false,
          reason: "Conditional format top/bottom rule is missing rank.",
        };
      }

      return {
        supported: true,
        rule: attachConditionalFormatCaptureContext({
          type: "top_bottom",
          topBottomType,
          rank,
          ...captureRuleFormatting(conditionalFormat.topBottom.format),
        }, captureContext),
      };
    },
    apply(range, targetAddress, rule) {
      if (!rule.topBottomType || typeof rule.rank !== "number" || !Number.isFinite(rule.rank)) {
        throw new Error("Conditional format checkpoint is invalid: top/bottom rule is incomplete.");
      }

      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
      const topBottomRule: Excel.ConditionalTopBottomRule = {
        type: rule.topBottomType,
        rank: rule.rank,
      };

      conditionalFormat.topBottom.rule = topBottomRule;
      applyRuleFormatting(conditionalFormat.topBottom.format, rule);

      if (rule.stopIfTrue !== undefined) {
        conditionalFormat.stopIfTrue = rule.stopIfTrue;
      }

      conditionalFormat.setRanges(targetAddress);
    },
  },
  preset_criteria: {
    loadForCapture(conditionalFormat) {
      conditionalFormat.preset.load("rule");
      conditionalFormat.preset.format.fill.load("color");
      conditionalFormat.preset.format.font.load("bold,italic,underline,color");
    },
    capture(conditionalFormat, captureContext) {
      const ruleData = conditionalFormat.preset.rule;
      const criterion = isConditionalFormatHandlersBasicPayloadShape(ruleData) ? ruleData.criterion : undefined;

      if (!isRecoveryConditionalPresetCriterion(criterion)) {
        return {
          supported: false,
          reason: "Unsupported conditional format preset criterion.",
        };
      }

      return {
        supported: true,
        rule: attachConditionalFormatCaptureContext({
          type: "preset_criteria",
          presetCriterion: criterion,
          ...captureRuleFormatting(conditionalFormat.preset.format),
        }, captureContext),
      };
    },
    apply(range, targetAddress, rule) {
      if (!rule.presetCriterion) {
        throw new Error("Conditional format checkpoint is invalid: preset-criteria rule is incomplete.");
      }

      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);
      const presetRule: Excel.ConditionalPresetCriteriaRule = {
        criterion: rule.presetCriterion,
      };

      conditionalFormat.preset.rule = presetRule;
      applyRuleFormatting(conditionalFormat.preset.format, rule);

      if (rule.stopIfTrue !== undefined) {
        conditionalFormat.stopIfTrue = rule.stopIfTrue;
      }

      conditionalFormat.setRanges(targetAddress);
    },
  },
} satisfies Record<BasicConditionalFormatRuleType, ConditionalFormatRuleHandler>;
