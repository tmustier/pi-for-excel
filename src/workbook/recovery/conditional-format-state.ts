/** Conditional-format capture/apply for workbook recovery snapshots. */

import { excelRun, getRange } from "../../excel/helpers.js";
import { isRecord } from "../../utils/type-guards.js";
import { cloneRecoveryConditionalFormatRules } from "./clone.js";
import type {
  RecoveryConditionalColorScaleCriterion,
  RecoveryConditionalDataBarRule,
  RecoveryConditionalFormatCaptureResult,
  RecoveryConditionalFormatRule,
  RecoveryConditionalFormatRuleType,
  RecoveryConditionalIcon,
  RecoveryConditionalIconCriterion,
} from "./types.js";

import {
  isRecoveryConditionalCellValueOperator,
  isRecoveryConditionalColorCriterionType,
  isRecoveryConditionalColorScaleState,
  isRecoveryConditionalDataBarAxisFormat,
  isRecoveryConditionalDataBarDirection,
  isRecoveryConditionalDataBarRuleType,
  isRecoveryConditionalDataBarState,
  isRecoveryConditionalIconCriterionOperator,
  isRecoveryConditionalIconCriterionType,
  isRecoveryConditionalIconSet,
  isRecoveryConditionalIconSetState,
  isRecoveryConditionalPresetCriterion,
  isRecoveryConditionalTextOperator,
  isRecoveryConditionalTopBottomCriterionType,
  normalizeConditionalFormatType,
  normalizeOptionalBoolean,
  normalizeOptionalString,
  normalizeUnderline,
} from "./conditional-format-normalization.js";

function captureRuleFormatting(format: Excel.ConditionalRangeFormat): {
  fillColor?: string;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
} {
  return {
    fillColor: normalizeOptionalString(format.fill.color),
    fontColor: normalizeOptionalString(format.font.color),
    bold: normalizeOptionalBoolean(format.font.bold),
    italic: normalizeOptionalBoolean(format.font.italic),
    underline: normalizeUnderline(format.font.underline),
  };
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

interface LoadedConditionalFormatEntry {
  conditionalFormat: Excel.ConditionalFormat;
  appliesTo: Excel.RangeAreas;
  normalizedType: RecoveryConditionalFormatRuleType;
}

interface ConditionalFormatRuleCaptureContext {
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

type ConditionalFormatRuleCaptureResult = ConditionalFormatRuleCaptureSuccess | ConditionalFormatRuleCaptureFailure;

interface ConditionalFormatRuleHandler {
  loadForCapture: (conditionalFormat: Excel.ConditionalFormat) => void;
  capture: (
    conditionalFormat: Excel.ConditionalFormat,
    captureContext: ConditionalFormatRuleCaptureContext,
  ) => ConditionalFormatRuleCaptureResult;
  apply: (range: Excel.Range, targetAddress: string, rule: RecoveryConditionalFormatRule) => void;
}

function normalizeConditionalFormatAddress(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function captureDataBarRule(value: unknown): RecoveryConditionalDataBarRule | null {
  if (!isRecord(value)) return null;

  const type = value.type;
  if (!isRecoveryConditionalDataBarRuleType(type)) {
    return null;
  }

  const formula = value.formula;
  if (formula !== undefined && typeof formula !== "string") {
    return null;
  }

  return {
    type,
    formula: typeof formula === "string" ? formula : undefined,
  };
}

function captureColorScaleCriterion(value: unknown): RecoveryConditionalColorScaleCriterion | null {
  if (!isRecord(value)) return null;

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

  return {
    type,
    formula: typeof formula === "string" ? formula : undefined,
    color: typeof color === "string" ? color : undefined,
  };
}

function captureConditionalIcon(value: unknown): RecoveryConditionalIcon | null {
  if (!isRecord(value)) return null;

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

function captureIconCriterion(value: unknown): RecoveryConditionalIconCriterion | null {
  if (!isRecord(value)) return null;

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

  return {
    type,
    operator,
    formula,
    customIcon,
  };
}

function toDataBarRule(rule: RecoveryConditionalDataBarRule): Excel.ConditionalDataBarRule {
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

function toColorScaleCriterion(
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

function toIconCriterion(criterion: RecoveryConditionalIconCriterion): Excel.ConditionalIconCriterion {
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

const CONDITIONAL_FORMAT_RULE_HANDLERS = {
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
        rule: {
          type: "custom",
          stopIfTrue: captureContext.stopIfTrue,
          formula,
          appliesToAddress: captureContext.appliesToAddress,
          ...captureRuleFormatting(conditionalFormat.custom.format),
        },
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
      const operator = isRecord(ruleData) ? ruleData.operator : undefined;

      if (!isRecoveryConditionalCellValueOperator(operator)) {
        return {
          supported: false,
          reason: "Unsupported conditional format rule operator.",
        };
      }

      const formula1 = isRecord(ruleData) ? ruleData.formula1 : undefined;
      const formula2 = isRecord(ruleData) ? ruleData.formula2 : undefined;

      if (typeof formula1 !== "string") {
        return {
          supported: false,
          reason: "Conditional format rule is missing formula1.",
        };
      }

      return {
        supported: true,
        rule: {
          type: "cell_value",
          stopIfTrue: captureContext.stopIfTrue,
          operator,
          formula1,
          formula2: typeof formula2 === "string" ? formula2 : undefined,
          appliesToAddress: captureContext.appliesToAddress,
          ...captureRuleFormatting(conditionalFormat.cellValue.format),
        },
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
      const operator = isRecord(ruleData) ? ruleData.operator : undefined;

      if (!isRecoveryConditionalTextOperator(operator)) {
        return {
          supported: false,
          reason: "Unsupported conditional format text operator.",
        };
      }

      const text = isRecord(ruleData) ? ruleData.text : undefined;
      if (typeof text !== "string") {
        return {
          supported: false,
          reason: "Conditional format text-comparison rule is missing text.",
        };
      }

      return {
        supported: true,
        rule: {
          type: "text_comparison",
          stopIfTrue: captureContext.stopIfTrue,
          textOperator: operator,
          text,
          appliesToAddress: captureContext.appliesToAddress,
          ...captureRuleFormatting(conditionalFormat.textComparison.format),
        },
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
      const topBottomType = isRecord(ruleData) ? ruleData.type : undefined;
      const rank = isRecord(ruleData) ? ruleData.rank : undefined;

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
        rule: {
          type: "top_bottom",
          stopIfTrue: captureContext.stopIfTrue,
          topBottomType,
          rank,
          appliesToAddress: captureContext.appliesToAddress,
          ...captureRuleFormatting(conditionalFormat.topBottom.format),
        },
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
      const criterion = isRecord(ruleData) ? ruleData.criterion : undefined;

      if (!isRecoveryConditionalPresetCriterion(criterion)) {
        return {
          supported: false,
          reason: "Unsupported conditional format preset criterion.",
        };
      }

      return {
        supported: true,
        rule: {
          type: "preset_criteria",
          stopIfTrue: captureContext.stopIfTrue,
          presetCriterion: criterion,
          appliesToAddress: captureContext.appliesToAddress,
          ...captureRuleFormatting(conditionalFormat.preset.format),
        },
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
  data_bar: {
    loadForCapture(conditionalFormat) {
      conditionalFormat.dataBar.load("axisColor,axisFormat,barDirection,lowerBoundRule,showDataBarOnly,upperBoundRule");
      conditionalFormat.dataBar.positiveFormat.load("fillColor,borderColor,gradientFill");
      conditionalFormat.dataBar.negativeFormat.load("fillColor,borderColor,matchPositiveFillColor,matchPositiveBorderColor");
    },
    capture(conditionalFormat, captureContext) {
      const dataBar = conditionalFormat.dataBar;
      const axisFormat = dataBar.axisFormat;
      const barDirection = dataBar.barDirection;

      if (!isRecoveryConditionalDataBarAxisFormat(axisFormat)) {
        return {
          supported: false,
          reason: "Unsupported conditional format data-bar axis format.",
        };
      }

      if (!isRecoveryConditionalDataBarDirection(barDirection)) {
        return {
          supported: false,
          reason: "Unsupported conditional format data-bar direction.",
        };
      }

      const showDataBarOnly = normalizeOptionalBoolean(dataBar.showDataBarOnly);
      if (showDataBarOnly === undefined) {
        return {
          supported: false,
          reason: "Conditional format data-bar showDataBarOnly is unsupported.",
        };
      }

      const lowerBoundRule = captureDataBarRule(dataBar.lowerBoundRule);
      if (!lowerBoundRule) {
        return {
          supported: false,
          reason: "Unsupported conditional format data-bar lower bound rule.",
        };
      }

      const upperBoundRule = captureDataBarRule(dataBar.upperBoundRule);
      if (!upperBoundRule) {
        return {
          supported: false,
          reason: "Unsupported conditional format data-bar upper bound rule.",
        };
      }

      const positiveFillColor = normalizeOptionalString(dataBar.positiveFormat.fillColor);
      if (positiveFillColor === undefined) {
        return {
          supported: false,
          reason: "Conditional format data-bar positive fill color is unavailable.",
        };
      }

      const positiveGradientFill = normalizeOptionalBoolean(dataBar.positiveFormat.gradientFill);
      if (positiveGradientFill === undefined) {
        return {
          supported: false,
          reason: "Conditional format data-bar positive gradient setting is unavailable.",
        };
      }

      const negativeFillColor = normalizeOptionalString(dataBar.negativeFormat.fillColor);
      if (negativeFillColor === undefined) {
        return {
          supported: false,
          reason: "Conditional format data-bar negative fill color is unavailable.",
        };
      }

      const negativeMatchPositiveFillColor = normalizeOptionalBoolean(dataBar.negativeFormat.matchPositiveFillColor);
      if (negativeMatchPositiveFillColor === undefined) {
        return {
          supported: false,
          reason: "Conditional format data-bar negative fill matching setting is unavailable.",
        };
      }

      const negativeMatchPositiveBorderColor = normalizeOptionalBoolean(dataBar.negativeFormat.matchPositiveBorderColor);
      if (negativeMatchPositiveBorderColor === undefined) {
        return {
          supported: false,
          reason: "Conditional format data-bar negative border matching setting is unavailable.",
        };
      }

      return {
        supported: true,
        rule: {
          type: "data_bar",
          stopIfTrue: captureContext.stopIfTrue,
          appliesToAddress: captureContext.appliesToAddress,
          dataBar: {
            axisColor: normalizeOptionalString(dataBar.axisColor),
            axisFormat,
            barDirection,
            showDataBarOnly,
            lowerBoundRule,
            upperBoundRule,
            positiveFillColor,
            positiveBorderColor: normalizeOptionalString(dataBar.positiveFormat.borderColor),
            positiveGradientFill,
            negativeFillColor,
            negativeBorderColor: normalizeOptionalString(dataBar.negativeFormat.borderColor),
            negativeMatchPositiveFillColor,
            negativeMatchPositiveBorderColor,
          },
        },
      };
    },
    apply(range, targetAddress, rule) {
      if (!isRecoveryConditionalDataBarState(rule.dataBar)) {
        throw new Error("Conditional format checkpoint is invalid: data-bar rule is incomplete.");
      }

      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
      const state = rule.dataBar;
      const dataBar = conditionalFormat.dataBar;

      if (typeof state.axisColor === "string") {
        dataBar.axisColor = state.axisColor;
      }

      dataBar.axisFormat = state.axisFormat;
      dataBar.barDirection = state.barDirection;
      dataBar.showDataBarOnly = state.showDataBarOnly;
      dataBar.lowerBoundRule = toDataBarRule(state.lowerBoundRule);
      dataBar.upperBoundRule = toDataBarRule(state.upperBoundRule);

      dataBar.positiveFormat.fillColor = state.positiveFillColor;
      if (typeof state.positiveBorderColor === "string") {
        dataBar.positiveFormat.borderColor = state.positiveBorderColor;
      }
      dataBar.positiveFormat.gradientFill = state.positiveGradientFill;

      dataBar.negativeFormat.fillColor = state.negativeFillColor;
      if (typeof state.negativeBorderColor === "string") {
        dataBar.negativeFormat.borderColor = state.negativeBorderColor;
      }
      dataBar.negativeFormat.matchPositiveFillColor = state.negativeMatchPositiveFillColor;
      dataBar.negativeFormat.matchPositiveBorderColor = state.negativeMatchPositiveBorderColor;

      if (rule.stopIfTrue !== undefined) {
        conditionalFormat.stopIfTrue = rule.stopIfTrue;
      }

      conditionalFormat.setRanges(targetAddress);
    },
  },
  color_scale: {
    loadForCapture(conditionalFormat) {
      conditionalFormat.colorScale.load("criteria");
    },
    capture(conditionalFormat, captureContext) {
      const criteria = conditionalFormat.colorScale.criteria;

      const minimum = captureColorScaleCriterion(criteria.minimum);
      if (!minimum) {
        return {
          supported: false,
          reason: "Unsupported conditional format color-scale minimum criterion.",
        };
      }

      const maximum = captureColorScaleCriterion(criteria.maximum);
      if (!maximum) {
        return {
          supported: false,
          reason: "Unsupported conditional format color-scale maximum criterion.",
        };
      }

      const midpointRaw = criteria.midpoint;
      let midpoint: RecoveryConditionalColorScaleCriterion | undefined;
      if (midpointRaw !== undefined) {
        const capturedMidpoint = captureColorScaleCriterion(midpointRaw);
        if (!capturedMidpoint) {
          return {
            supported: false,
            reason: "Unsupported conditional format color-scale midpoint criterion.",
          };
        }

        midpoint = capturedMidpoint;
      }

      return {
        supported: true,
        rule: {
          type: "color_scale",
          stopIfTrue: captureContext.stopIfTrue,
          appliesToAddress: captureContext.appliesToAddress,
          colorScale: {
            minimum,
            midpoint,
            maximum,
          },
        },
      };
    },
    apply(range, targetAddress, rule) {
      if (!isRecoveryConditionalColorScaleState(rule.colorScale)) {
        throw new Error("Conditional format checkpoint is invalid: color-scale rule is incomplete.");
      }

      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
      const state = rule.colorScale;
      const criteria: Excel.ConditionalColorScaleCriteria = {
        minimum: toColorScaleCriterion(state.minimum),
        maximum: toColorScaleCriterion(state.maximum),
      };

      if (state.midpoint) {
        criteria.midpoint = toColorScaleCriterion(state.midpoint);
      }

      conditionalFormat.colorScale.criteria = criteria;

      if (rule.stopIfTrue !== undefined) {
        conditionalFormat.stopIfTrue = rule.stopIfTrue;
      }

      conditionalFormat.setRanges(targetAddress);
    },
  },
  icon_set: {
    loadForCapture(conditionalFormat) {
      conditionalFormat.iconSet.load("style,reverseIconOrder,showIconOnly,criteria");
    },
    capture(conditionalFormat, captureContext) {
      const iconSet = conditionalFormat.iconSet;
      const style = iconSet.style;
      if (!isRecoveryConditionalIconSet(style)) {
        return {
          supported: false,
          reason: "Unsupported conditional format icon-set style.",
        };
      }

      const reverseIconOrder = normalizeOptionalBoolean(iconSet.reverseIconOrder);
      if (reverseIconOrder === undefined) {
        return {
          supported: false,
          reason: "Conditional format icon-set reverseIconOrder is unavailable.",
        };
      }

      const showIconOnly = normalizeOptionalBoolean(iconSet.showIconOnly);
      if (showIconOnly === undefined) {
        return {
          supported: false,
          reason: "Conditional format icon-set showIconOnly is unavailable.",
        };
      }

      const criteriaRaw = iconSet.criteria;
      if (!Array.isArray(criteriaRaw) || criteriaRaw.length === 0) {
        return {
          supported: false,
          reason: "Conditional format icon-set criteria are unavailable.",
        };
      }

      const criteria: RecoveryConditionalIconCriterion[] = [];
      for (const criterion of criteriaRaw) {
        const captured = captureIconCriterion(criterion);
        if (!captured) {
          return {
            supported: false,
            reason: "Unsupported conditional format icon-set criterion.",
          };
        }

        criteria.push(captured);
      }

      return {
        supported: true,
        rule: {
          type: "icon_set",
          stopIfTrue: captureContext.stopIfTrue,
          appliesToAddress: captureContext.appliesToAddress,
          iconSet: {
            style,
            reverseIconOrder,
            showIconOnly,
            criteria,
          },
        },
      };
    },
    apply(range, targetAddress, rule) {
      if (!isRecoveryConditionalIconSetState(rule.iconSet)) {
        throw new Error("Conditional format checkpoint is invalid: icon-set rule is incomplete.");
      }

      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
      const state = rule.iconSet;
      conditionalFormat.iconSet.style = state.style;
      conditionalFormat.iconSet.reverseIconOrder = state.reverseIconOrder;
      conditionalFormat.iconSet.showIconOnly = state.showIconOnly;
      conditionalFormat.iconSet.criteria = state.criteria.map((criterion) => toIconCriterion(criterion));

      if (rule.stopIfTrue !== undefined) {
        conditionalFormat.stopIfTrue = rule.stopIfTrue;
      }

      conditionalFormat.setRanges(targetAddress);
    },
  },
} satisfies Record<RecoveryConditionalFormatRuleType, ConditionalFormatRuleHandler>;

async function captureConditionalFormatRulesInRange(
  context: Excel.RequestContext,
  range: Excel.Range,
): Promise<RecoveryConditionalFormatCaptureResult> {
  const collection = range.conditionalFormats;
  collection.load("items/type,items/stopIfTrue");
  await context.sync();

  const entries: LoadedConditionalFormatEntry[] = [];

  for (const conditionalFormat of collection.items) {
    const normalizedType = normalizeConditionalFormatType(conditionalFormat.type);
    if (!normalizedType) {
      return {
        supported: false,
        rules: [],
        reason: `Unsupported conditional format type: ${String(conditionalFormat.type)}`,
      };
    }

    const appliesTo = conditionalFormat.getRanges();
    appliesTo.load("address");

    const handler = CONDITIONAL_FORMAT_RULE_HANDLERS[normalizedType];
    handler.loadForCapture(conditionalFormat);

    entries.push({ conditionalFormat, appliesTo, normalizedType });
  }

  await context.sync();

  const rules: RecoveryConditionalFormatRule[] = [];

  for (const entry of entries) {
    const handler = CONDITIONAL_FORMAT_RULE_HANDLERS[entry.normalizedType];
    const captureResult = handler.capture(entry.conditionalFormat, {
      stopIfTrue: normalizeOptionalBoolean(entry.conditionalFormat.stopIfTrue),
      appliesToAddress: normalizeConditionalFormatAddress(entry.appliesTo.address),
    });

    if (!captureResult.supported) {
      return {
        supported: false,
        rules: [],
        reason: captureResult.reason,
      };
    }

    rules.push(captureResult.rule);
  }

  return {
    supported: true,
    rules,
  };
}

function resolveConditionalFormatTargetAddress(
  fallbackAddress: string,
  rule: RecoveryConditionalFormatRule,
): string {
  return normalizeConditionalFormatAddress(rule.appliesToAddress) ?? fallbackAddress;
}

function applyConditionalFormatRule(
  range: Excel.Range,
  fallbackAddress: string,
  rule: RecoveryConditionalFormatRule,
): void {
  const targetAddress = resolveConditionalFormatTargetAddress(fallbackAddress, rule);
  const handler = CONDITIONAL_FORMAT_RULE_HANDLERS[rule.type];
  handler.apply(range, targetAddress, rule);
}

export async function captureConditionalFormatState(address: string): Promise<RecoveryConditionalFormatCaptureResult> {
  return excelRun<RecoveryConditionalFormatCaptureResult>(async (context) => {
    const { range } = getRange(context, address);
    return captureConditionalFormatRulesInRange(context, range);
  });
}

export async function applyConditionalFormatState(
  address: string,
  targetRules: readonly RecoveryConditionalFormatRule[],
): Promise<RecoveryConditionalFormatCaptureResult> {
  return excelRun<RecoveryConditionalFormatCaptureResult>(async (context) => {
    const { range } = getRange(context, address);
    const currentState = await captureConditionalFormatRulesInRange(context, range);

    if (!currentState.supported) {
      throw new Error(currentState.reason ?? "Conditional format checkpoint cannot be restored safely.");
    }

    range.conditionalFormats.clearAll();

    for (const rule of targetRules) {
      applyConditionalFormatRule(range, address, rule);
    }

    await context.sync();

    return {
      supported: true,
      rules: cloneRecoveryConditionalFormatRules(currentState.rules),
    };
  });
}

