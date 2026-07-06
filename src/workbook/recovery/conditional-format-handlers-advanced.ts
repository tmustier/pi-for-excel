/** Advanced conditional-format handlers for data bars, color scales, and icon sets. */

import {
  attachConditionalFormatCaptureContext,
  type ConditionalFormatRuleHandler,
} from "./conditional-format-handlers-basic.js";
import {
  captureColorScaleCriterion,
  captureDataBarRule,
  captureIconCriterion,
  isRecoveryConditionalColorScaleState,
  isRecoveryConditionalDataBarAxisFormat,
  isRecoveryConditionalDataBarDirection,
  isRecoveryConditionalDataBarState,
  isRecoveryConditionalIconSet,
  isRecoveryConditionalIconSetState,
  normalizeOptionalBoolean,
  normalizeOptionalString,
  toColorScaleCriterion,
  toDataBarRule,
  toIconCriterion,
} from "./conditional-format-normalization.js";
import type {
  RecoveryConditionalColorScaleCriterion,
  RecoveryConditionalFormatRule,
} from "./types.js";

type AdvancedConditionalFormatRuleType = Extract<
  RecoveryConditionalFormatRule["type"],
  "data_bar" | "color_scale" | "icon_set"
>;

export const ADVANCED_CONDITIONAL_FORMAT_RULE_HANDLERS = {
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

      const dataBarState: NonNullable<RecoveryConditionalFormatRule["dataBar"]> = {
        axisFormat,
        barDirection,
        showDataBarOnly,
        lowerBoundRule,
        upperBoundRule,
        positiveFillColor,
        positiveGradientFill,
        negativeFillColor,
        negativeMatchPositiveFillColor,
        negativeMatchPositiveBorderColor,
      };

      const axisColor = normalizeOptionalString(dataBar.axisColor);
      if (axisColor !== undefined) {
        dataBarState.axisColor = axisColor;
      }
      const positiveBorderColor = normalizeOptionalString(dataBar.positiveFormat.borderColor);
      if (positiveBorderColor !== undefined) {
        dataBarState.positiveBorderColor = positiveBorderColor;
      }
      const negativeBorderColor = normalizeOptionalString(dataBar.negativeFormat.borderColor);
      if (negativeBorderColor !== undefined) {
        dataBarState.negativeBorderColor = negativeBorderColor;
      }

      return {
        supported: true,
        rule: attachConditionalFormatCaptureContext({
          type: "data_bar",
          dataBar: dataBarState,
        }, captureContext),
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

      const colorScale: NonNullable<RecoveryConditionalFormatRule["colorScale"]> = {
        minimum,
        maximum,
      };
      if (midpoint !== undefined) {
        colorScale.midpoint = midpoint;
      }

      return {
        supported: true,
        rule: attachConditionalFormatCaptureContext({
          type: "color_scale",
          colorScale,
        }, captureContext),
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

      const criteria: NonNullable<RecoveryConditionalFormatRule["iconSet"]>["criteria"] = [];
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
        rule: attachConditionalFormatCaptureContext({
          type: "icon_set",
          iconSet: {
            style,
            reverseIconOrder,
            showIconOnly,
            criteria,
          },
        }, captureContext),
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
} satisfies Record<AdvancedConditionalFormatRuleType, ConditionalFormatRuleHandler>;
