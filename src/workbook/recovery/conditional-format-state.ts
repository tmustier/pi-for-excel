/** Conditional-format capture/apply for workbook recovery snapshots. */

import { excelRun, getRange } from "../../excel/helpers.js";
import { cloneRecoveryConditionalFormatRules } from "./clone.js";
import {
  BASIC_CONDITIONAL_FORMAT_RULE_HANDLERS,
  type ConditionalFormatRuleCaptureContext,
  type ConditionalFormatRuleHandler,
} from "./conditional-format-handlers-basic.js";
import { ADVANCED_CONDITIONAL_FORMAT_RULE_HANDLERS } from "./conditional-format-handlers-advanced.js";
import type {
  RecoveryConditionalFormatCaptureResult,
  RecoveryConditionalFormatRule,
  RecoveryConditionalFormatRuleType,
} from "./types.js";

import {
  normalizeConditionalFormatAddress,
  normalizeConditionalFormatType,
  normalizeOptionalBoolean,
} from "./conditional-format-normalization.js";

interface LoadedConditionalFormatEntry {
  conditionalFormat: Excel.ConditionalFormat;
  appliesTo: Excel.RangeAreas;
  normalizedType: RecoveryConditionalFormatRuleType;
}

const CONDITIONAL_FORMAT_RULE_HANDLERS = {
  ...BASIC_CONDITIONAL_FORMAT_RULE_HANDLERS,
  ...ADVANCED_CONDITIONAL_FORMAT_RULE_HANDLERS,
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
    const captureContext: ConditionalFormatRuleCaptureContext = {};
    const stopIfTrue = normalizeOptionalBoolean(entry.conditionalFormat.stopIfTrue);
    if (stopIfTrue !== undefined) {
      captureContext.stopIfTrue = stopIfTrue;
    }
    const appliesToAddress = normalizeConditionalFormatAddress(entry.appliesTo.address);
    if (appliesToAddress !== undefined) {
      captureContext.appliesToAddress = appliesToAddress;
    }
    const captureResult = handler.capture(entry.conditionalFormat, captureContext);

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

