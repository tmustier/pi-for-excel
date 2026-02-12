import { excelRun, getRange } from "../excel/helpers.js";
import { isRecord } from "../utils/type-guards.js";

export type RecoveryConditionalCellValueOperator =
  | "Between"
  | "NotBetween"
  | "EqualTo"
  | "NotEqualTo"
  | "GreaterThan"
  | "LessThan"
  | "GreaterThanOrEqual"
  | "LessThanOrEqual";

export interface RecoveryConditionalFormatRule {
  type: "custom" | "cell_value";
  stopIfTrue?: boolean;
  formula?: string;
  operator?: RecoveryConditionalCellValueOperator;
  formula1?: string;
  formula2?: string;
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

function isRecoveryConditionalCellValueOperator(value: unknown): value is RecoveryConditionalCellValueOperator {
  if (typeof value !== "string") return false;

  for (const operator of SUPPORTED_CELL_VALUE_OPERATORS) {
    if (operator === value) {
      return true;
    }
  }

  return false;
}

function normalizeConditionalFormatType(type: unknown): "custom" | "cell_value" | null {
  if (type === "Custom" || type === "custom") {
    return "custom";
  }

  if (type === "CellValue" || type === "cellValue") {
    return "cell_value";
  }

  return null;
}

function normalizeOptionalString(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function normalizeOptionalBoolean(value: unknown): boolean | undefined {
  return typeof value === "boolean" ? value : undefined;
}

function normalizeUnderline(value: unknown): boolean | undefined {
  if (typeof value === "boolean") return value;

  if (typeof value === "string") {
    return value !== "None";
  }

  return undefined;
}

function emptyCommentThreadState(): RecoveryCommentThreadState {
  return {
    exists: false,
    content: "",
    resolved: false,
    replies: [],
  };
}

function localAddressPart(address: string): string {
  const trimmed = address.trim();
  const separatorIndex = trimmed.lastIndexOf("!");
  if (separatorIndex < 0) {
    return trimmed;
  }

  return trimmed.slice(separatorIndex + 1);
}

export function firstCellAddress(address: string): string {
  const local = localAddressPart(address);
  const firstArea = local.split(",")[0] ?? local;
  const first = firstArea.split(":")[0] ?? firstArea;
  return first.trim();
}

function cloneRecoveryConditionalFormatRule(rule: RecoveryConditionalFormatRule): RecoveryConditionalFormatRule {
  return {
    type: rule.type,
    stopIfTrue: rule.stopIfTrue,
    formula: rule.formula,
    operator: rule.operator,
    formula1: rule.formula1,
    formula2: rule.formula2,
    fillColor: rule.fillColor,
    fontColor: rule.fontColor,
    bold: rule.bold,
    italic: rule.italic,
    underline: rule.underline,
    appliesToAddress: rule.appliesToAddress,
  };
}

export function cloneRecoveryConditionalFormatRules(
  rules: readonly RecoveryConditionalFormatRule[],
): RecoveryConditionalFormatRule[] {
  return rules.map((rule) => cloneRecoveryConditionalFormatRule(rule));
}

export function cloneRecoveryCommentThreadState(state: RecoveryCommentThreadState): RecoveryCommentThreadState {
  return {
    exists: state.exists,
    content: state.content,
    resolved: state.resolved,
    replies: [...state.replies],
  };
}

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
}

function normalizeConditionalFormatAddress(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

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
    entries.push({ conditionalFormat, appliesTo });

    if (normalizedType === "custom") {
      conditionalFormat.custom.load("rule");
      conditionalFormat.custom.format.fill.load("color");
      conditionalFormat.custom.format.font.load("bold,italic,underline,color");
      continue;
    }

    conditionalFormat.cellValue.load("rule");
    conditionalFormat.cellValue.format.fill.load("color");
    conditionalFormat.cellValue.format.font.load("bold,italic,underline,color");
  }

  await context.sync();

  const rules: RecoveryConditionalFormatRule[] = [];

  for (const entry of entries) {
    const conditionalFormat = entry.conditionalFormat;
    const normalizedType = normalizeConditionalFormatType(conditionalFormat.type);
    if (!normalizedType) {
      return {
        supported: false,
        rules: [],
        reason: `Unsupported conditional format type: ${String(conditionalFormat.type)}`,
      };
    }

    const appliesToAddress = normalizeConditionalFormatAddress(entry.appliesTo.address);

    if (normalizedType === "custom") {
      const custom = conditionalFormat.custom;
      rules.push({
        type: "custom",
        stopIfTrue: normalizeOptionalBoolean(conditionalFormat.stopIfTrue),
        formula: normalizeOptionalString(custom.rule.formula),
        appliesToAddress,
        ...captureRuleFormatting(custom.format),
      });
      continue;
    }

    const cellValue = conditionalFormat.cellValue;
    const ruleData = cellValue.rule;
    const operator = isRecord(ruleData) ? ruleData.operator : undefined;

    if (!isRecoveryConditionalCellValueOperator(operator)) {
      return {
        supported: false,
        rules: [],
        reason: "Unsupported conditional format rule operator.",
      };
    }

    const formula1 = isRecord(ruleData) ? ruleData.formula1 : undefined;
    const formula2 = isRecord(ruleData) ? ruleData.formula2 : undefined;

    if (typeof formula1 !== "string") {
      return {
        supported: false,
        rules: [],
        reason: "Conditional format rule is missing formula1.",
      };
    }

    rules.push({
      type: "cell_value",
      stopIfTrue: normalizeOptionalBoolean(conditionalFormat.stopIfTrue),
      operator,
      formula1,
      formula2: typeof formula2 === "string" ? formula2 : undefined,
      appliesToAddress,
      ...captureRuleFormatting(cellValue.format),
    });
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

  if (rule.type === "custom") {
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
    return;
  }

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

interface LoadedCommentThread {
  state: RecoveryCommentThreadState;
  comment: Excel.Comment | null;
}

async function loadCommentThreadInRange(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  range: Excel.Range,
): Promise<LoadedCommentThread> {
  range.load("address");

  const commentCollection = sheet.comments;
  commentCollection.load("items");
  await context.sync();

  if (commentCollection.items.length === 0) {
    return {
      state: emptyCommentThreadState(),
      comment: null,
    };
  }

  const entries = commentCollection.items.map((comment) => {
    comment.load("content,resolved");
    comment.replies.load("items");
    const location = comment.getLocation();
    location.load("address");
    return { comment, location };
  });

  await context.sync();

  const targetCell = firstCellAddress(range.address).toUpperCase();
  let match: { comment: Excel.Comment } | null = null;

  for (const entry of entries) {
    if (firstCellAddress(entry.location.address).toUpperCase() === targetCell) {
      match = { comment: entry.comment };
      break;
    }
  }

  if (!match) {
    return {
      state: emptyCommentThreadState(),
      comment: null,
    };
  }

  for (const reply of match.comment.replies.items) {
    reply.load("content");
  }

  if (match.comment.replies.items.length > 0) {
    await context.sync();
  }

  return {
    state: {
      exists: true,
      content: match.comment.content,
      resolved: match.comment.resolved,
      replies: match.comment.replies.items.map((reply) => reply.content),
    },
    comment: match.comment,
  };
}

export async function captureCommentThreadState(address: string): Promise<RecoveryCommentThreadState> {
  return excelRun<RecoveryCommentThreadState>(async (context) => {
    const { sheet, range } = getRange(context, address);
    const loaded = await loadCommentThreadInRange(context, sheet, range);
    return cloneRecoveryCommentThreadState(loaded.state);
  });
}

export async function applyCommentThreadState(
  address: string,
  targetState: RecoveryCommentThreadState,
): Promise<RecoveryCommentThreadState> {
  return excelRun<RecoveryCommentThreadState>(async (context) => {
    const { sheet, range } = getRange(context, address);
    const loaded = await loadCommentThreadInRange(context, sheet, range);

    if (!targetState.exists) {
      if (loaded.comment) {
        loaded.comment.delete();
        await context.sync();
      }

      return cloneRecoveryCommentThreadState(loaded.state);
    }

    if (loaded.comment) {
      loaded.comment.delete();
      await context.sync();
    }

    const restoredComment = sheet.comments.add(range, targetState.content);

    for (const reply of targetState.replies) {
      restoredComment.replies.add(reply);
    }

    restoredComment.resolved = targetState.resolved;
    await context.sync();

    return cloneRecoveryCommentThreadState(loaded.state);
  });
}
