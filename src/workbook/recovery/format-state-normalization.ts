import type { RecoveryFormatBorderState } from "./types.js";

export const RECOVERY_BORDER_KEYS = [
  "borderTop",
  "borderBottom",
  "borderLeft",
  "borderRight",
  "borderInsideHorizontal",
  "borderInsideVertical",
] as const;

export type RecoveryBorderKey = (typeof RECOVERY_BORDER_KEYS)[number];

type RecoveryBorderEdge =
  | "EdgeTop"
  | "EdgeBottom"
  | "EdgeLeft"
  | "EdgeRight"
  | "InsideHorizontal"
  | "InsideVertical";

export const BORDER_KEY_TO_EDGE: Record<RecoveryBorderKey, RecoveryBorderEdge> = {
  borderTop: "EdgeTop",
  borderBottom: "EdgeBottom",
  borderLeft: "EdgeLeft",
  borderRight: "EdgeRight",
  borderInsideHorizontal: "InsideHorizontal",
  borderInsideVertical: "InsideVertical",
};

type RecoveryUnderlineStyle = "None" | "Single" | "Double" | "SingleAccountant" | "DoubleAccountant";

const RECOVERY_UNDERLINE_STYLES: readonly RecoveryUnderlineStyle[] = [
  "None",
  "Single",
  "Double",
  "SingleAccountant",
  "DoubleAccountant",
];

export function isRecoveryUnderlineStyle(value: DynamicValue): value is RecoveryUnderlineStyle {
  if (typeof value !== "string") return false;

  for (const candidate of RECOVERY_UNDERLINE_STYLES) {
    if (candidate === value) return true;
  }

  return false;
}

type RecoveryHorizontalAlignment =
  | "General"
  | "Left"
  | "Center"
  | "Right"
  | "Fill"
  | "Justify"
  | "CenterAcrossSelection"
  | "Distributed";

const RECOVERY_HORIZONTAL_ALIGNMENTS: readonly RecoveryHorizontalAlignment[] = [
  "General",
  "Left",
  "Center",
  "Right",
  "Fill",
  "Justify",
  "CenterAcrossSelection",
  "Distributed",
];

export function isRecoveryHorizontalAlignment(value: DynamicValue): value is RecoveryHorizontalAlignment {
  if (typeof value !== "string") return false;

  for (const candidate of RECOVERY_HORIZONTAL_ALIGNMENTS) {
    if (candidate === value) return true;
  }

  return false;
}

type RecoveryVerticalAlignment = "Top" | "Center" | "Bottom" | "Justify" | "Distributed";

const RECOVERY_VERTICAL_ALIGNMENTS: readonly RecoveryVerticalAlignment[] = [
  "Top",
  "Center",
  "Bottom",
  "Justify",
  "Distributed",
];

export function isRecoveryVerticalAlignment(value: DynamicValue): value is RecoveryVerticalAlignment {
  if (typeof value !== "string") return false;

  for (const candidate of RECOVERY_VERTICAL_ALIGNMENTS) {
    if (candidate === value) return true;
  }

  return false;
}

type RecoveryRangeBorderStyle =
  | "None"
  | "Continuous"
  | "Dash"
  | "DashDot"
  | "DashDotDot"
  | "Dot"
  | "Double"
  | "SlantDashDot";

const RECOVERY_RANGE_BORDER_STYLES: readonly RecoveryRangeBorderStyle[] = [
  "None",
  "Continuous",
  "Dash",
  "DashDot",
  "DashDotDot",
  "Dot",
  "Double",
  "SlantDashDot",
];

function isRecoveryRangeBorderStyle(value: DynamicValue): value is RecoveryRangeBorderStyle {
  if (typeof value !== "string") return false;

  for (const candidate of RECOVERY_RANGE_BORDER_STYLES) {
    if (candidate === value) return true;
  }

  return false;
}

type RecoveryRangeBorderWeight = "Hairline" | "Thin" | "Medium" | "Thick";

const RECOVERY_RANGE_BORDER_WEIGHTS: readonly RecoveryRangeBorderWeight[] = [
  "Hairline",
  "Thin",
  "Medium",
  "Thick",
];

function isRecoveryRangeBorderWeight(value: DynamicValue): value is RecoveryRangeBorderWeight {
  if (typeof value !== "string") return false;

  for (const candidate of RECOVERY_RANGE_BORDER_WEIGHTS) {
    if (candidate === value) return true;
  }

  return false;
}

export function normalizeOptionalString(value: DynamicValue): string | undefined {
  return typeof value === "string" ? value : undefined;
}

export function normalizeOptionalBoolean(value: DynamicValue): boolean | undefined {
  return typeof value === "boolean" ? value : undefined;
}

export function normalizeOptionalNumber(value: DynamicValue): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

export function captureBorderState(border: Excel.RangeBorder): RecoveryFormatBorderState | null {
  const styleRaw = border.style;
  if (!isRecoveryRangeBorderStyle(styleRaw)) {
    return null;
  }

  const weightRaw = border.weight;
  const colorRaw = border.color;

  return {
    style: styleRaw,
    weight: isRecoveryRangeBorderWeight(weightRaw) ? weightRaw : undefined,
    color: normalizeOptionalString(colorRaw),
  };
}

export function applyBorderState(border: Excel.RangeBorder, state: RecoveryFormatBorderState): void {
  if (!isRecoveryRangeBorderStyle(state.style)) {
    throw new Error("Format checkpoint is invalid: border style is unsupported.");
  }

  border.style = state.style;

  if (state.style !== "None" && state.weight !== undefined) {
    if (!isRecoveryRangeBorderWeight(state.weight)) {
      throw new Error("Format checkpoint is invalid: border weight is unsupported.");
    }
    border.weight = state.weight;
  }

  if (typeof state.color === "string") {
    border.color = state.color;
  }
}
