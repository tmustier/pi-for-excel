export const MAX_PREVIEW_CHARS = 120;

const FUNCTION_SUMMARIES: Readonly<Record<string, string>> = {
  SUM: "adds values",
  AVERAGE: "calculates an average",
  MIN: "returns the smallest value",
  MAX: "returns the largest value",
  IF: "applies conditional logic",
  IFS: "evaluates multiple conditional branches",
  IFERROR: "substitutes a fallback when an error occurs",
  XLOOKUP: "looks up a matching value and returns a related result",
  VLOOKUP: "looks up a value in the first column and returns a related result",
  HLOOKUP: "looks up a value in the first row and returns a related result",
  INDEX: "returns a value from a row/column position",
  MATCH: "finds a position within a range",
  SUMIFS: "adds values that meet multiple criteria",
  COUNTIF: "counts cells that meet one criterion",
  COUNTIFS: "counts cells that meet multiple criteria",
  ROUND: "rounds a number",
  ROUNDUP: "rounds a number up",
  ROUNDDOWN: "rounds a number down",
  TEXT: "formats a value as text",
  CONCAT: "concatenates values into text",
  CONCATENATE: "concatenates values into text",
};

function truncate(value: string, maxChars: number): string {
  if (value.length <= maxChars) return value;
  return `${value.slice(0, Math.max(maxChars - 1, 1))}…`;
}

export function previewCellValue(value: DynamicValue): string {
  if (value === null || value === undefined || value === "") {
    return "(blank)";
  }

  let rendered: string;
  if (typeof value === "string") {
    rendered = value;
  } else if (typeof value === "number" || typeof value === "boolean") {
    rendered = String(value);
  } else {
    rendered = JSON.stringify(value);
  }

  return truncate(rendered, MAX_PREVIEW_CHARS);
}

export function extractFormulaFunctionNames(formula: string): string[] {
  const seen = new Set<string>();
  const names: string[] = [];

  const pattern = /\b([A-Za-z][A-Za-z0-9_.]*)\s*\(/gu;
  for (const match of formula.matchAll(pattern)) {
    const candidate = match[1]?.toUpperCase();
    if (!candidate || seen.has(candidate)) continue;
    seen.add(candidate);
    names.push(candidate);
  }

  return names;
}

function formatFunctionList(functionNames: readonly string[]): string {
  if (functionNames.length === 0) return "computes a result from referenced cells";

  const labels = functionNames.map((name) => FUNCTION_SUMMARIES[name] ?? `uses ${name}`);
  const [firstLabel, secondLabel] = labels;
  if (firstLabel === undefined) return "computes a result from referenced cells";
  if (secondLabel === undefined) return firstLabel;
  if (labels.length === 2) return `${firstLabel} and ${secondLabel}`;
  const lastLabel = labels[labels.length - 1];
  if (lastLabel === undefined) return `${firstLabel} and ${secondLabel}`;
  return `${labels.slice(0, -1).join(", ")}, and ${lastLabel}`;
}

export interface ExplainFormulaNarrativeInput {
  valuePreview: string;
  functionNames: readonly string[];
  referenceCount: number;
  truncated: boolean;
}

export function buildExplainFormulaNarrative(input: ExplainFormulaNarrativeInput): string {
  const refsLabel = `${input.referenceCount} direct reference${input.referenceCount === 1 ? "" : "s"}`;
  const functionSummary = formatFunctionList(input.functionNames);

  const parts = [
    `Current value: ${input.valuePreview}.`,
    `The formula ${functionSummary} across ${refsLabel}.`,
  ];

  if (input.truncated) {
    parts.push("Reference preview is truncated; inspect cited cells for the complete lineage.");
  }

  return parts.join(" ");
}
