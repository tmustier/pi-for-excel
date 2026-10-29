import type { WorkbookCellChangeSummary } from "./cell-diff.js";

export const MAX_EXPLANATION_PROMPT_CHARS = 1_200;
export const MAX_EXPLANATION_TEXT_CHARS = 420;

const MAX_CHANGE_SAMPLES = 6;
const MAX_SAMPLE_VALUE_CHARS = 40;
const MAX_CITATIONS = 8;

export interface ChangeExplanationInput {
  toolName: string;
  blocked: boolean;
  changedCount?: number;
  summary?: string;
  error?: string;
  inputAddress?: string;
  outputAddress?: string;
  changes?: WorkbookCellChangeSummary;
}

export interface ChangeExplanation {
  prompt: string;
  text: string;
  citations: string[];
  truncated: boolean;
  usedFallback: boolean;
}

function trimToLength(value: string, maxChars: number): { text: string; truncated: boolean } {
  if (value.length <= maxChars) {
    return { text: value, truncated: false };
  }

  if (maxChars <= 1) {
    return { text: value.slice(0, Math.max(maxChars, 0)), truncated: true };
  }

  return {
    text: `${value.slice(0, maxChars - 1)}…`,
    truncated: true,
  };
}

function sanitizeInline(value: string): string {
  return value.replace(/\s+/gu, " ").trim();
}

function normalizeSampleValue(value: string): { value: string; truncated: boolean } {
  const normalized = sanitizeInline(value);
  const trimmed = trimToLength(normalized, MAX_SAMPLE_VALUE_CHARS);
  return {
    value: trimmed.text,
    truncated: trimmed.truncated,
  };
}

function uniqueCitations(input: ChangeExplanationInput): string[] {
  const citations: string[] = [];

  const addCitation = (address: string | undefined): void => {
    if (!address) return;
    const trimmed = address.trim();
    if (trimmed.length === 0) return;
    if (citations.includes(trimmed)) return;
    citations.push(trimmed);
  };

  if (input.changes) {
    for (const change of input.changes.sample) {
      addCitation(change.address);
      if (citations.length >= MAX_CITATIONS) {
        return citations;
      }
    }
  }

  addCitation(input.outputAddress);
  addCitation(input.inputAddress);

  return citations.slice(0, MAX_CITATIONS);
}

function buildPrompt(input: ChangeExplanationInput): { prompt: string; truncated: boolean } {
  const lines: string[] = [
    "You explain spreadsheet mutations in plain language.",
    `Tool: ${input.toolName}`,
    `Blocked: ${input.blocked ? "yes" : "no"}`,
    `Changed count: ${typeof input.changedCount === "number" ? String(input.changedCount) : "unknown"}`,
    `Output address: ${input.outputAddress ?? "(none)"}`,
    `Input address: ${input.inputAddress ?? "(none)"}`,
    `Summary: ${input.summary ?? "(none)"}`,
    `Error: ${input.error ?? "(none)"}`,
    "Sample changes:",
  ];

  let truncated = false;

  if (input.changes && input.changes.sample.length > 0) {
    const samples = input.changes.sample.slice(0, MAX_CHANGE_SAMPLES);
    if (input.changes.sample.length > MAX_CHANGE_SAMPLES) {
      truncated = true;
    }

    for (const change of samples) {
      const before = normalizeSampleValue(change.beforeValue);
      const after = normalizeSampleValue(change.afterValue);
      truncated = truncated || before.truncated || after.truncated;

      const formulaChanged = change.beforeFormula !== change.afterFormula;
      const formulaSuffix = formulaChanged ? " (formula changed)" : "";
      lines.push(`- ${change.address}: ${before.value} -> ${after.value}${formulaSuffix}`);
    }
  } else {
    lines.push("- (none)");
  }

  lines.push("Return a concise explanation (2-4 sentences) and reference cited addresses when available.");

  const prompt = lines.join("\n");
  const boundedPrompt = trimToLength(prompt, MAX_EXPLANATION_PROMPT_CHARS);

  return {
    prompt: boundedPrompt.text,
    truncated: truncated || boundedPrompt.truncated,
  };
}

function buildStatusLine(input: ChangeExplanationInput): string {
  if (input.blocked) {
    return "This mutation was blocked, so workbook data was not changed.";
  }

  if (typeof input.changedCount === "number") {
    if (input.changedCount > 0) {
      return `This operation changed ${input.changedCount} cell${input.changedCount === 1 ? "" : "s"}.`;
    }

    return "This operation completed without changing cell values.";
  }

  return "This operation mutated workbook state.";
}

function buildDetailLine(input: ChangeExplanationInput): string | null {
  if (input.error) {
    return `Reason: ${sanitizeInline(input.error)}.`;
  }

  if (input.summary) {
    return `Summary: ${sanitizeInline(input.summary)}.`;
  }

  if (input.outputAddress) {
    return `Primary target: ${input.outputAddress}.`;
  }

  return null;
}

function buildSampleLine(input: ChangeExplanationInput): { text: string | null; truncated: boolean } {
  if (!input.changes || input.changes.sample.length === 0) {
    return { text: null, truncated: false };
  }

  const sample = input.changes.sample.slice(0, 2);
  const snippets: string[] = [];
  let truncated = input.changes.sample.length > 2;

  for (const change of sample) {
    const before = normalizeSampleValue(change.beforeValue);
    const after = normalizeSampleValue(change.afterValue);
    truncated = truncated || before.truncated || after.truncated;

    snippets.push(`${change.address}: ${before.value} -> ${after.value}`);
  }

  return {
    text: `Examples: ${snippets.join("; ")}.`,
    truncated,
  };
}

function buildCitationLine(citations: readonly string[]): string | null {
  if (citations.length === 0) return null;

  const shown = citations.slice(0, 3);
  const suffix = citations.length > shown.length ? ` (+${citations.length - shown.length} more)` : "";
  return `Inspect: ${shown.join(", ")}${suffix}.`;
}

export function buildChangeExplanation(input: ChangeExplanationInput): ChangeExplanation {
  const citations = uniqueCitations(input);
  const promptResult = buildPrompt(input);

  const statusLine = buildStatusLine(input);
  const detailLine = buildDetailLine(input);
  const sampleLine = buildSampleLine(input);
  const citationLine = buildCitationLine(citations);

  const lines = [statusLine, detailLine, sampleLine.text, citationLine]
    .filter((line): line is string => typeof line === "string" && line.trim().length > 0);

  let usedFallback = false;

  if (lines.length === 0) {
    usedFallback = true;
    lines.push("Not enough audit metadata is available to explain this change yet.");
  }

  if (
    !input.summary &&
    !input.error &&
    !input.outputAddress &&
    !input.inputAddress &&
    (!input.changes || input.changes.sample.length === 0)
  ) {
    usedFallback = true;
    if (lines.length < 2) {
      lines.push("Not enough audit metadata is available to explain this change yet.");
    }
  }

  const textResult = trimToLength(lines.join("\n\n"), MAX_EXPLANATION_TEXT_CHARS);

  return {
    prompt: promptResult.prompt,
    text: textResult.text,
    citations,
    truncated: promptResult.truncated || sampleLine.truncated || textResult.truncated,
    usedFallback,
  };
}
