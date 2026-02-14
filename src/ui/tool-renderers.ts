/**
 * Tool renderers for Pi-for-Excel.
 *
 * Renders Excel tool calls as compact, collapsed-by-default cards with
 * human-readable descriptions. Expand to see raw Input/Output.
 */

import type { ImageContent, TextContent, ToolResultMessage } from "@mariozechner/pi-ai";
import { registerToolRenderer } from "@mariozechner/pi-web-ui/dist/tools/renderer-registry.js";
import type { ToolRenderer, ToolRenderResult } from "@mariozechner/pi-web-ui/dist/tools/types.js";
import { html, type TemplateResult } from "lit";
import { createRef, ref } from "lit/directives/ref.js";
import { renderCollapsibleToolCardHeader, renderToolCardHeader } from "./tool-card-header.js";
import { cellRef, cellRefDisplay, cellRefs } from "./cell-link.js";
import { humanizeToolInput } from "./humanize-params.js";
import { humanizeColorsInText } from "./color-names.js";
import { stripYamlFrontmatter } from "./markdown-preprocess.js";
import { TOOL_NAMES_WITH_RENDERER, type UiToolName } from "../tools/capabilities.js";
import {
  isCommentsDetails,
  isConditionalFormatDetails,
  isExplainFormulaDetails,
  isFillFormulaDetails,
  isFormatCellsDetails,
  isModifyStructureDetails,
  isPythonTransformRangeDetails,
  isReadRangeCsvDetails,
  isSkillsReadDetails,
  isTraceDependenciesDetails,
  isViewSettingsDetails,
  isWorkbookHistoryDetails,
  isWriteCellsDetails,
  type RecoveryCheckpointDetails,
  type WriteCellsDetails,
} from "../tools/tool-details.js";
import { getToolExecutionMode } from "../tools/execution-policy.js";
import {
  buildChangeExplanation,
  type ChangeExplanationInput,
} from "../audit/change-explanation.js";
import { renderCsvTable } from "./render-csv-table.js";
import { renderDepTree } from "./render-dep-tree.js";

// Ensure <markdown-block> custom element is registered before we render it.
import "@mariozechner/mini-lit/dist/MarkdownBlock.js";

type ToolState = "inprogress" | "complete" | "error";
type SupportedToolName = UiToolName;

/* ── Helpers ────────────────────────────────────────────────── */

function formatParamsJson(params: unknown): string {
  if (params === undefined) return "";

  try {
    if (typeof params === "string") {
      try {
        return JSON.stringify(JSON.parse(params), null, 2);
      } catch {
        return params;
      }
    }
    return JSON.stringify(params, null, 2);
  } catch {
    return typeof params === "string" || typeof params === "number" || typeof params === "boolean" ? String(params) : JSON.stringify(params);
  }
}

function safeParseParams(params: unknown): Record<string, unknown> {
  if (!params) return {};
  if (typeof params === "object" && params !== null) return params as Record<string, unknown>;
  if (typeof params === "string") {
    try {
      const parsed: unknown = JSON.parse(params);
      if (typeof parsed === "object" && parsed !== null) return parsed as Record<string, unknown>;
      return {};
    } catch { return {}; }
  }
  return {};
}

function splitToolResultContent(result: ToolResultMessage<unknown>): {
  text: string;
  images: ImageContent[];
} {
  const text = (result.content ?? [])
    .filter((c): c is TextContent => c.type === "text")
    .map((c) => c.text)
    .join("\n");

  const images = (result.content ?? []).filter((c): c is ImageContent => c.type === "image");

  return { text, images };
}

function tryFormatJsonOutput(text: string): { isJson: boolean; formatted: string } {
  const trimmed = text.trim();
  if (!trimmed) return { isJson: false, formatted: text };

  try {
    const parsed: unknown = JSON.parse(trimmed);
    return { isJson: true, formatted: JSON.stringify(parsed, null, 2) };
  } catch {
    return { isJson: false, formatted: text };
  }
}

/**
 * Heuristic: does the text contain markdown syntax that benefits from
 * rendering via `<markdown-block>` rather than plain text?
 *
 * Checks for: tables, headers, lists, bold/italic, links, code fences,
 * blockquotes, horizontal rules, and emoji sentinels (✅ ⛔ etc.).
 */
function looksLikeMarkdown(text: string): boolean {
  // Table rows: "| ... | ... |"
  if (/^\s*\|.+\|/m.test(text)) return true;
  // ATX headers: "# ", "## ", etc.
  if (/^#{1,6}\s+\S/m.test(text)) return true;
  // Unordered list items: "- item" or "* item"
  if (/^[ \t]*[-*]\s+\S/m.test(text)) return true;
  // Ordered list items: "1. item"
  if (/^[ \t]*\d+\.\s+\S/m.test(text)) return true;
  // Bold / italic
  if (/\*\*[^*]+\*\*/.test(text)) return true;
  if (/__[^_]+__/.test(text)) return true;
  // Links: [text](url)
  if (/\[[^\]]+\]\([^)]+\)/.test(text)) return true;
  // Fenced code blocks
  if (/^```/m.test(text)) return true;
  // Blockquotes: "> "
  if (/^>\s+\S/m.test(text)) return true;
  // Horizontal rules: "---" or "***" or "___" (alone on a line)
  if (/^[-*_]{3,}\s*$/m.test(text)) return true;
  // Common sentinels our tools emit (emoji prefixes)
  if (/^(?:✅|⛔|⚠️|ℹ️|📊|📋|🔍)/m.test(text)) return true;

  return false;
}

function stripMarkdownInline(line: string): string {
  return line
    .replace(/^#+\s+/, "")
    .replace(/^\s*[-*]\s+/, "")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/__([^_]+)__/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .trim();
}

function extractSummaryLine(text: string): string | null {
  for (const rawLine of text.split("\n")) {
    const t = rawLine.trim();
    if (!t) continue;
    if (t.startsWith("|")) continue;

    const stripped = stripMarkdownInline(t);
    if (stripped) return stripped;
  }
  return null;
}

function detectStandaloneImagePath(text: string): string | null {
  const t = text.trim();
  if (!t) return null;
  if (t.includes("\n")) return null;

  const isImage = /\.(png|jpe?g|gif|webp|svg)$/i.test(t);
  if (!isImage) return null;

  const isUnixAbs = t.startsWith("/");
  const isWinAbs = /^[A-Za-z]:\\/.test(t);
  const isFileUrl = t.startsWith("file://");

  return isUnixAbs || isWinAbs || isFileUrl ? t : null;
}

function pathBasename(path: string): string {
  const parts = path.split(/[\\/]/).filter(Boolean);
  return parts[parts.length - 1] ?? path;
}

function toFileUrl(path: string): string {
  if (path.startsWith("file://")) return path;

  const win = /^([A-Za-z]):\\(.*)$/.exec(path);
  if (win) {
    const drive = win[1].toUpperCase();
    const rest = win[2]
      .split("\\")
      .map((seg) => encodeURIComponent(seg))
      .join("/");
    return `file:///${drive}:/${rest}`;
  }

  const encoded = path
    .split("/")
    .map((seg) => encodeURIComponent(seg))
    .join("/");
  return `file://${encoded}`;
}

function renderImages(images: ImageContent[]): TemplateResult {
  if (!images.length) return html``;

  return html`
    <div class="mt-2 grid grid-cols-1 gap-2">
      ${images.map((img) => {
        const src = `data:${img.mimeType};base64,${img.data}`;
        return html`
          <div class="border border-border rounded-lg overflow-hidden bg-background">
            <img src=${src} class="block w-full h-auto" />
          </div>
        `;
      })}
    </div>
  `;
}

function getWorkbookCellChanges(details: unknown): WriteCellsDetails["changes"] | undefined {
  if (isWriteCellsDetails(details)) {
    return details.changes;
  }

  if (isFillFormulaDetails(details)) {
    return details.changes;
  }

  if (isPythonTransformRangeDetails(details)) {
    return details.changes;
  }

  return undefined;
}

function formatDiffValue(value: string): string {
  return value.length > 0 ? value : "∅";
}

function renderWorkbookCellDiff(details: unknown): TemplateResult {
  const changes = getWorkbookCellChanges(details);
  if (!changes || changes.changedCount <= 0) return html``;

  return html`
    <div class="pi-tool-card__section">
      <div class="pi-tool-card__section-label">Changes (${changes.changedCount})</div>
      <div class="pi-tool-card__diff">
        <table class="pi-tool-card__diff-table">
          <thead>
            <tr>
              <th>Cell</th>
              <th>Before</th>
              <th>After</th>
            </tr>
          </thead>
          <tbody>
            ${changes.sample.map((change) => html`
              <tr>
                <td class="pi-tool-card__diff-cell">${cellRef(change.address)}</td>
                <td>
                  <div class="pi-tool-card__diff-value">${formatDiffValue(change.beforeValue)}</div>
                  ${change.beforeFormula || change.afterFormula
                    ? html`<div class="pi-tool-card__diff-formula">ƒ ${change.beforeFormula ?? "—"}</div>`
                    : html``}
                </td>
                <td>
                  <div class="pi-tool-card__diff-value">${formatDiffValue(change.afterValue)}</div>
                  ${change.beforeFormula || change.afterFormula
                    ? html`<div class="pi-tool-card__diff-formula">ƒ ${change.afterFormula ?? "—"}</div>`
                    : html``}
                </td>
              </tr>
            `)}
          </tbody>
        </table>
        ${changes.truncated
          ? html`<div class="pi-tool-card__diff-note">Showing first ${changes.sample.length} changed cell(s).</div>`
          : html``}
      </div>
    </div>
  `;
}

function renderExplainFormulaDetails(details: unknown): TemplateResult | null {
  if (!isExplainFormulaDetails(details)) return null;

  if (!details.hasFormula) {
    return html`<div class="pi-tool-card__plain-text">${details.explanation}</div>`;
  }

  return html`
    <div class="pi-formula-explain">
      <div class="pi-tool-card__plain-text">${details.explanation}</div>
      <div class="pi-formula-explain__formula">
        <span class="pi-formula-explain__label">Formula:</span>
        <code>${details.formula}</code>
      </div>
      <div class="pi-formula-explain__refs">
        <div class="pi-formula-explain__label">Direct references (${details.references.length})</div>
        <ul class="pi-formula-explain__list">
          ${details.references.length > 0
            ? details.references.map((reference) => html`
              <li>
                ${cellRefs(reference.address)}
                ${reference.valuePreview ? html`<span class="pi-formula-explain__preview"> → ${reference.valuePreview}</span>` : html``}
              </li>
            `)
            : html`<li>(No direct references detected)</li>`}
        </ul>
      </div>
      ${details.truncated
        ? html`<div class="pi-tool-card__diff-note">Showing the first ${details.references.length} reference(s).</div>`
        : html``}
    </div>
  `;
}

function optionalString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim().length > 0 ? value : undefined;
}

function extractResultError(resultText: string | undefined): string | undefined {
  if (!resultText) return undefined;

  const summary = resultSummary(resultText);
  if (!summary) return undefined;

  const normalized = summary.replace(/^\s*⚠️\s*/u, "").trim();
  if (/^error\b/ui.test(normalized)) {
    return normalized.replace(/^error:\s*/ui, "").trim();
  }

  return undefined;
}

function buildChangeExplanationInputForTool(
  toolName: SupportedToolName,
  params: unknown,
  resultText: string | undefined,
  details: unknown,
): ChangeExplanationInput | null {
  if (getToolExecutionMode(toolName, params) !== "mutate") return null;

  const p = safeParseParams(params);
  const range = optionalString(p.range);
  const startCell = optionalString(p.start_cell);
  const sheet = optionalString(p.sheet);
  const action = optionalString(p.action);

  const summary = resultSummary(resultText ?? "") ?? undefined;
  const error = extractResultError(resultText);
  const blockedFromText = Boolean(resultText && isBlocked(resultText));

  if (isWriteCellsDetails(details)) {
    return {
      toolName,
      blocked: details.blocked,
      changedCount: details.changes?.changedCount,
      summary,
      error,
      outputAddress: details.address ?? startCell,
      changes: details.changes,
    };
  }

  if (isFillFormulaDetails(details)) {
    return {
      toolName,
      blocked: details.blocked,
      changedCount: details.changes?.changedCount,
      summary,
      error,
      outputAddress: details.address ?? range,
      changes: details.changes,
    };
  }

  if (isPythonTransformRangeDetails(details)) {
    return {
      toolName,
      blocked: details.blocked || blockedFromText,
      changedCount: details.changes?.changedCount,
      summary,
      error: details.error ?? error,
      inputAddress: details.inputAddress,
      outputAddress: details.outputAddress,
      changes: details.changes,
    };
  }

  if (isWorkbookHistoryDetails(details) && details.action === "restore") {
    const historyError = optionalString(details.error);
    return {
      toolName,
      blocked: blockedFromText || Boolean(historyError),
      changedCount: details.changedCount,
      summary,
      error: historyError ?? error,
      outputAddress: details.address,
    };
  }

  if (toolName === "format_cells") {
    const detailsAddress = isFormatCellsDetails(details) ? details.address : undefined;
    return {
      toolName,
      blocked: blockedFromText,
      summary,
      error,
      outputAddress: detailsAddress ?? range,
    };
  }

  if (toolName === "conditional_format") {
    return {
      toolName,
      blocked: blockedFromText,
      summary,
      error,
      outputAddress: range,
    };
  }

  if (toolName === "modify_structure") {
    return {
      toolName,
      blocked: blockedFromText,
      summary,
      error,
      outputAddress: range ?? sheet,
    };
  }

  if (toolName === "comments") {
    return {
      toolName,
      blocked: blockedFromText,
      changedCount: blockedFromText ? 0 : 1,
      summary,
      error,
      outputAddress: range,
    };
  }

  if (toolName === "view_settings") {
    const outputAddress = range ?? sheet;
    return {
      toolName,
      blocked: blockedFromText,
      changedCount: blockedFromText ? 0 : 1,
      summary,
      error,
      outputAddress,
    };
  }

  if (toolName === "workbook_history" && action === "restore") {
    return {
      toolName,
      blocked: blockedFromText,
      summary,
      error,
    };
  }

  return null;
}

function renderCitations(citations: readonly string[]): TemplateResult {
  if (citations.length === 0) {
    return html`<span class="pi-tool-card__explain-citations-empty">No range citations available.</span>`;
  }

  return html`${citations.map((address, index) => html`${index > 0 ? html`, ` : html``}${cellRefs(address)}`)}`;
}

function renderChangeExplanationSection(
  toolName: SupportedToolName,
  params: unknown,
  resultText: string | undefined,
  details: unknown,
): TemplateResult {
  const input = buildChangeExplanationInputForTool(toolName, params, resultText, details);
  if (!input) return html``;

  const explanation = buildChangeExplanation(input);

  return html`
    <div class="pi-tool-card__section">
      <details class="pi-tool-card__explain">
        <summary class="pi-tool-card__explain-toggle">Explain these changes</summary>
        <div class="pi-tool-card__explain-body">
          <div class="pi-tool-card__plain-text">${explanation.text}</div>
          <div class="pi-tool-card__explain-citations">
            <span class="pi-tool-card__explain-citations-label">Citations:</span>
            ${renderCitations(explanation.citations)}
          </div>
          ${explanation.usedFallback
            ? html`<div class="pi-tool-card__diff-note">Limited metadata available; explanation is high-level.</div>`
            : html``}
          ${explanation.truncated
            ? html`<div class="pi-tool-card__diff-note">Explanation uses a bounded metadata sample.</div>`
            : html``}
        </div>
      </details>
    </div>
  `;
}

/* ── Human-readable descriptions ────────────────────────────── */

/** Strip "(N×M)" / "(NxM)" dimension notation — not intuitive for users. */
function stripDimensions(text: string): string {
  return text.replace(/\s*\(\d+[×x]\d+\)/gi, "").trim();
}

/**
 * Compact multi-range addresses by factoring out a shared sheet prefix.
 *   "Summary!A3,Summary!A13,Summary!A22" → "Summary!A3,A13,A22"
 *   "Costs!A18:C18, Costs!A19:C19"       → "Costs!A18:C18,A19:C19"
 */
function compactRange(range: string): string {
  const parts = range.split(/\s*,\s*/);
  if (parts.length <= 1) return range;

  const parsed = parts.map((p) => {
    const bang = p.indexOf("!");
    return bang >= 0
      ? { sheet: p.substring(0, bang), addr: p.substring(bang + 1) }
      : { sheet: "", addr: p };
  });

  const first = parsed[0].sheet;
  if (first && parsed.every((p) => p.sheet === first)) {
    return `${first}!${parsed.map((p) => p.addr).join(",")}`;
  }
  return range;
}

function qualifyRangeAddress(range: string | undefined, sheet: string | undefined): string | undefined {
  if (!range) return undefined;
  if (range.includes("!")) return range;
  return sheet ? `${sheet}!${range}` : range;
}

/**
 * Compact sheet-qualified ranges inside bold markdown markers.
 *   "Formatted **Sheet1!A1,Sheet1!B2**: ..." → "Formatted **Sheet1!A1, B2**: ..."
 */
function compactRangesInMarkdown(text: string): string {
  return text.replace(/\*\*([^*]+)\*\*/g, (_match, inner: string) => {
    if (!inner.includes("!")) return `**${inner}**`;
    const compacted = compactRange(inner);
    // Add spaces after commas for readability
    const spaced = compacted.replace(/,(?!\s)/g, ", ");
    return `**${spaced}**`;
  });
}

/** Extract target address from write_cells / fill_formula result text. */
function extractWrittenAddress(text: string): string | null {
  // "Written to **Sheet1!A1:C10** (…)" or "Filled formula across **Sheet1!A1:B20** (…)"
  const m = /(?:Written to|Filled formula across)\s+\*\*([^*]+)\*\*/.exec(text);
  return m ? m[1] : null;
}

/** Count formula errors mentioned in tool result text. */
function countResultErrors(text: string): number {
  const m = /(\d+)\s+formula error/i.exec(text);
  return m ? parseInt(m[1], 10) : 0;
}

/** True when result text starts with the blocked sentinel. */
function isBlocked(text: string): boolean {
  return text.trimStart().startsWith("⛔");
}

/**
 * Tools whose success result just echoes the Input section.
 * For these, we show a compact "✓ Done" instead of the full text.
 * Only applies when the result isn't an error/block/warning.
 */
const ECHO_RESULT_PATTERNS: Record<string, RegExp> = {
  format_cells: /^Formatted\s+\*\*/,
  conditional_format: /^(?:Added|Cleared) conditional format/,
};

/** True when the result is a redundant success echo of the input. */
function isEchoResult(toolName: string, text: string): boolean {
  const pattern = ECHO_RESULT_PATTERNS[toolName];
  if (!pattern) return false;
  const trimmed = text.trim();
  // Don't suppress if there are warnings or errors
  if (/⚠️|error|warning/i.test(trimmed)) return false;
  if (isBlocked(trimmed)) return false;
  return pattern.test(trimmed);
}

/** Result-aware summary for result text that is already user-friendly. */
function resultSummary(text: string): string | null {
  const line = extractSummaryLine(text);
  return line ? stripDimensions(line) : null;
}

function mutationBadge(changedCount: number | undefined, errorCount: number | undefined): string {
  const parts: string[] = [];

  if (typeof changedCount === "number" && changedCount > 0) {
    parts.push(`${changedCount} changed`);
  }

  if (typeof errorCount === "number" && errorCount > 0) {
    parts.push(`${errorCount} error${errorCount !== 1 ? "s" : ""}`);
  }

  return parts.length > 0 ? ` — ${parts.join(", ")}` : "";
}

function withRecoveryBadge(base: string, recovery: RecoveryCheckpointDetails | undefined): string {
  if (!recovery || recovery.status !== "not_available") {
    return base;
  }

  return base.length > 0 ? `${base}, no backup` : " — no backup";
}

function recoveryBadgeForDetails(details: unknown): string {
  if (isFormatCellsDetails(details)) {
    return withRecoveryBadge("", details.recovery);
  }

  if (isConditionalFormatDetails(details)) {
    return withRecoveryBadge("", details.recovery);
  }

  if (isModifyStructureDetails(details)) {
    return withRecoveryBadge("", details.recovery);
  }

  if (isCommentsDetails(details)) {
    return withRecoveryBadge("", details.recovery);
  }

  if (isViewSettingsDetails(details)) {
    return withRecoveryBadge("", details.recovery);
  }

  return "";
}

/** Append error / blocked badge to the detail string. */
function badge(
  toolName: SupportedToolName,
  resultText: string | undefined,
  details: unknown,
): string {
  if (toolName === "write_cells" && isWriteCellsDetails(details)) {
    if (details.blocked) return " — blocked";
    return withRecoveryBadge(
      mutationBadge(details.changes?.changedCount, details.formulaErrorCount),
      details.recovery,
    );
  }

  if (toolName === "fill_formula" && isFillFormulaDetails(details)) {
    if (details.blocked) return " — blocked";
    return withRecoveryBadge(
      mutationBadge(details.changes?.changedCount, details.formulaErrorCount),
      details.recovery,
    );
  }

  if (toolName === "python_transform_range" && isPythonTransformRangeDetails(details)) {
    if (details.blocked) return " — blocked";
    if (typeof details.error === "string" && details.error.length > 0) return " — error";
    return withRecoveryBadge(
      mutationBadge(details.changes?.changedCount, details.formulaErrorCount),
      details.recovery,
    );
  }

  if (!resultText) return "";
  if (isBlocked(resultText)) return " — blocked";
  const n = countResultErrors(resultText);
  if (n > 0) return ` — ${n} error${n !== 1 ? "s" : ""}`;
  return "";
}

interface ToolDesc {
  /** Bold verb, e.g. "Read", "Wrote", "Format" */
  action: string;
  /** Normal-weight rest, e.g. "Costs!A1:C19" */
  detail: string;
  /** Raw Excel address for click-to-navigate. Omit for non-range details. */
  address?: string;
}

/** Split a result-text summary line into action (first word) + rest. */
function splitFirstWord(text: string): ToolDesc {
  const i = text.indexOf(" ");
  return i > 0
    ? { action: text.substring(0, i), detail: text.substring(i + 1) }
    : { action: text, detail: "" };
}

/** Structured description: bold action + normal-weight detail. */
function describeToolCall(
  toolName: SupportedToolName,
  params: unknown,
  resultText: string | undefined,
  details: unknown,
): ToolDesc {
  const p = safeParseParams(params);
  const range = p.range as string | undefined;
  const startCell = p.start_cell as string | undefined;

  switch (toolName) {
    // ── Read tools ──
    case "read_range": {
      const mode = p.mode as string | undefined;
      const label = mode === "csv" ? "Export" : "Read";
      return { action: label, detail: range ? compactRange(range) + (mode === "csv" ? " (CSV)" : "") : "range", address: range };
    }
    case "get_workbook_overview": {
      const sheet = p.sheet as string | undefined;
      return { action: "Overview", detail: sheet ?? "" };
    }

    // ── Write tools ──
    case "write_cells": {
      const b = badge(toolName, resultText, details);

      if (isWriteCellsDetails(details) && details.address) {
        const action = details.blocked ? "Write" : "Edit";
        return { action, detail: details.address + b, address: details.address };
      }

      const addr = resultText ? extractWrittenAddress(resultText) : null;
      return addr
        ? { action: "Edit", detail: addr + b, address: addr }
        : { action: "Write", detail: (startCell ?? "cells") + b, address: startCell };
    }
    case "fill_formula": {
      const b = badge(toolName, resultText, details);

      if (isFillFormulaDetails(details) && details.address) {
        const action = details.blocked ? "Fill" : "Filled";
        return { action, detail: details.address + b, address: details.address };
      }

      const addr = resultText ? extractWrittenAddress(resultText) : null;
      return addr
        ? { action: "Filled", detail: addr + b, address: addr }
        : { action: "Fill", detail: (range ? compactRange(range) : "formula") + b, address: range };
    }
    case "python_transform_range": {
      const b = badge(toolName, resultText, details);

      if (isPythonTransformRangeDetails(details)) {
        const address = details.outputAddress ?? details.inputAddress;
        if (address) {
          const hasError = typeof details.error === "string" && details.error.length > 0;
          const action = details.blocked || hasError ? "Transform" : "Transformed";
          return { action, detail: address + b, address };
        }
      }

      const outputStart = p.output_start_cell as string | undefined;
      const fallbackAddress = outputStart ?? range;
      return {
        action: "Transform",
        detail: (fallbackAddress ?? "range") + b,
        address: fallbackAddress,
      };
    }

    // ── Format tools ──
    case "format_cells": {
      const addr = isFormatCellsDetails(details) ? details.address : undefined;
      const resolved = addr ?? range;
      const recovery = recoveryBadgeForDetails(details);
      return {
        action: "Format",
        detail: (resolved ? compactRange(resolved) : "cells") + recovery,
        address: resolved,
      };
    }
    case "conditional_format": {
      const recovery = recoveryBadgeForDetails(details);
      return {
        action: "Cond. format",
        detail: (range ? compactRange(range) : "cells") + recovery,
        address: range,
      };
    }

    // ── Result-text tools (split first word as action) ──
    case "modify_structure": {
      const recovery = recoveryBadgeForDetails(details);

      if (resultText) {
        const s = resultSummary(resultText);
        if (s) {
          const parts = splitFirstWord(s);
          return { ...parts, detail: `${parts.detail}${recovery}` };
        }
      }

      const act = p.action as string | undefined;
      const name = (p.name ?? p.new_name) as string | undefined;
      if (act === "add_sheet") return { action: "Add", detail: `${name ? `sheet "${name}"` : "sheet"}${recovery}` };
      if (act === "rename_sheet") return { action: "Rename", detail: `${name ? `to "${name}"` : "sheet"}${recovery}` };
      if (act === "delete_sheet") return { action: "Delete", detail: `sheet${recovery}` };
      return { action: "Modify", detail: `structure${recovery}` };
    }
    case "search_workbook": {
      if (resultText) { const s = resultSummary(resultText); if (s) return splitFirstWord(s); }
      const q = p.query as string | undefined;
      return { action: "Search", detail: q ? `"${q}"` : "workbook" };
    }

    // ── Other tools ──
    case "trace_dependencies": {
      const cell = (p.cell ?? p.range) as string | undefined;
      const mode = p.mode === "dependents" ? "dependents" : "precedents";
      return {
        action: mode === "dependents" ? "Trace dependents" : "Trace precedents",
        detail: cell ?? mode,
        address: cell,
      };
    }
    case "explain_formula": {
      const cell = p.cell as string | undefined;
      return {
        action: "Explain formula",
        detail: cell ?? "cell",
        address: cell,
      };
    }
    case "comments": {
      const op = p.action as string | undefined;
      const addr = range ? compactRange(range) : "range";
      const recovery = recoveryBadgeForDetails(details);

      switch (op) {
        case "read":
          return { action: "Comments", detail: addr, address: range };
        case "add":
          return { action: "Add", detail: `comment ${addr}${recovery}`, address: range };
        case "update":
          return { action: "Update", detail: `comment ${addr}${recovery}`, address: range };
        case "reply":
          return { action: "Reply", detail: `${addr}${recovery}`, address: range };
        case "delete":
          return { action: "Delete", detail: `comment ${addr}${recovery}`, address: range };
        case "resolve":
          return { action: "Resolve", detail: `${addr}${recovery}`, address: range };
        case "reopen":
          return { action: "Reopen", detail: `${addr}${recovery}`, address: range };
        default:
          return { action: "Comment", detail: `${addr}${recovery}`, address: range };
      }
    }
    case "view_settings": {
      const op = p.action as string | undefined;
      const targetSheet = p.sheet as string | undefined;
      const targetSheetLabel = targetSheet ?? "active sheet";
      const targetRange = p.range as string | undefined;
      const detailsAddress = isViewSettingsDetails(details) ? details.address : undefined;
      const qualifiedRange = detailsAddress ?? qualifyRangeAddress(targetRange, targetSheet);
      const recovery = recoveryBadgeForDetails(details);

      if (!op || op === "get") {
        return { action: "View", detail: "settings" };
      }

      if (op === "activate") {
        return { action: "Activate", detail: `${targetSheetLabel}${recovery}` };
      }

      if (op === "freeze_at") {
        const freezeTarget = qualifiedRange ?? targetSheetLabel;
        return {
          action: "Freeze",
          detail: `${compactRange(freezeTarget)}${recovery}`,
          address: qualifiedRange,
        };
      }

      if (op.startsWith("hide_") || op.startsWith("show_")) {
        return {
          action: op.startsWith("hide_") ? "Hide" : "Show",
          detail: `${op.replace(/^(hide_|show_)/u, "").replace(/_/gu, " ")} (${targetSheetLabel})${recovery}`,
        };
      }

      return { action: "Set", detail: `${op.replace(/_/gu, " ")} (${targetSheetLabel})${recovery}` };
    }
    case "instructions": {
      const level = p.level as string | undefined;
      const action = p.action as string | undefined;
      const scope = level ? `${level} rules` : "rules";
      if (action === "replace") {
        return { action: "Set", detail: scope };
      }
      if (action === "append") {
        return { action: "Remember", detail: scope };
      }
      return { action: "Update", detail: scope };
    }
    case "conventions": {
      const action = p.action as string | undefined;
      if (action === "get") return { action: "View", detail: "conventions" };
      if (action === "reset") return { action: "Reset", detail: "conventions" };
      return { action: "Update", detail: "conventions" };
    }
    case "workbook_history": {
      const action = p.action as string | undefined;
      const snapshotId = p.snapshot_id as string | undefined;
      if (action === "restore") {
        return {
          action: "Restore",
          detail: snapshotId ? `backup ${snapshotId}` : "latest backup",
        };
      }
      if (action === "delete") {
        return {
          action: "Delete",
          detail: snapshotId ? `backup ${snapshotId}` : "latest backup",
        };
      }
      if (action === "clear") {
        return { action: "Clear", detail: "backups" };
      }
      return { action: "List", detail: "backups" };
    }
    case "skills": {
      const action = p.action as string | undefined;
      const name = p.name as string | undefined;
      const refresh = p.refresh === true;

      if (action === "read") {
        const detailName = isSkillsReadDetails(details)
          ? details.skillName
          : name ?? "name";
        const sourceSuffix = isSkillsReadDetails(details)
          ? (details.sourceKind === "external" ? " (external)" : " (bundled)")
          : "";

        if (isSkillsReadDetails(details) && details.cacheHit) {
          return { action: "Read skill", detail: `${detailName}${sourceSuffix} (cached)` };
        }

        if (refresh) {
          return { action: "Refresh skill", detail: `${detailName}${sourceSuffix}` };
        }

        return { action: "Read skill", detail: `${detailName}${sourceSuffix}` };
      }

      return { action: "List skills", detail: "" };
    }
    case "web_search": {
      const query = p.query as string | undefined;
      return { action: "Web search", detail: query ? `\"${query}\"` : "query" };
    }
    case "fetch_page": {
      const url = p.url as string | undefined;
      return { action: "Fetch page", detail: url ?? "url" };
    }
    case "mcp": {
      if (typeof p.tool === "string") {
        return { action: "MCP call", detail: p.tool };
      }
      if (typeof p.connect === "string") {
        return { action: "MCP connect", detail: p.connect };
      }
      if (typeof p.describe === "string") {
        return { action: "MCP describe", detail: p.describe };
      }
      if (typeof p.search === "string") {
        return { action: "MCP search", detail: `\"${p.search}\"` };
      }
      if (typeof p.server === "string") {
        return { action: "MCP list", detail: p.server };
      }
      return { action: "MCP", detail: "status" };
    }
    case "files": {
      const action = p.action as string | undefined;
      const path = p.path as string | undefined;

      if (action === "list") return { action: "Files", detail: "list" };
      if (action === "read") return { action: "Read file", detail: path ?? "path" };
      if (action === "write") return { action: "Write file", detail: path ?? "path" };
      if (action === "delete") return { action: "Delete file", detail: path ?? "path" };
      return { action: "Files", detail: action ?? "action" };
    }
    default: {
      if (resultText) { const s = resultSummary(resultText); if (s) return splitFirstWord(s); }
      return { action: "Tool", detail: "" };
    }
  }
}

/* ── Renderer ───────────────────────────────────────────────── */

function createExcelMarkdownRenderer(toolName: SupportedToolName): ToolRenderer<unknown, unknown> {
  return {
    render(
      params: unknown,
      result: ToolResultMessage<unknown> | undefined,
      isStreaming?: boolean,
    ): ToolRenderResult {
      const state: ToolState = result
        ? (result.isError ? "error" : "complete")
        : isStreaming
          ? "inprogress"
          : "complete";

      const paramsJson = formatParamsJson(params);
      const contentRef = createRef<HTMLDivElement>();
      const chevronRef = createRef<HTMLElement>();

      // Always start collapsed — the description tells the user what happened
      const defaultExpanded = false;

      const resultText = result ? splitToolResultContent(result).text : undefined;
      const desc = describeToolCall(toolName, params, resultText, result?.details);
      const detailContent = desc.address
        ? (desc.detail && desc.detail !== desc.address
          ? cellRefDisplay(desc.detail, desc.address)
          : cellRefs(desc.address))
        : desc.detail;
      const title = html`<span class="pi-tool-card__title"><strong>${desc.action}</strong>${desc.detail ? html` <span class="pi-tool-card__detail-text">${detailContent}</span>` : ""}</span>`;

      // ── With result ─────────────────────────────────────
      if (result) {
        const { text, images } = splitToolResultContent(result);
        const standaloneImagePath = detectStandaloneImagePath(text);
        const json = tryFormatJsonOutput(text);
        const cleanedText = stripYamlFrontmatter(text);
        const humanizedText = compactRangesInMarkdown(humanizeColorsInText(cleanedText));
        const useMarkdown = !json.isJson && looksLikeMarkdown(cleanedText);
        const csvTable = isReadRangeCsvDetails(result.details)
          ? renderCsvTable(result.details)
          : null;
        const traceDetails = isTraceDependenciesDetails(result.details)
          ? result.details
          : null;
        const depTree = traceDetails
          ? renderDepTree(traceDetails.root, traceDetails.mode ?? "precedents")
          : null;
        const depSectionLabel = traceDetails
          ? traceDetails.mode === "dependents"
            ? "Dependents"
            : "Precedents"
          : "Dependencies";
        const formulaExplanation = renderExplainFormulaDetails(result.details);

        return {
          content: html`
            <div class="pi-tool-card" data-state=${state} data-tool-name=${toolName}>
              <div class="pi-tool-card__header">
                ${renderCollapsibleToolCardHeader(state, title, contentRef, chevronRef, defaultExpanded)}
              </div>
              <div ${ref(contentRef)}
                class="pi-tool-card__body overflow-hidden transition-all duration-300 max-h-0"
              >
                <div class="pi-tool-card__inner">
                  <div class="pi-tool-card__detail">
                    <span class="pi-tool-card__tool-id">${toolName}</span>
                  </div>
                  ${paramsJson ? html`
                    <div class="pi-tool-card__section">
                      <div class="pi-tool-card__section-label">Input</div>
                      ${humanizeToolInput(toolName, params) ?? html`<code-block .code=${paramsJson} language="json"></code-block>`}
                    </div>
                  ` : ""}
                  <div class="pi-tool-card__section">
                    <div class="pi-tool-card__section-label">${formulaExplanation !== null ? "Formula explanation" : depTree !== null ? depSectionLabel : csvTable !== null ? "Data" : "Result"}</div>
                    ${formulaExplanation !== null
                      ? formulaExplanation
                      : csvTable !== null
                      ? csvTable
                      : depTree !== null
                      ? depTree
                      : isEchoResult(toolName, text)
                      ? html`<div class="pi-tool-card__plain-text pi-tool-card__echo-result">✓ Done</div>`
                      : standaloneImagePath
                      ? html`
                        <div class="text-sm">
                          <div>Image:
                            <a href=${toFileUrl(standaloneImagePath)} target="_blank"
                              rel="noopener noreferrer" class="underline">
                              ${pathBasename(standaloneImagePath)}
                            </a>
                          </div>
                          <div class="mt-1 text-xs font-mono text-muted-foreground break-all">
                            ${standaloneImagePath}
                          </div>
                        </div>
                      `
                      : json.isJson
                        ? html`<code-block .code=${json.formatted} language="json"></code-block>`
                        : useMarkdown
                          ? html`<div class="pi-tool-card__markdown"><markdown-block .content=${humanizedText || "(no output)"}></markdown-block></div>`
                          : html`<div class="pi-tool-card__plain-text">${humanizedText || "(no output)"}</div>`}
                    ${renderImages(images)}
                  </div>
                  ${renderWorkbookCellDiff(result.details)}
                  ${renderChangeExplanationSection(toolName, params, text, result.details)}
                </div>
              </div>
            </div>
          `,
          isCustom: true,
        };
      }

      // ── Streaming / pending with params ──────────────────
      if (paramsJson) {
        return {
          content: html`
            <div class="pi-tool-card" data-state=${state} data-tool-name=${toolName}>
              <div class="pi-tool-card__header">
                ${renderCollapsibleToolCardHeader(state, title, contentRef, chevronRef, defaultExpanded)}
              </div>
              <div ${ref(contentRef)}
                class="pi-tool-card__body overflow-hidden transition-all duration-300 max-h-0"
              >
                <div class="pi-tool-card__inner">
                  <div class="pi-tool-card__detail">
                    <span class="pi-tool-card__tool-id">${toolName}</span>
                  </div>
                  <div class="pi-tool-card__section">
                    <div class="pi-tool-card__section-label">Input</div>
                    ${humanizeToolInput(toolName, params) ?? html`<code-block .code=${paramsJson} language="json"></code-block>`}
                  </div>
                </div>
              </div>
            </div>
          `,
          isCustom: true,
        };
      }

      // ── No params or result yet ──────────────────────────
      return {
        content: html`
          <div class="pi-tool-card" data-state=${state} data-tool-name=${toolName}>
            <div class="pi-tool-card__header">
              ${renderToolCardHeader(state, title)}
            </div>
          </div>
        `,
        isCustom: true,
      };
    },
  };
}

const CUSTOM_RENDERED_TOOL_NAMES: readonly SupportedToolName[] = TOOL_NAMES_WITH_RENDERER;

for (const name of CUSTOM_RENDERED_TOOL_NAMES) {
  registerToolRenderer(name, createExcelMarkdownRenderer(name));
}
