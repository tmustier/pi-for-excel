/**
 * Tool renderers for Pi-for-Excel.
 *
 * Renders Excel tool calls as compact, collapsed-by-default cards with
 * human-readable descriptions. Expand to see raw Input/Output.
 */

import type { ImageContent, TextContent, ToolResultMessage } from "@earendil-works/pi-ai";
import { registerToolRenderer } from "./messages/tool-renderer-registry.js";
import type { ToolRenderer, ToolRenderResult } from "./messages/tool-renderer-registry.js";
import { t } from "../language/index.js";
import { html, type TemplateResult } from "lit";
import { ref } from "lit/directives/ref.js";
import { renderDisclosureToolCardHeader, renderToolCardHeader, type ToolCardDisclosure } from "./tool-card-header.js";
import { cellRef, cellRefDisplay, cellRefs } from "./cell-link.js";
import { humanizeToolInput } from "./humanize-params.js";
import { humanizeColorsInText } from "./color-names.js";
import { stripYamlFrontmatter } from "./markdown-preprocess.js";
import { TOOL_NAMES_WITH_RENDERER, type UiToolName } from "../tools/capabilities.js";
import {
  mountSearchSetupCard,
  shouldShowSearchSetupCard,
} from "./web-search-setup-card.js";
import {
  mountBridgeSetupCard,
  shouldShowBridgeSetupCard,
} from "./bridge-setup-card.js";
import {
  decodeToolDetails,
  type ExcelToolDetails,
  type RecoveryCheckpointDetails,
  type WriteCellsDetails,
} from "../tools/tool-details.js";
import { getToolExecutionMode } from "../tools/execution-policy.js";
import {
  buildChangeExplanation,
  type ChangeExplanationInput,
} from "../audit/change-explanation.js";
import { computeRangeAddress, parseRangeRef } from "../excel/helpers.js";
import { renderCsvTable } from "./render-csv-table.js";
import { renderDepTree } from "./render-dep-tree.js";

// The taskpane registers <markdown-block> and <code-block> through register-components.

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

export function splitToolResultContent(result: ToolResultMessage): {
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

export function extractSummaryLine(text: string): string | null {
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
    const drive = win[1];
    const restPath = win[2];
    if (drive && restPath !== undefined) {
      const rest = restPath
        .split("\\")
        .map((seg) => encodeURIComponent(seg))
        .join("/");
      return `file:///${drive.toUpperCase()}:/${rest}`;
    }
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
    <div class="pi-tool-images">
      ${images.map((img) => {
        const src = `data:${img.mimeType};base64,${img.data}`;
        return html`
          <div class="pi-tool-image-frame">
            <img src=${src} alt=${t("tool-card.imageAlt")} class="pi-tool-image" />
          </div>
        `;
      })}
    </div>
  `;
}

type ToolDetails = ExcelToolDetails | undefined;

function getWorkbookCellChanges(details: ToolDetails): WriteCellsDetails["changes"] | undefined {
  switch (details?.kind) {
    case "write_cells":
    case "fill_formula":
    case "python_transform_range":
      return details.changes;
    default:
      return undefined;
  }
}

function formatDiffValue(value: string): string {
  return value.length > 0 ? value : "∅";
}

function renderWorkbookCellDiff(details: ToolDetails): TemplateResult {
  const changes = getWorkbookCellChanges(details);
  if (!changes || changes.changedCount <= 0) return html``;

  return html`
    <div class="pi-tool-card__section">
      <div class="pi-tool-card__section-label">${t("tool-card.changes", { count: changes.changedCount })}</div>
      <div class="pi-tool-card__diff">
        <table class="pi-tool-card__diff-table">
          <thead>
            <tr>
              <th>${t("tool-card.diff.cell")}</th>
              <th>${t("tool-card.diff.before")}</th>
              <th>${t("tool-card.diff.after")}</th>
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
          ? html`<div class="pi-tool-card__diff-note">${t("tool-card.diff.sample", { count: changes.sample.length })}</div>`
          : html``}
      </div>
    </div>
  `;
}

function renderChartImageDetails(details: ToolDetails, hasImageContent: boolean): TemplateResult {
  if (hasImageContent || details?.kind !== "charts" || !details.image) return html``;

  const src = `data:${details.image.mimeType};base64,${details.image.base64}`;
  const alt = details.name ? `Chart ${details.name}` : "Chart image";

  return html`
    <div class="pi-tool-image-frame">
      <img src=${src} alt=${alt} class="pi-tool-image" />
    </div>
  `;
}

function renderExplainFormulaDetails(details: ToolDetails): TemplateResult | null {
  if (details?.kind !== "explain_formula") return null;

  if (!details.hasFormula) {
    return html`<div class="pi-tool-card__plain-text">${details.explanation}</div>`;
  }

  return html`
    <div class="pi-formula-explain">
      <div class="pi-tool-card__plain-text">${details.explanation}</div>
      <div class="pi-formula-explain__formula">
        <span class="pi-formula-explain__label">${t("tool-card.formula.label")}</span>
        <code>${details.formula}</code>
      </div>
      <div class="pi-formula-explain__refs">
        <div class="pi-formula-explain__label">${t("tool-card.formula.references", { count: details.references.length })}</div>
        <ul class="pi-formula-explain__list">
          ${details.references.length > 0
            ? details.references.map((reference) => html`
              <li>
                ${cellRefs(reference.address)}
                ${reference.valuePreview ? html`<span class="pi-formula-explain__preview"> → ${reference.valuePreview}</span>` : html``}
              </li>
            `)
            : html`<li>${t("tool-card.formula.noReferences")}</li>`}
        </ul>
      </div>
      ${details.truncated
        ? html`<div class="pi-tool-card__diff-note">${t("tool-card.formula.sample", { count: details.references.length })}</div>`
        : html``}
    </div>
  `;
}

function optionalString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim().length > 0 ? value : undefined;
}

function createChangeExplanationInput(args: {
  toolName: SupportedToolName;
  blocked: boolean;
  changedCount?: number | undefined;
  summary?: string | undefined;
  error?: string | undefined;
  inputAddress?: string | undefined;
  outputAddress?: string | undefined;
  changes?: ChangeExplanationInput["changes"] | undefined;
}): ChangeExplanationInput {
  return {
    toolName: args.toolName,
    blocked: args.blocked,
    ...(args.changedCount !== undefined ? { changedCount: args.changedCount } : {}),
    ...(args.summary !== undefined ? { summary: args.summary } : {}),
    ...(args.error !== undefined ? { error: args.error } : {}),
    ...(args.inputAddress !== undefined ? { inputAddress: args.inputAddress } : {}),
    ...(args.outputAddress !== undefined ? { outputAddress: args.outputAddress } : {}),
    ...(args.changes !== undefined ? { changes: args.changes } : {}),
  };
}

export function extractResultError(resultText: string | undefined): string | undefined {
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
  details: ToolDetails,
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

  if (details?.kind === "write_cells") {
    return createChangeExplanationInput({
      toolName,
      blocked: details.blocked,
      changedCount: details.changes?.changedCount,
      summary,
      error,
      outputAddress: details.address ?? startCell,
      changes: details.changes,
    });
  }

  if (details?.kind === "fill_formula") {
    return createChangeExplanationInput({
      toolName,
      blocked: details.blocked,
      changedCount: details.changes?.changedCount,
      summary,
      error,
      outputAddress: details.address ?? range,
      changes: details.changes,
    });
  }

  if (details?.kind === "python_transform_range") {
    return createChangeExplanationInput({
      toolName,
      blocked: details.blocked || blockedFromText,
      changedCount: details.changes?.changedCount,
      summary,
      error: details.error ?? error,
      inputAddress: details.inputAddress,
      outputAddress: details.outputAddress,
      changes: details.changes,
    });
  }

  if (details?.kind === "workbook_history" && details.action === "restore") {
    const historyError = optionalString(details.error);
    return createChangeExplanationInput({
      toolName,
      blocked: blockedFromText || Boolean(historyError),
      changedCount: details.changedCount,
      summary,
      error: historyError ?? error,
      outputAddress: details.address,
    });
  }

  if (toolName === "format_cells") {
    const detailsAddress = details?.kind === "format_cells" ? details.address : undefined;
    return createChangeExplanationInput({
      toolName,
      blocked: blockedFromText,
      summary,
      error,
      outputAddress: detailsAddress ?? range,
    });
  }

  if (toolName === "conditional_format") {
    return createChangeExplanationInput({
      toolName,
      blocked: blockedFromText,
      summary,
      error,
      outputAddress: range,
    });
  }

  if (toolName === "modify_structure") {
    return createChangeExplanationInput({
      toolName,
      blocked: blockedFromText,
      summary,
      error,
      outputAddress: range ?? sheet,
    });
  }

  if (toolName === "comments") {
    return createChangeExplanationInput({
      toolName,
      blocked: blockedFromText,
      changedCount: blockedFromText ? 0 : 1,
      summary,
      error,
      outputAddress: range,
    });
  }

  if (toolName === "view_settings") {
    const outputAddress = range ?? sheet;
    return createChangeExplanationInput({
      toolName,
      blocked: blockedFromText,
      changedCount: blockedFromText ? 0 : 1,
      summary,
      error,
      outputAddress,
    });
  }

  if (toolName === "charts") {
    const detailsAddress = details?.kind === "charts" ? details.address : undefined;
    const detailsSource = details?.kind === "charts" ? details.sourceRange : undefined;
    return createChangeExplanationInput({
      toolName,
      blocked: blockedFromText,
      changedCount: blockedFromText ? 0 : 1,
      summary,
      error,
      outputAddress: detailsAddress ?? detailsSource ?? range,
    });
  }

  if (toolName === "workbook_history" && action === "restore") {
    return createChangeExplanationInput({
      toolName,
      blocked: blockedFromText,
      summary,
      error,
    });
  }

  return null;
}

function renderCitations(citations: readonly string[]): TemplateResult {
  if (citations.length === 0) {
    return html`<span class="pi-tool-card__explain-citations-empty">${t("tool-card.explain.noCitations")}</span>`;
  }

  return html`${citations.map((address, index) => html`${index > 0 ? html`, ` : html``}${cellRefs(address)}`)}`;
}

function renderChangeExplanationSection(
  toolName: SupportedToolName,
  params: unknown,
  resultText: string | undefined,
  details: ToolDetails,
): TemplateResult {
  const input = buildChangeExplanationInputForTool(toolName, params, resultText, details);
  if (!input) return html``;

  const explanation = buildChangeExplanation(input);

  return html`
    <div class="pi-tool-card__section">
      <details class="pi-tool-card__explain">
        <summary class="pi-tool-card__explain-toggle">${t("tool-card.explain.toggle")}</summary>
        <div class="pi-tool-card__explain-body">
          <div class="pi-tool-card__plain-text">${explanation.text}</div>
          <div class="pi-tool-card__explain-citations">
            <span class="pi-tool-card__explain-citations-label">${t("tool-card.explain.citations")}</span>
            ${renderCitations(explanation.citations)}
          </div>
          ${explanation.usedFallback
            ? html`<div class="pi-tool-card__diff-note">${t("tool-card.explain.fallback")}</div>`
            : html``}
          ${explanation.truncated
            ? html`<div class="pi-tool-card__diff-note">${t("tool-card.explain.truncated")}</div>`
            : html``}
        </div>
      </details>
    </div>
  `;
}

/* ── Human-readable descriptions ────────────────────────────── */

/** Strip "(N×M)" / "(NxM)" dimension notation — not intuitive for users. */
export function stripDimensions(text: string): string {
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

  const first = parsed[0]?.sheet;
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
  return m?.[1] ?? null;
}

/** Count formula errors mentioned in tool result text. */
function countResultErrors(text: string): number {
  const m = /(\d+)\s+formula error/i.exec(text);
  return m?.[1] ? parseInt(m[1], 10) : 0;
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

function recoveryBadgeForDetails(details: ToolDetails): string {
  switch (details?.kind) {
    case "format_cells":
    case "conditional_format":
    case "modify_structure":
    case "comments":
    case "view_settings":
    case "charts":
      return withRecoveryBadge("", details.recovery);
    default:
      return "";
  }
}

/** Append error / blocked badge to the detail string. */
function badge(
  toolName: SupportedToolName,
  resultText: string | undefined,
  details: ToolDetails,
): string {
  if (toolName === "write_cells" && details?.kind === "write_cells") {
    if (details.blocked) return " — blocked";
    return withRecoveryBadge(
      mutationBadge(details.changes?.changedCount, details.formulaErrorCount),
      details.recovery,
    );
  }

  if (toolName === "fill_formula" && details?.kind === "fill_formula") {
    if (details.blocked) return " — blocked";
    return withRecoveryBadge(
      mutationBadge(details.changes?.changedCount, details.formulaErrorCount),
      details.recovery,
    );
  }

  if (toolName === "python_transform_range" && details?.kind === "python_transform_range") {
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

export interface ToolDesc {
  /** Bold verb, e.g. "Read", "Wrote", "Format" */
  action: string;
  /** Normal-weight rest, e.g. "Costs!A1:C19" */
  detail: string;
  /** Raw Excel address for click-to-navigate. Omit for non-range details. */
  address?: string;
}

function toolDescWithAddress(action: string, detail: string, address: string | undefined): ToolDesc {
  return address !== undefined ? { action, detail, address } : { action, detail };
}

/** Split a result-text summary line into action (first word) + rest. */
function splitFirstWord(text: string): ToolDesc {
  const i = text.indexOf(" ");
  return i > 0
    ? { action: text.substring(0, i), detail: text.substring(i + 1) }
    : { action: text, detail: "" };
}

interface ToolDescriptionContext {
  params: Record<string, unknown>;
  range: string | undefined;
  startCell: string | undefined;
  resultText: string | undefined;
  details: ToolDetails;
}

type ToolDescriptionHandler = (context: ToolDescriptionContext) => ToolDesc;

/**
 * Header for a write the user stopped (`details.stopped`). How far it got:
 * write_cells / fill_formula report `writtenCount` (0 = nothing changed);
 * python_transform_range reports `rowsWritten × colsWritten`, and no
 * `rowsWritten` when Stop came before the destination was known. A stopped
 * write_cells / fill_formula without `writtenCount` (older results) cannot
 * say how much landed, so the header asks the user to check the range.
 */
function describeStoppedWrite(details: ToolDetails): ToolDesc | null {
  let written: number | undefined;
  let address: string | undefined;
  switch (details?.kind) {
    case "write_cells":
    case "fill_formula":
      if (details.stopped !== true) return null;
      written = details.writtenCount;
      address = details.address;
      break;
    case "python_transform_range":
      if (details.stopped !== true) return null;
      written = details.rowsWritten === undefined ? 0 : details.rowsWritten * (details.colsWritten ?? 1);
      address = details.outputAddress;
      break;
    default:
      return null;
  }

  const action = t("tool-card.stopped.action");
  if (written === 0) return { action, detail: t("tool-card.stopped.nothing") };
  if (written === undefined) {
    return address
      ? toolDescWithAddress(action, t("tool-card.stopped.unknown", { address }), address)
      : { action, detail: t("tool-card.stopped.nothing") };
  }
  return address
    ? toolDescWithAddress(action, t("tool-card.stopped.after", { count: written, address }), address)
    : { action, detail: t("tool-card.stopped.afterCount", { count: written }) };
}

function describeWriteCells(context: ToolDescriptionContext): ToolDesc {
  const { details, resultText, startCell } = context;
  const stopped = describeStoppedWrite(details);
  if (stopped) return stopped;
  const receipt = badge("write_cells", resultText, details);
  if (details?.kind === "write_cells" && details.address) {
    return toolDescWithAddress(
      details.blocked ? "Write" : "Edit",
      details.address + receipt,
      details.address,
    );
  }

  const address = resultText ? extractWrittenAddress(resultText) : null;
  if (address) return toolDescWithAddress("Edit", address + receipt, address);

  const values = context.params.values;
  const rows = Array.isArray(values) ? values.length : 0;
  const cols = Array.isArray(values)
    ? Math.max(0, ...values.map((row) => (Array.isArray(row) ? row.length : 0)))
    : 0;
  const target = startCell ? pendingWriteTarget(startCell, rows, cols) : undefined;
  return toolDescWithAddress("Write", (target ?? "cells") + receipt, target);
}

/** The whole range a `write_cells` call will fill ("Sheet1!B2:C4"), before it runs. */
function pendingWriteTarget(startCell: string, rows: number, cols: number): string {
  if (rows <= 0 || cols <= 0 || (rows === 1 && cols === 1)) return startCell;

  const { sheet, address } = parseRangeRef(startCell);
  try {
    const range = computeRangeAddress(address, rows, cols);
    return sheet ? `${startCell.slice(0, startCell.lastIndexOf("!") + 1)}${range}` : range;
  } catch {
    return startCell;
  }
}

function describeFillFormula(context: ToolDescriptionContext): ToolDesc {
  const { details, range, resultText } = context;
  const stopped = describeStoppedWrite(details);
  if (stopped) return stopped;
  const receipt = badge("fill_formula", resultText, details);
  if (details?.kind === "fill_formula" && details.address) {
    return toolDescWithAddress(
      details.blocked ? "Fill" : "Filled",
      details.address + receipt,
      details.address,
    );
  }

  const address = resultText ? extractWrittenAddress(resultText) : null;
  return address
    ? toolDescWithAddress("Filled", address + receipt, address)
    : toolDescWithAddress("Fill", (range ? compactRange(range) : "formula") + receipt, range);
}

function describePythonTransform(context: ToolDescriptionContext): ToolDesc {
  const { details, params, range, resultText } = context;
  const stopped = describeStoppedWrite(details);
  if (stopped) return stopped;
  const receipt = badge("python_transform_range", resultText, details);
  if (details?.kind === "python_transform_range") {
    const address = details.outputAddress ?? details.inputAddress;
    if (address) {
      const hasError = typeof details.error === "string" && details.error.length > 0;
      return toolDescWithAddress(
        details.blocked || hasError ? "Transform" : "Transformed",
        address + receipt,
        address,
      );
    }
  }

  const outputStart = params.output_start_cell as string | undefined;
  const fallbackAddress = outputStart ?? range;
  return toolDescWithAddress("Transform", (fallbackAddress ?? "range") + receipt, fallbackAddress);
}

function describeModifyStructure(context: ToolDescriptionContext): ToolDesc {
  const { details, params, resultText } = context;
  const recovery = recoveryBadgeForDetails(details);
  const summary = resultText ? resultSummary(resultText) : null;
  if (summary) {
    const parts = splitFirstWord(summary);
    return { ...parts, detail: `${parts.detail}${recovery}` };
  }

  const action = params.action as string | undefined;
  const name = (params.name ?? params.new_name) as string | undefined;
  const sheet = params.sheet as string | undefined;
  if (action === "add_sheet") return { action: "Add", detail: `${name ? `sheet "${name}"` : "sheet"}${recovery}` };
  if (action === "rename_sheet") return { action: "Rename", detail: `${name ? `to "${name}"` : "sheet"}${recovery}` };
  if (action === "delete_sheet") return { action: "Delete", detail: `sheet${sheet ? ` "${sheet}"` : ""}${recovery}` };
  if (action === "insert_rows" || action === "delete_rows" || action === "insert_columns" || action === "delete_columns") {
    const count = typeof params.count === "number" ? params.count : 1;
    const position = typeof params.position === "number" ? params.position : undefined;
    const unit = action.endsWith("rows") ? "row" : "column";
    const where = position === undefined ? ""
      : count > 1 ? ` ${position}–${position + count - 1}` : ` ${position}`;
    return {
      action: action.startsWith("insert") ? "Insert" : "Delete",
      detail: `${count > 1 ? `${count} ${unit}s` : unit}${where}${sheet ? ` in ${sheet}` : ""}${recovery}`,
    };
  }
  return { action: "Modify", detail: `structure${recovery}` };
}

function describeCharts(context: ToolDescriptionContext): ToolDesc {
  const { details, params } = context;
  const operation = params.action as string | undefined;
  const recovery = recoveryBadgeForDetails(details);
  const chartDetails = details?.kind === "charts" ? details : undefined;
  const chartName = chartDetails?.name ??
    (params.name as string | undefined) ??
    (params.new_name as string | undefined);
  const source = chartDetails?.sourceRange ?? (params.source_range as string | undefined);

  if (operation === "list") {
    const count = chartDetails?.count !== undefined ? ` (${chartDetails.count})` : "";
    return { action: "Charts", detail: `${(params.sheet as string | undefined) ?? "workbook"}${count}` };
  }
  if (operation === "create" || operation === "update") {
    const status = operation === "create" ? "created" : "updated";
    const fallback = operation === "create" ? "new" : "chart";
    return toolDescWithAddress(
      "Chart",
      `'${chartName ?? fallback}' ${status}${source ? ` — ${compactRange(source)}` : ""}${recovery}`,
      source,
    );
  }
  if (operation === "delete") {
    return { action: "Chart", detail: `'${chartName ?? "chart"}' deleted${recovery}` };
  }
  if (operation === "get_image") {
    const size = chartDetails?.image
      ? ` (${chartDetails.image.width}×${chartDetails.image.height}px)`
      : "";
    return { action: "Chart image", detail: `'${chartName ?? "chart"}'${size}` };
  }
  return { action: "Charts", detail: chartName ?? "chart" };
}

function describeComments(context: ToolDescriptionContext): ToolDesc {
  const { details, params, range } = context;
  const operation = params.action as string | undefined;
  const address = range ? compactRange(range) : "range";
  const recovery = recoveryBadgeForDetails(details);
  switch (operation) {
    case "read": return toolDescWithAddress("Comments", address, range);
    case "add": return toolDescWithAddress("Add", `comment ${address}${recovery}`, range);
    case "update": return toolDescWithAddress("Update", `comment ${address}${recovery}`, range);
    case "reply": return toolDescWithAddress("Reply", `${address}${recovery}`, range);
    case "delete": return toolDescWithAddress("Delete", `comment ${address}${recovery}`, range);
    case "resolve": return toolDescWithAddress("Resolve", `${address}${recovery}`, range);
    case "reopen": return toolDescWithAddress("Reopen", `${address}${recovery}`, range);
    default: return toolDescWithAddress("Comment", `${address}${recovery}`, range);
  }
}

function describeViewSettings(context: ToolDescriptionContext): ToolDesc {
  const { details, params } = context;
  const operation = params.action as string | undefined;
  const targetSheet = params.sheet as string | undefined;
  const targetSheetLabel = targetSheet ?? "active sheet";
  const targetRange = params.range as string | undefined;
  const detailsAddress = details?.kind === "view_settings" ? details.address : undefined;
  const qualifiedRange = detailsAddress ?? qualifyRangeAddress(targetRange, targetSheet);
  const recovery = recoveryBadgeForDetails(details);

  if (!operation || operation === "get") return { action: "View", detail: "settings" };
  if (operation === "activate") return { action: "Activate", detail: `${targetSheetLabel}${recovery}` };
  if (operation === "freeze_at") {
    const freezeTarget = qualifiedRange ?? targetSheetLabel;
    return toolDescWithAddress("Freeze", `${compactRange(freezeTarget)}${recovery}`, qualifiedRange);
  }
  if (operation.startsWith("hide_") || operation.startsWith("show_")) {
    return {
      action: operation.startsWith("hide_") ? "Hide" : "Show",
      detail: `${operation.replace(/^(hide_|show_)/u, "").replace(/_/gu, " ")} (${targetSheetLabel})${recovery}`,
    };
  }
  return { action: "Set", detail: `${operation.replace(/_/gu, " ")} (${targetSheetLabel})${recovery}` };
}

function describeInstructions(context: ToolDescriptionContext): ToolDesc {
  const level = context.params.level as string | undefined;
  const action = context.params.action as string | undefined;
  const scope = level ? `${level} rules` : "rules";
  if (action === "replace") return { action: "Set", detail: scope };
  if (action === "append") return { action: "Remember", detail: scope };
  return { action: "Update", detail: scope };
}

function describeWorkbookHistory(context: ToolDescriptionContext): ToolDesc {
  const action = context.params.action as string | undefined;
  const snapshotId = context.params.snapshot_id as string | undefined;
  if (action === "restore" || action === "delete") {
    return {
      action: action === "restore" ? "Restore" : "Delete",
      detail: snapshotId ? `backup ${snapshotId}` : "latest backup",
    };
  }
  if (action === "clear") return { action: "Clear", detail: "backups" };
  return { action: "List", detail: "backups" };
}

function describeSkills(context: ToolDescriptionContext): ToolDesc {
  const { details, params } = context;
  const action = params.action as string | undefined;
  const name = params.name as string | undefined;
  if (action === "read") {
    const readDetails = details?.kind === "skills_read" ? details : undefined;
    const detailName = readDetails?.skillName ?? name ?? "name";
    const sourceSuffix = readDetails
      ? (readDetails.sourceKind === "external" ? " (external)" : " (bundled)")
      : "";
    if (readDetails?.cacheHit) return { action: "Read skill", detail: `${detailName}${sourceSuffix} (cached)` };
    if (params.refresh === true) return { action: "Refresh skill", detail: `${detailName}${sourceSuffix}` };
    return { action: "Read skill", detail: `${detailName}${sourceSuffix}` };
  }
  if (action === "install") {
    const installedName = details?.kind === "skills_install" ? details.skillName : name ?? "skill";
    return { action: "Install skill", detail: installedName };
  }
  if (action === "uninstall") {
    const uninstallDetails = details?.kind === "skills_uninstall" ? details : undefined;
    const removedName = uninstallDetails?.skillName ?? name ?? "skill";
    const removedSuffix = uninstallDetails ? (uninstallDetails.removed ? "" : " (not found)") : "";
    return { action: "Uninstall skill", detail: `${removedName}${removedSuffix}` };
  }
  return { action: "List skills", detail: "" };
}

function describeMcp(context: ToolDescriptionContext): ToolDesc {
  const { params } = context;
  if (typeof params.tool === "string") return { action: "MCP call", detail: params.tool };
  if (typeof params.connect === "string") return { action: "MCP connect", detail: params.connect };
  if (typeof params.describe === "string") return { action: "MCP describe", detail: params.describe };
  if (typeof params.search === "string") return { action: "MCP search", detail: `"${params.search}"` };
  if (typeof params.server === "string") return { action: "MCP list", detail: params.server };
  return { action: "MCP", detail: "status" };
}

const TOOL_DESCRIPTION_HANDLERS = {
  read_range: ({ params, range }) => {
    const mode = params.mode as string | undefined;
    const label = mode === "csv" ? "Export" : "Read";
    return toolDescWithAddress(label, range ? compactRange(range) + (mode === "csv" ? " (CSV)" : "") : "range", range);
  },
  get_workbook_overview: ({ params }) => ({
    action: t("tools.action.overview"),
    detail: (params.sheet as string | undefined) ?? "",
  }),
  write_cells: describeWriteCells,
  fill_formula: describeFillFormula,
  python_transform_range: describePythonTransform,
  format_cells: ({ details, range }) => {
    const address = details?.kind === "format_cells" ? details.address : undefined;
    const resolved = address ?? range;
    return toolDescWithAddress(
      "Format",
      (resolved ? compactRange(resolved) : "cells") + recoveryBadgeForDetails(details),
      resolved,
    );
  },
  conditional_format: ({ details, range }) => toolDescWithAddress(
    "Cond. format",
    (range ? compactRange(range) : "cells") + recoveryBadgeForDetails(details),
    range,
  ),
  modify_structure: describeModifyStructure,
  search_workbook: ({ params, resultText }) => {
    const summary = resultText ? resultSummary(resultText) : null;
    if (summary) return splitFirstWord(summary);
    const query = params.query as string | undefined;
    return { action: "Search", detail: query ? `"${query}"` : "workbook" };
  },
  trace_dependencies: ({ params }) => {
    const cell = (params.cell ?? params.range) as string | undefined;
    const mode = params.mode === "dependents" ? "dependents" : "precedents";
    return toolDescWithAddress(
      mode === "dependents" ? "Trace dependents" : "Trace precedents",
      cell ?? mode,
      cell,
    );
  },
  explain_formula: ({ params }) => {
    const cell = params.cell as string | undefined;
    return toolDescWithAddress("Explain formula", cell ?? "cell", cell);
  },
  charts: describeCharts,
  comments: describeComments,
  view_settings: describeViewSettings,
  instructions: describeInstructions,
  conventions: ({ params }) => {
    const action = params.action as string | undefined;
    if (action === "get") return { action: "View", detail: "conventions" };
    if (action === "reset") return { action: "Reset", detail: "conventions" };
    return { action: "Update", detail: "conventions" };
  },
  workbook_history: describeWorkbookHistory,
  skills: describeSkills,
  web_search: ({ params }) => {
    const query = params.query as string | undefined;
    return { action: "Web search", detail: query ? `"${query}"` : "query" };
  },
  fetch_page: ({ params }) => ({ action: "Fetch page", detail: (params.url as string | undefined) ?? "url" }),
  mcp: describeMcp,
  files: ({ params }) => {
    const action = params.action as string | undefined;
    const path = params.path as string | undefined;
    if (action === "list") return { action: "Files", detail: "list" };
    if (action === "read") return { action: "Read file", detail: path ?? "path" };
    if (action === "write") return { action: "Write file", detail: path ?? "path" };
    if (action === "delete") return { action: "Delete file", detail: path ?? "path" };
    return { action: "Files", detail: action ?? "action" };
  },
  extensions_manager: ({ params }) => {
    const action = params.action as string | undefined;
    const target = (params.name ?? params.extension_id) as string | undefined;
    if (action === "install_code") return { action: "Install extension", detail: target ?? "extension" };
    if (action === "set_enabled") {
      return { action: params.enabled === false ? "Disable extension" : "Enable extension", detail: target ?? "extension" };
    }
    if (action === "reload") return { action: "Reload extension", detail: target ?? "extension" };
    if (action === "uninstall") return { action: "Uninstall extension", detail: target ?? "extension" };
    return { action: "Extensions", detail: "list" };
  },
  execute_office_js: ({ params }) => ({
    action: t("tools.action.runOfficeJs"),
    detail: (params.explanation as string | undefined) ?? "code",
  }),
  execute_wps_js: ({ params }) => ({
    action: t("tools.action.runWpsJs"),
    detail: (params.explanation as string | undefined) ?? "code",
  }),
} satisfies Record<SupportedToolName, ToolDescriptionHandler>;

/** Structured description: bold action + normal-weight detail. */
export function describeToolCall(
  toolName: SupportedToolName,
  params: unknown,
  resultText: string | undefined,
  details: ToolDetails,
): ToolDesc {
  const parsedParams = safeParseParams(params);
  return TOOL_DESCRIPTION_HANDLERS[toolName]({
    params: parsedParams,
    range: parsedParams.range as string | undefined,
    startCell: parsedParams.start_cell as string | undefined,
    resultText,
    details,
  });
}

/* ── Renderer ───────────────────────────────────────────────── */

/** The card's failure test: tools report most failures as an "Error …" result rather than throwing. */
export function isFailedToolResult(result: ToolResultMessage): boolean {
  return result.isError || extractResultError(splitToolResultContent(result).text) !== undefined;
}

/**
 * Expanded body of a tool call: the humanized input, then the result
 * (data table, dependency tree, markdown, image, or JSON), the cell diff and
 * the change explanation. Shared by the tool card and the turn group row.
 */
export function renderToolCardBody(
  toolName: SupportedToolName,
  params: unknown,
  result: ToolResultMessage | undefined,
): TemplateResult {
  const paramsJson = formatParamsJson(params);
  const inputSection = paramsJson ? html`
    <div class="pi-tool-card__section">
      <div class="pi-tool-card__section-label">${t("messages.toolInput")}</div>
      ${humanizeToolInput(toolName, params) ?? html`<code-block .code=${paramsJson} language="json"></code-block>`}
    </div>
  ` : html``;

  if (!result) {
    return html`
      <div class="pi-tool-card__inner">
        <div class="pi-tool-card__detail">
          <span class="pi-tool-card__tool-id">${toolName}</span>
        </div>
        ${inputSection}
      </div>
    `;
  }

  // `details` is unknown at this seam: older persisted sessions may carry any shape.
  const details = decodeToolDetails(result.details);
  const { text, images } = splitToolResultContent(result);
  const standaloneImagePath = detectStandaloneImagePath(text);
  const json = tryFormatJsonOutput(text);
  const cleanedText = stripYamlFrontmatter(text);
  const humanizedText = compactRangesInMarkdown(humanizeColorsInText(cleanedText));
  const useMarkdown = !json.isJson && looksLikeMarkdown(cleanedText);
  const csvTable = details?.kind === "read_range_csv"
    ? renderCsvTable(details)
    : null;
  const traceDetails = details?.kind === "trace_dependencies"
    ? details
    : null;
  const depTree = traceDetails
    ? renderDepTree(traceDetails.root, traceDetails.mode ?? "precedents")
    : null;
  const depSectionLabel = traceDetails
    ? traceDetails.mode === "dependents"
      ? t("tool-card.section.dependents")
      : t("tool-card.section.precedents")
    : t("tool-card.section.dependencies");
  const formulaExplanation = renderExplainFormulaDetails(details);

  return html`
    <div class="pi-tool-card__inner">
      <div class="pi-tool-card__detail">
        <span class="pi-tool-card__tool-id">${toolName}</span>
      </div>
      ${inputSection}
      <div class="pi-tool-card__section">
        <div class="pi-tool-card__section-label">${formulaExplanation !== null ? t("tool-card.section.formulaExplanation") : depTree !== null ? depSectionLabel : csvTable !== null ? t("tool-card.section.data") : t("tool-card.section.result")}</div>
        ${formulaExplanation !== null
          ? formulaExplanation
          : csvTable !== null
          ? csvTable
          : depTree !== null
          ? depTree
          : isEchoResult(toolName, text)
          ? html`<div class="pi-tool-card__plain-text pi-tool-card__echo-result">${t("tool-card.done")}</div>`
          : standaloneImagePath
          ? html`
            <div class="pi-tool-image-link">
              <div>${t("tool-card.image")}
                <a href=${toFileUrl(standaloneImagePath)} target="_blank"
                  rel="noopener noreferrer" class="pi-tool-image-link__anchor">
                  ${pathBasename(standaloneImagePath)}
                </a>
              </div>
              <div class="pi-tool-image-link__path">
                ${standaloneImagePath}
              </div>
            </div>
          `
          : json.isJson
            ? html`<code-block .code=${json.formatted} language="json"></code-block>`
            : useMarkdown
              ? html`<div class="pi-tool-card__markdown"><markdown-block .content=${humanizedText || t("messages.noOutput")}></markdown-block></div>`
              : html`<div class="pi-tool-card__plain-text">${humanizedText || t("messages.noOutput")}</div>`}
        ${renderImages(images)}
        ${renderChartImageDetails(details, images.length > 0)}
      </div>
      ${renderWorkbookCellDiff(details)}
      ${renderChangeExplanationSection(toolName, params, text, details)}
    </div>
  `;
}

/**
 * Guided setup for a failure the user can fix (web search not configured,
 * a bridge unreachable). Shown next to the call whether or not it is
 * expanded.
 */
export function renderToolSetupCards(result: ToolResultMessage | undefined): TemplateResult | null {
  const details = decodeToolDetails(result?.details);
  const searchSetupDetails = details !== undefined && shouldShowSearchSetupCard(details) ? details : null;
  const initSearchSetup = (el: Element | undefined): void => {
    if (el instanceof HTMLElement && searchSetupDetails) {
      mountSearchSetupCard(el, searchSetupDetails);
    }
  };
  const bridgeSetupDetails = details !== undefined && shouldShowBridgeSetupCard(details) ? details : null;
  if (!searchSetupDetails && !bridgeSetupDetails) return null;
  const initBridgeSetup = (el: Element | undefined): void => {
    if (el instanceof HTMLElement && bridgeSetupDetails) {
      mountBridgeSetupCard(el, bridgeSetupDetails);
    }
  };
  return html`
    ${searchSetupDetails !== null ? html`<div ${ref(initSearchSetup)}></div>` : html``}
    ${bridgeSetupDetails !== null ? html`<div ${ref(initBridgeSetup)}></div>` : html``}
  `;
}

function createExcelMarkdownRenderer(toolName: SupportedToolName): ToolRenderer {
  return {
    render(
      params: unknown,
      result: ToolResultMessage | undefined,
      isStreaming: boolean,
      disclosure: ToolCardDisclosure,
    ): ToolRenderResult {
      const resultText = result ? splitToolResultContent(result).text : undefined;
      const state: ToolState = result
        ? (isFailedToolResult(result) ? "error" : "complete")
        : isStreaming
          ? "inprogress"
          : "complete";

      const paramsJson = formatParamsJson(params);
      const bodyClass = disclosure.expanded
        ? "pi-tool-card__body"
        : "pi-tool-card__body pi-tool-card__body--collapsed";

      // `details` is unknown at this seam: older persisted sessions may carry any shape.
      const details = decodeToolDetails(result?.details);
      const desc = describeToolCall(toolName, params, resultText, details);
      const detailContent = desc.address
        ? (desc.detail && desc.detail !== desc.address
          ? cellRefDisplay(desc.detail, desc.address)
          : cellRefs(desc.address))
        : desc.detail;
      const title = html`<span class="pi-tool-card__title"><strong>${desc.action}</strong>${desc.detail ? html` <span class="pi-tool-card__detail-text">${detailContent}</span>` : ""}</span>`;

      // ── No params or result yet ──────────────────────────
      if (!result && !paramsJson) {
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
      }

      return {
        content: html`
          <div class="pi-tool-card" data-state=${state} data-tool-name=${toolName}>
            <div class="pi-tool-card__header">
              ${renderDisclosureToolCardHeader(state, title, disclosure)}
            </div>
            <div class=${bodyClass}>${renderToolCardBody(toolName, params, result)}</div>
          </div>
          ${renderToolSetupCards(result)}
        `,
        isCustom: true,
      };
    },
  };
}

const CUSTOM_RENDERED_TOOL_NAMES: readonly SupportedToolName[] = TOOL_NAMES_WITH_RENDERER;
const CUSTOM_RENDERED_TOOL_NAME_SET = new Set<string>(CUSTOM_RENDERED_TOOL_NAMES);

export function isSupportedToolName(name: string): name is SupportedToolName {
  return CUSTOM_RENDERED_TOOL_NAME_SET.has(name);
}

/**
 * One-line description of a tool call that has not run yet, as its tool
 * card shows it (for example "Write Sheet1!B2:C4"). Approvals use it so the
 * user sees the same words they will see on the card.
 */
export function describePendingToolCall(toolName: string, params: unknown): string | null {
  if (!isSupportedToolName(toolName)) return null;
  const desc = describeToolCall(toolName, params, undefined, undefined);
  return desc.detail ? `${desc.action} ${desc.detail}` : desc.action;
}

for (const name of CUSTOM_RENDERED_TOOL_NAMES) {
  registerToolRenderer(name, createExcelMarkdownRenderer(name));
}
