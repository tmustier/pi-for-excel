/**
 * Turn group rows: one sentence per tool call, in the tense of its state.
 *
 * The tool card describers (tool-renderers.ts) know what each call did and
 * where ("Fill", "'Cash Flow'!D10:L10"). A row conjugates that verb for its
 * state (`Filled`, `Filling`, `Could not fill`, `Stopped filling`), keeps
 * the reference in mono, and moves the card's badge ("— 9 changed") into a
 * single meta fact ("9 cells"). Verbs stay in English here, as the card
 * verbs do; the structural strings are localized.
 */

import type { ToolCall, ToolResultMessage } from "@earendil-works/pi-ai";
import { html, nothing, type TemplateResult } from "lit";
import { Code, Globe, Puzzle, Table, Terminal, type IconNode } from "lucide";

import { t } from "../language/index.js";
import { decodeToolDetails, type ExcelToolDetails } from "../tools/tool-details.js";
import { cellRefDisplay, cellRefs } from "./cell-link.js";
import {
  describeToolCall,
  extractResultError,
  extractSummaryLine,
  isSupportedToolName,
  renderToolCardBody,
  splitToolResultContent,
  stripDimensions,
} from "./tool-renderers.js";

/** How a row reads: what the call is doing, did, failed to do, or was stopped doing. */
export type ToolRowTense = "live" | "done" | "failed" | "stopped";

/** The family of tool a row belongs to; the group header stacks one icon per family used. */
export type ToolKind = "sheet" | "python" | "web" | "code" | "extension";

export interface ToolRowDescription {
  sentence: TemplateResult;
  /** One fact about the result ("9 cells", "6 results"), or null. */
  meta: string | null;
  /** The error text the model received, for failed rows. */
  reason: string | null;
}

interface VerbForms {
  /** Infinitive, as a button label and after "Could not": "Fill" */
  base: string;
  past: string;
  live: string;
}

const VERBS: Record<string, VerbForms> = {
  Read: { base: "Read", past: "Read", live: "Reading" },
  Export: { base: "Export", past: "Exported", live: "Exporting" },
  Overview: { base: "Read the overview of", past: "Read the overview of", live: "Reading the overview of" },
  Edit: { base: "Edit", past: "Edited", live: "Editing" },
  Write: { base: "Write", past: "Wrote", live: "Writing" },
  Fill: { base: "Fill", past: "Filled", live: "Filling" },
  Filled: { base: "Fill", past: "Filled", live: "Filling" },
  Transform: { base: "Transform", past: "Transformed", live: "Transforming" },
  Transformed: { base: "Transform", past: "Transformed", live: "Transforming" },
  Format: { base: "Format", past: "Formatted", live: "Formatting" },
  "Cond. format": { base: "Conditionally format", past: "Conditionally formatted", live: "Conditionally formatting" },
  Add: { base: "Add", past: "Added", live: "Adding" },
  Rename: { base: "Rename", past: "Renamed", live: "Renaming" },
  Delete: { base: "Delete", past: "Deleted", live: "Deleting" },
  Insert: { base: "Insert", past: "Inserted", live: "Inserting" },
  Modify: { base: "Modify", past: "Modified", live: "Modifying" },
  // Result summaries already read in the past tense ("Deleted 3 row(s) …").
  Deleted: { base: "Delete", past: "Deleted", live: "Deleting" },
  Inserted: { base: "Insert", past: "Inserted", live: "Inserting" },
  Added: { base: "Add", past: "Added", live: "Adding" },
  Renamed: { base: "Rename", past: "Renamed", live: "Renaming" },
  Duplicated: { base: "Duplicate", past: "Duplicated", live: "Duplicating" },
  Hidden: { base: "Hide", past: "Hid", live: "Hiding" },
  Unhidden: { base: "Unhide", past: "Unhid", live: "Unhiding" },
  Search: { base: "Search", past: "Searched", live: "Searching" },
  "Trace precedents": { base: "Trace precedents of", past: "Traced precedents of", live: "Tracing precedents of" },
  "Trace dependents": { base: "Trace dependents of", past: "Traced dependents of", live: "Tracing dependents of" },
  "Explain formula": { base: "Explain the formula in", past: "Explained the formula in", live: "Explaining the formula in" },
  Chart: { base: "Chart", past: "Chart", live: "Charting" },
  Charts: { base: "List charts in", past: "Listed charts in", live: "Listing charts in" },
  "Chart image": { base: "Capture chart", past: "Captured chart", live: "Capturing chart" },
  Comments: { base: "Read comments on", past: "Read comments on", live: "Reading comments on" },
  Comment: { base: "Comment on", past: "Commented on", live: "Commenting on" },
  Reply: { base: "Reply on", past: "Replied on", live: "Replying on" },
  Resolve: { base: "Resolve", past: "Resolved", live: "Resolving" },
  Reopen: { base: "Reopen", past: "Reopened", live: "Reopening" },
  Update: { base: "Update", past: "Updated", live: "Updating" },
  View: { base: "View", past: "Viewed", live: "Viewing" },
  Activate: { base: "Activate", past: "Activated", live: "Activating" },
  Freeze: { base: "Freeze", past: "Froze", live: "Freezing" },
  Hide: { base: "Hide", past: "Hid", live: "Hiding" },
  Show: { base: "Show", past: "Showed", live: "Showing" },
  Set: { base: "Set", past: "Set", live: "Setting" },
  Remember: { base: "Remember", past: "Remembered", live: "Remembering" },
  Reset: { base: "Reset", past: "Reset", live: "Resetting" },
  Restore: { base: "Restore", past: "Restored", live: "Restoring" },
  Clear: { base: "Clear", past: "Cleared", live: "Clearing" },
  List: { base: "List", past: "Listed", live: "Listing" },
  "Read skill": { base: "Read skill", past: "Read skill", live: "Reading skill" },
  "Refresh skill": { base: "Refresh skill", past: "Refreshed skill", live: "Refreshing skill" },
  "Install skill": { base: "Install skill", past: "Installed skill", live: "Installing skill" },
  "Uninstall skill": { base: "Uninstall skill", past: "Uninstalled skill", live: "Uninstalling skill" },
  "List skills": { base: "List skills", past: "Listed skills", live: "Listing skills" },
  "MCP call": { base: "Call", past: "Called", live: "Calling" },
  "MCP connect": { base: "Connect to", past: "Connected to", live: "Connecting to" },
  "MCP describe": { base: "Describe", past: "Described", live: "Describing" },
  "MCP search": { base: "Search MCP for", past: "Searched MCP for", live: "Searching MCP for" },
  "MCP list": { base: "List tools on", past: "Listed tools on", live: "Listing tools on" },
  MCP: { base: "Check MCP", past: "Checked MCP", live: "Checking MCP" },
  "Web search": { base: "Search the web for", past: "Searched the web for", live: "Searching the web for" },
  "Fetch page": { base: "Fetch", past: "Fetched", live: "Fetching" },
  Files: { base: "List files", past: "Listed files", live: "Listing files" },
  "Read file": { base: "Read file", past: "Read file", live: "Reading file" },
  "Write file": { base: "Write file", past: "Wrote file", live: "Writing file" },
  "Delete file": { base: "Delete file", past: "Deleted file", live: "Deleting file" },
  Extensions: { base: "List extensions", past: "Listed extensions", live: "Listing extensions" },
  "Install extension": { base: "Install extension", past: "Installed extension", live: "Installing extension" },
  "Enable extension": { base: "Enable extension", past: "Enabled extension", live: "Enabling extension" },
  "Disable extension": { base: "Disable extension", past: "Disabled extension", live: "Disabling extension" },
  "Reload extension": { base: "Reload extension", past: "Reloaded extension", live: "Reloading extension" },
  "Uninstall extension": { base: "Uninstall extension", past: "Uninstalled extension", live: "Uninstalling extension" },
  "Run Office.js": { base: "Run Office.js", past: "Ran Office.js", live: "Running Office.js" },
  "Run WPS JSAPI": { base: "Run WPS JSAPI", past: "Ran WPS JSAPI", live: "Running WPS JSAPI" },
};

const TOOL_KINDS: Record<string, ToolKind> = {
  python_transform_range: "python",
  python_run: "python",
  libreoffice_convert: "python",
  web_search: "web",
  fetch_page: "web",
  execute_office_js: "code",
  execute_wps_js: "code",
  mcp: "extension",
  extensions_manager: "extension",
};

/** In the order the group header stacks them. */
export const TOOL_KIND_ICONS: ReadonlyMap<ToolKind, IconNode> = new Map([
  ["sheet", Table],
  ["python", Terminal],
  ["web", Globe],
  ["code", Code],
  ["extension", Puzzle],
]);

export function toolKind(toolName: string): ToolKind {
  return TOOL_KINDS[toolName] ?? (isSupportedToolName(toolName) ? "sheet" : "extension");
}

function lowerFirst(text: string): string {
  return text.charAt(0).toLowerCase() + text.slice(1);
}

/** The imperative verb for an approval button: "Fill", "Delete". */
export function toolActionVerb(toolName: string, params: unknown): string | null {
  if (!isSupportedToolName(toolName)) return null;
  const desc = describeToolCall(toolName, params, undefined, undefined);
  return VERBS[desc.action]?.base.split(" ")[0] ?? null;
}

function cellCount(count: number): string {
  return count === 1 ? t("turn.row.cell") : t("turn.row.cells", { count });
}

/** One fact about what the call produced, from its details or its result text. */
function describeMeta(toolName: string, resultText: string | undefined, details: ExcelToolDetails | undefined, badge: string | null): string | null {
  switch (details?.kind) {
    case "write_cells":
    case "fill_formula": {
      if (details.stopped && details.writtenCount !== undefined) return cellCount(details.writtenCount);
      const changed = details.changes?.changedCount;
      if (changed !== undefined) return cellCount(changed);
      break;
    }
    case "python_transform_range": {
      if (details.stopped) {
        return cellCount(details.rowsWritten === undefined ? 0 : details.rowsWritten * (details.colsWritten ?? 1));
      }
      const changed = details.changes?.changedCount;
      if (changed !== undefined) return cellCount(changed);
      break;
    }
    case "read_range_csv":
      return cellCount(details.values.reduce((count, row) => count + row.length, 0));
    case "web_search":
      if (details.resultCount !== undefined) {
        return details.resultCount === 1 ? t("turn.row.result") : t("turn.row.results", { count: details.resultCount });
      }
      break;
    case "charts":
      if (details.count !== undefined) return details.count === 1 ? t("turn.row.chart") : t("turn.row.charts", { count: details.count });
      break;
    case "skills_list":
    case "files_list":
      return details.count === 1 ? t("turn.row.item") : t("turn.row.items", { count: details.count });
    default:
      break;
  }
  if (toolName === "read_range" && resultText) {
    const match = /\((\d+)[×x](\d+)\)/u.exec(extractSummaryLine(resultText) ?? "");
    if (match) return cellCount(Number(match[1]) * Number(match[2]));
  }
  if (toolName === "search_workbook" && resultText) {
    const match = /\b(\d+)\s+match/u.exec(extractSummaryLine(resultText) ?? "");
    const count = match ? Number(match[1]) : null;
    if (count !== null) return count === 1 ? t("turn.row.match") : t("turn.row.matches", { count });
  }
  return badge;
}

/**
 * The row for a tool call. `tense` is the row's state; the sentence keeps
 * the card describer's reference (mono, clickable when it is a range).
 */
export function describeToolRow(
  toolName: string,
  params: ToolCall["arguments"],
  result: ToolResultMessage | undefined,
  tense: ToolRowTense,
  label?: string,
): ToolRowDescription {
  const resultText = result ? splitToolResultContent(result).text : undefined;
  const details = decodeToolDetails(result?.details);
  let reason: string | null = null;
  if (tense === "failed") {
    const error = details && "error" in details ? details.error : null;
    const firstLine = (resultText ?? "").split("\n").find((line) => line.trim() !== "") ?? "";
    reason = extractResultError(resultText) || error || firstLine
      .replace(/^\s*(?:⚠️|❌)?\s*(?:\*\*)?error(?:\*\*)?:?\s*/iu, "").replace(/\*\*/gu, "").trim() || null;
  }

  if (!isSupportedToolName(toolName)) {
    const name = label ?? toolName;
    const text = tense === "live" ? t("turn.row.running", { name })
      : tense === "failed" ? t("turn.row.couldNot", { action: t("turn.row.run", { name }) })
      : tense === "stopped" ? t("turn.row.stopped", { action: t("turn.row.running", { name }) })
      : t("turn.row.ran", { name });
    return { sentence: html`${text}`, meta: null, reason };
  }

  const desc = describeToolCall(toolName, params, resultText, details);
  const badgeIndex = desc.detail.indexOf(" — ");
  const { detail, badge } = desc.detail.startsWith("— ") ? { detail: "", badge: desc.detail.slice(2) }
    : badgeIndex < 0 ? { detail: desc.detail, badge: null }
    : { detail: desc.detail.slice(0, badgeIndex), badge: desc.detail.slice(badgeIndex + 3) };
  // A stopped write already says "Stopped" and how far it got.
  const forms = VERBS[desc.action];
  const verb = desc.action === t("tool-card.stopped.action") ? desc.action
    : tense === "failed" ? t("turn.row.couldNot", { action: lowerFirst(forms?.base ?? desc.action) })
    : tense === "stopped" ? t("turn.row.stopped", { action: lowerFirst(forms?.live ?? desc.action) })
    : tense === "live" ? forms?.live ?? desc.action : forms?.past ?? desc.action;
  // Result summaries end in a full stop ("Added sheet "Q3"."); a row does not.
  const refText = stripDimensions(detail).replace(/\.$/u, "");
  const reference = desc.address
    ? (refText && refText !== desc.address ? cellRefDisplay(refText, desc.address) : cellRefs(desc.address))
    : refText;
  const sentence = reference === ""
    ? html`${verb}`
    : html`${verb} <span class="pi-turn-row__ref">${reference}</span>`;
  return {
    sentence,
    meta: tense === "failed" || tense === "live" ? null : describeMeta(toolName, resultText, details, tense === "stopped" ? null : badge),
    reason,
  };
}

function formatJson(value: unknown): { code: string; language: string } {
  const text = typeof value === "string" ? value : JSON.stringify(value, null, 2) ?? String(value);
  try {
    return { code: JSON.stringify(JSON.parse(text), null, 2), language: "json" };
  } catch {
    return { code: text, language: "text" };
  }
}

/**
 * The row's expanded body: the tool card's input and result sections for a
 * workbook tool; the raw input and output for an extension tool.
 */
export function renderToolRowDetail(toolName: string, params: ToolCall["arguments"], result: ToolResultMessage | undefined): TemplateResult {
  if (isSupportedToolName(toolName)) return renderToolCardBody(toolName, params, result);
  const input = params === undefined || params === null ? null : formatJson(params);
  const output = result ? formatJson(result.content.filter((part) => part.type === "text").map((part) => part.text).join("\n") || t("messages.noOutput")) : null;
  const section = (label: string, { code, language }: { code: string; language: string }) => html`<div>
    <div class="pi-tool-fallback__label">${label}</div>
    <code-block .code=${code} language=${language}></code-block>
  </div>`;
  return html`<div class="pi-tool-fallback">
    ${input ? section(t("messages.toolInput"), input) : nothing}
    ${output ? section(t("messages.toolOutput"), output) : nothing}
  </div>`;
}
