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

import type { ToolResultMessage } from "@earendil-works/pi-ai";
import { html, type TemplateResult } from "lit";
import { Code, Globe, Puzzle, Table, Terminal, type IconNode } from "lucide";

import { t } from "../language/index.js";
import { decodeToolDetails } from "../tools/tool-details.js";
import { cellRefDisplay, cellRefs } from "./cell-link.js";
import { renderTool } from "./messages/tool-renderer-registry.js";
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
  /** Plain-text form of the sentence, for announcements and tests. */
  text: string;
  /** One fact about the result ("9 cells", "6 results"), or null. */
  meta: string | null;
  kind: ToolKind;
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

const KIND_ICONS: Record<ToolKind, IconNode> = {
  sheet: Table,
  python: Terminal,
  web: Globe,
  code: Code,
  extension: Puzzle,
};

export function toolKindIcon(kind: ToolKind): IconNode {
  return KIND_ICONS[kind];
}

export function toolKind(toolName: string): ToolKind {
  return TOOL_KINDS[toolName] ?? (isSupportedToolName(toolName) ? "sheet" : "extension");
}

function lowerFirst(text: string): string {
  return text.charAt(0).toLowerCase() + text.slice(1);
}

/** The card verb, such as Fill or Web search, in the tense a row needs. */
function conjugate(action: string, tense: ToolRowTense): string {
  const forms = VERBS[action];
  if (!forms) {
    // Unknown verb (a result summary's first word, or an extension's label): leave it alone.
    return tense === "failed" ? t("turn.row.couldNot", { action: lowerFirst(action) })
      : tense === "stopped" ? t("turn.row.stopped", { action: lowerFirst(action) })
      : action;
  }
  switch (tense) {
    case "live": return forms.live;
    case "done": return forms.past;
    case "failed": return t("turn.row.couldNot", { action: lowerFirst(forms.base) });
    case "stopped": return t("turn.row.stopped", { action: lowerFirst(forms.live) });
  }
}

/** The imperative verb for an approval button: "Fill", "Delete". */
export function toolActionVerb(toolName: string, params: unknown): string | null {
  if (!isSupportedToolName(toolName)) return null;
  const desc = describeToolCall(toolName, params, undefined, undefined);
  return VERBS[desc.action]?.base.split(" ")[0] ?? null;
}

/** The card's " — 9 changed, no backup" suffix, as the meta fact and the remainder. */
function splitBadge(detail: string): { detail: string; badge: string | null } {
  if (detail.startsWith("— ")) return { detail: "", badge: detail.slice(2) };
  const index = detail.indexOf(" — ");
  if (index < 0) return { detail, badge: null };
  return { detail: detail.slice(0, index), badge: detail.slice(index + 3) };
}

function cellCount(count: number): string {
  return count === 1 ? t("turn.row.cell") : t("turn.row.cells", { count });
}

/** One fact about what the call produced, from its details or its result text. */
function describeMeta(toolName: string, resultText: string | undefined, result: ToolResultMessage | undefined, badge: string | null): string | null {
  const details = decodeToolDetails(result?.details);
  switch (details?.kind) {
    case "write_cells":
    case "fill_formula": {
      if (details.stopped && typeof details.writtenCount === "number") return cellCount(details.writtenCount);
      const changed = details.changes?.changedCount;
      if (typeof changed === "number") return cellCount(changed);
      break;
    }
    case "python_transform_range": {
      if (details.stopped) {
        return cellCount(details.rowsWritten === undefined ? 0 : details.rowsWritten * (details.colsWritten ?? 1));
      }
      const changed = details.changes?.changedCount;
      if (typeof changed === "number") return cellCount(changed);
      break;
    }
    case "read_range_csv":
      return cellCount(details.values.reduce((count, row) => count + row.length, 0));
    case "web_search":
      if (typeof details.resultCount === "number") {
        return details.resultCount === 1 ? t("turn.row.result") : t("turn.row.results", { count: details.resultCount });
      }
      break;
    case "charts":
      if (typeof details.count === "number") return details.count === 1 ? t("turn.row.chart") : t("turn.row.charts", { count: details.count });
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
  params: unknown,
  result: ToolResultMessage | undefined,
  tense: ToolRowTense,
  label?: string,
): ToolRowDescription {
  const kind = toolKind(toolName);
  const resultText = result ? splitToolResultContent(result).text : undefined;
  const reason = tense === "failed" ? failureReason(result, resultText) : null;

  if (!isSupportedToolName(toolName)) {
    const name = label ?? toolName;
    const text = tense === "live" ? t("turn.row.running", { name })
      : tense === "failed" ? t("turn.row.couldNot", { action: t("turn.row.run", { name }) })
      : tense === "stopped" ? t("turn.row.stopped", { action: t("turn.row.running", { name }) })
      : t("turn.row.ran", { name });
    return { sentence: html`${text}`, text, meta: null, kind, reason };
  }

  const desc = describeToolCall(toolName, params, resultText, decodeToolDetails(result?.details));
  const { detail, badge } = splitBadge(desc.detail);
  // A stopped write already says "Stopped" and how far it got.
  const verb = desc.action === t("tool-card.stopped.action") ? desc.action : conjugate(desc.action, tense);
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
    text: refText ? `${verb} ${refText}` : verb,
    meta: tense === "failed" || tense === "live" ? null : describeMeta(toolName, resultText, result, tense === "stopped" ? null : badge),
    kind,
    reason,
  };
}

/** The error text the model received, with the "Error:" prefix removed. */
function failureReason(result: ToolResultMessage | undefined, resultText: string | undefined): string | null {
  if (!result) return null;
  const fromSummary = extractResultError(resultText);
  if (fromSummary) return fromSummary;
  const details = decodeToolDetails(result.details);
  if (details && "error" in details && typeof details.error === "string" && details.error) return details.error;
  const firstLine = (resultText ?? "").split("\n").find((line) => line.trim() !== "") ?? "";
  return firstLine.replace(/^\s*(?:⚠️|❌)?\s*(?:\*\*)?error(?:\*\*)?:?\s*/iu, "").replace(/\*\*/gu, "").trim() || null;
}

/** The row's expanded body: the tool card's input and result sections. */
export function renderToolRowDetail(toolName: string, params: unknown, result: ToolResultMessage | undefined): TemplateResult {
  if (isSupportedToolName(toolName)) return renderToolCardBody(toolName, params, result);
  return renderTool(toolName, params, result, false, { expanded: true, toggle: () => {} }).content;
}
