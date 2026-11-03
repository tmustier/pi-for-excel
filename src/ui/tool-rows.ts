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

export type ToolRowTense = "live" | "done" | "failed" | "stopped";

export type ToolKind = "sheet" | "python" | "web" | "code" | "extension";

export interface ToolRowDescription {
  sentence: TemplateResult;
  meta: string | null;
  reason: string | null;
}

interface VerbForms {
  base: string;
  past: string;
  live: string;
  button?: string;
}

const FILL: VerbForms = { base: "turn.verb.fill.base", past: "turn.verb.fill.past", live: "turn.verb.fill.live" };
const TRANSFORM: VerbForms = { base: "turn.verb.transform.base", past: "turn.verb.transform.past", live: "turn.verb.transform.live" };
const ADD: VerbForms = { base: "turn.verb.add.base", past: "turn.verb.add.past", live: "turn.verb.add.live" };
const RENAME: VerbForms = { base: "turn.verb.rename.base", past: "turn.verb.rename.past", live: "turn.verb.rename.live" };
const DELETE: VerbForms = { base: "turn.verb.delete.base", past: "turn.verb.delete.past", live: "turn.verb.delete.live" };
const INSERT: VerbForms = { base: "turn.verb.insert.base", past: "turn.verb.insert.past", live: "turn.verb.insert.live" };
const HIDE: VerbForms = { base: "turn.verb.hide.base", past: "turn.verb.hide.past", live: "turn.verb.hide.live" };

const VERBS: Record<string, VerbForms> = {
  Read: { base: "turn.verb.read.base", past: "turn.verb.read.past", live: "turn.verb.read.live" },
  Export: { base: "turn.verb.export.base", past: "turn.verb.export.past", live: "turn.verb.export.live" },
  Overview: { base: "turn.verb.read_the_overview_of.base", past: "turn.verb.read_the_overview_of.past", live: "turn.verb.read_the_overview_of.live", button: "turn.verb.read_the_overview_of.button" },
  Edit: { base: "turn.verb.edit.base", past: "turn.verb.edit.past", live: "turn.verb.edit.live" },
  Write: { base: "turn.verb.write.base", past: "turn.verb.write.past", live: "turn.verb.write.live" },
  Fill: FILL,
  Filled: FILL,
  Transform: TRANSFORM,
  Transformed: TRANSFORM,
  Format: { base: "turn.verb.format.base", past: "turn.verb.format.past", live: "turn.verb.format.live" },
  "Cond. format": { base: "turn.verb.conditionally_format.base", past: "turn.verb.conditionally_format.past", live: "turn.verb.conditionally_format.live" },
  Add: ADD,
  Rename: RENAME,
  Delete: DELETE,
  Insert: INSERT,
  Modify: { base: "turn.verb.modify.base", past: "turn.verb.modify.past", live: "turn.verb.modify.live" },
  Deleted: DELETE,
  Inserted: INSERT,
  Added: ADD,
  Renamed: RENAME,
  Duplicated: { base: "turn.verb.duplicate.base", past: "turn.verb.duplicate.past", live: "turn.verb.duplicate.live" },
  Hidden: HIDE,
  Unhidden: { base: "turn.verb.unhide.base", past: "turn.verb.unhide.past", live: "turn.verb.unhide.live" },
  Search: { base: "turn.verb.search.base", past: "turn.verb.search.past", live: "turn.verb.search.live" },
  "Trace precedents": { base: "turn.verb.trace_precedents_of.base", past: "turn.verb.trace_precedents_of.past", live: "turn.verb.trace_precedents_of.live", button: "turn.verb.trace_precedents_of.button" },
  "Trace dependents": { base: "turn.verb.trace_dependents_of.base", past: "turn.verb.trace_dependents_of.past", live: "turn.verb.trace_dependents_of.live", button: "turn.verb.trace_dependents_of.button" },
  "Explain formula": { base: "turn.verb.explain_the_formula_in.base", past: "turn.verb.explain_the_formula_in.past", live: "turn.verb.explain_the_formula_in.live", button: "turn.verb.explain_the_formula_in.button" },
  Create: { base: "turn.verb.create.base", past: "turn.verb.create.past", live: "turn.verb.create.live" },
  Charts: { base: "turn.verb.list_charts_in.base", past: "turn.verb.list_charts_in.past", live: "turn.verb.list_charts_in.live", button: "turn.verb.list_charts_in.button" },
  "Chart image": { base: "turn.verb.capture_chart.base", past: "turn.verb.capture_chart.past", live: "turn.verb.capture_chart.live" },
  Comments: { base: "turn.verb.read_comments_on.base", past: "turn.verb.read_comments_on.past", live: "turn.verb.read_comments_on.live", button: "turn.verb.read_comments_on.button" },
  Comment: { base: "turn.verb.comment_on.base", past: "turn.verb.comment_on.past", live: "turn.verb.comment_on.live", button: "turn.verb.comment_on.button" },
  Reply: { base: "turn.verb.reply_on.base", past: "turn.verb.reply_on.past", live: "turn.verb.reply_on.live", button: "turn.verb.reply_on.button" },
  Resolve: { base: "turn.verb.resolve.base", past: "turn.verb.resolve.past", live: "turn.verb.resolve.live" },
  Reopen: { base: "turn.verb.reopen.base", past: "turn.verb.reopen.past", live: "turn.verb.reopen.live" },
  Update: { base: "turn.verb.update.base", past: "turn.verb.update.past", live: "turn.verb.update.live" },
  View: { base: "turn.verb.view.base", past: "turn.verb.view.past", live: "turn.verb.view.live" },
  Activate: { base: "turn.verb.activate.base", past: "turn.verb.activate.past", live: "turn.verb.activate.live" },
  Freeze: { base: "turn.verb.freeze.base", past: "turn.verb.freeze.past", live: "turn.verb.freeze.live" },
  Hide: HIDE,
  Show: { base: "turn.verb.show.base", past: "turn.verb.show.past", live: "turn.verb.show.live" },
  Set: { base: "turn.verb.set.base", past: "turn.verb.set.past", live: "turn.verb.set.live" },
  Remember: { base: "turn.verb.remember.base", past: "turn.verb.remember.past", live: "turn.verb.remember.live" },
  Reset: { base: "turn.verb.reset.base", past: "turn.verb.reset.past", live: "turn.verb.reset.live" },
  Restore: { base: "turn.verb.restore.base", past: "turn.verb.restore.past", live: "turn.verb.restore.live" },
  Clear: { base: "turn.verb.clear.base", past: "turn.verb.clear.past", live: "turn.verb.clear.live" },
  List: { base: "turn.verb.list.base", past: "turn.verb.list.past", live: "turn.verb.list.live" },
  "Read skill": { base: "turn.verb.read_skill.base", past: "turn.verb.read_skill.past", live: "turn.verb.read_skill.live" },
  "Refresh skill": { base: "turn.verb.refresh_skill.base", past: "turn.verb.refresh_skill.past", live: "turn.verb.refresh_skill.live" },
  "Install skill": { base: "turn.verb.install_skill.base", past: "turn.verb.install_skill.past", live: "turn.verb.install_skill.live" },
  "Uninstall skill": { base: "turn.verb.uninstall_skill.base", past: "turn.verb.uninstall_skill.past", live: "turn.verb.uninstall_skill.live" },
  "List skills": { base: "turn.verb.list_skills.base", past: "turn.verb.list_skills.past", live: "turn.verb.list_skills.live" },
  "MCP call": { base: "turn.verb.call.base", past: "turn.verb.call.past", live: "turn.verb.call.live" },
  "MCP connect": { base: "turn.verb.connect_to.base", past: "turn.verb.connect_to.past", live: "turn.verb.connect_to.live", button: "turn.verb.connect_to.button" },
  "MCP describe": { base: "turn.verb.describe.base", past: "turn.verb.describe.past", live: "turn.verb.describe.live" },
  "MCP search": { base: "turn.verb.search_mcp_for.base", past: "turn.verb.search_mcp_for.past", live: "turn.verb.search_mcp_for.live", button: "turn.verb.search_mcp_for.button" },
  "MCP list": { base: "turn.verb.list_tools_on.base", past: "turn.verb.list_tools_on.past", live: "turn.verb.list_tools_on.live", button: "turn.verb.list_tools_on.button" },
  MCP: { base: "turn.verb.check_mcp.base", past: "turn.verb.check_mcp.past", live: "turn.verb.check_mcp.live" },
  "Web search": { base: "turn.verb.search_the_web_for.base", past: "turn.verb.search_the_web_for.past", live: "turn.verb.search_the_web_for.live", button: "turn.verb.search_the_web_for.button" },
  "Fetch page": { base: "turn.verb.fetch.base", past: "turn.verb.fetch.past", live: "turn.verb.fetch.live" },
  Files: { base: "turn.verb.list_files.base", past: "turn.verb.list_files.past", live: "turn.verb.list_files.live" },
  "Read file": { base: "turn.verb.read_file.base", past: "turn.verb.read_file.past", live: "turn.verb.read_file.live" },
  "Write file": { base: "turn.verb.write_file.base", past: "turn.verb.write_file.past", live: "turn.verb.write_file.live" },
  "Delete file": { base: "turn.verb.delete_file.base", past: "turn.verb.delete_file.past", live: "turn.verb.delete_file.live" },
  Extensions: { base: "turn.verb.list_extensions.base", past: "turn.verb.list_extensions.past", live: "turn.verb.list_extensions.live" },
  "Install extension": { base: "turn.verb.install_extension.base", past: "turn.verb.install_extension.past", live: "turn.verb.install_extension.live" },
  "Enable extension": { base: "turn.verb.enable_extension.base", past: "turn.verb.enable_extension.past", live: "turn.verb.enable_extension.live" },
  "Disable extension": { base: "turn.verb.disable_extension.base", past: "turn.verb.disable_extension.past", live: "turn.verb.disable_extension.live" },
  "Reload extension": { base: "turn.verb.reload_extension.base", past: "turn.verb.reload_extension.past", live: "turn.verb.reload_extension.live" },
  "Uninstall extension": { base: "turn.verb.uninstall_extension.base", past: "turn.verb.uninstall_extension.past", live: "turn.verb.uninstall_extension.live" },
  "Run Office.js": { base: "turn.verb.run_office_js.base", past: "turn.verb.run_office_js.past", live: "turn.verb.run_office_js.live" },
  "Run WPS JSAPI": { base: "turn.verb.run_wps_jsapi.base", past: "turn.verb.run_wps_jsapi.past", live: "turn.verb.run_wps_jsapi.live" },
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

export function toolActionVerb(toolName: string, params: unknown): string | null {
  if (!isSupportedToolName(toolName)) return null;
  const forms = VERBS[describeToolCall(toolName, params, undefined, undefined).action];
  return forms ? t(forms.button ?? forms.base) : null;
}

export function describePendingToolCall(toolName: string, params: unknown): string | null {
  if (!isSupportedToolName(toolName)) return null;
  const desc = describeToolCall(toolName, params, undefined, undefined);
  const forms = VERBS[desc.action];
  const verb = forms ? t(forms.base) : desc.action;
  return desc.detail ? `${verb} ${desc.detail}` : verb;
}

function cellCount(count: number): string {
  return count === 1 ? t("turn.row.cell") : t("turn.row.cells", { count });
}

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
      : tense === "stopped" ? t("turn.row.stopped", { action: lowerFirst(t("turn.row.running", { name })) })
      : t("turn.row.ran", { name });
    return { sentence: html`${text}`, meta: null, reason };
  }

  const desc = describeToolCall(toolName, params, resultText, details);
  const badgeIndex = desc.detail.indexOf(" — ");
  const { detail, badge } = desc.detail.startsWith("— ") ? { detail: "", badge: desc.detail.slice(2) }
    : badgeIndex < 0 ? { detail: desc.detail, badge: null }
    : { detail: desc.detail.slice(0, badgeIndex), badge: desc.detail.slice(badgeIndex + 3) };
  const forms = VERBS[desc.action];
  const form = (key: "base" | "past" | "live") => forms ? t(forms[key]) : desc.action;
  const verb = desc.action === t("tool-card.stopped.action") ? desc.action
    : tense === "failed" ? t("turn.row.couldNot", { action: lowerFirst(form("base")) })
    : tense === "stopped" ? t("turn.row.stopped", { action: lowerFirst(form("live")) })
    : form(tense === "live" ? "live" : "past");
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
