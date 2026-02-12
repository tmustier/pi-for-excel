/**
 * Human-readable tool input rendering.
 *
 * Converts raw tool parameters into clean key-value lists with
 * color chips, data previews, and friendly descriptions — aimed
 * at Excel-savvy, less-technical users.
 */

import { html, nothing, type TemplateResult } from "lit";
import { cellRef, cellRefs } from "./cell-link.js";
import { formatColorLabel } from "./color-names.js";
import type { CoreToolName } from "../tools/registry.js";

/* ── Types ──────────────────────────────────────────────────── */

interface ParamItem {
  label: string;
  value: TemplateResult | string;
}

/* ── Helpers ────────────────────────────────────────────────── */

function safe(params: unknown): Record<string, unknown> {
  if (!params) return {};
  if (typeof params === "object" && params !== null)
    return params as Record<string, unknown>;
  if (typeof params === "string") {
    try {
      const p: unknown = JSON.parse(params);
      return typeof p === "object" && p !== null
        ? (p as Record<string, unknown>)
        : {};
    } catch {
      return {};
    }
  }
  return {};
}

/** Safely convert an unknown value to string. */
function str(v: unknown): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  return JSON.stringify(v);
}

/** Safely read a number, returning undefined if not a number. */
function num(v: unknown): number | undefined {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = Number(v);
    return isNaN(n) ? undefined : n;
  }
  return undefined;
}

/** Inline color chip (small filled circle) + human-readable name. */
function colorChip(hex: string): TemplateResult {
  const label = formatColorLabel(hex);
  // For well-known names, show just the name. For raw hex, show hex.
  const display = label === hex ? hex : label;
  return html`<span
      class="pi-color-chip"
      style="background:${hex}"
    ></span
    ><span class="pi-color-chip-label">${display}</span>`;
}

/* ── Range parsing & grouping ───────────────────────────────── */

/** Split "Sheet1!A1:B2" → { sheet: "Sheet1", address: "A1:B2" }. */
function splitRangeRef(ref: string): { sheet: string; address: string } {
  const bang = ref.indexOf("!");
  if (bang >= 0) {
    return {
      sheet: ref.substring(0, bang).replace(/^'|'$/g, ""),
      address: ref.substring(bang + 1),
    };
  }
  return { sheet: "", address: ref };
}

interface RangeDisplayResult {
  /** Common sheet name if all ranges share one, otherwise empty. */
  sheet: string;
  /** Address display (sheet prefix stripped, truncated with "+N more"). */
  display: TemplateResult | string;
}

/**
 * Parse a comma/semicolon-separated range string, extract a common
 * sheet prefix, strip it from individual addresses, and truncate.
 *
 * Examples:
 *   "Sheet1!A1, Sheet1!B2, Sheet1!C3" → sheet="Sheet1", display="A1, B2, C3"
 *   "A1, B2, C3"                       → sheet="",       display="A1, B2, C3"
 *   "Sheet1!A1:D1, Sheet2!E1"          → sheet="",       display="Sheet1!A1:D1, Sheet2!E1"
 */
function formatRangeForDisplay(range: string, maxShow = 3): RangeDisplayResult {
  const parts = range
    .split(/\s*[,;]\s*/)
    .map((s) => s.trim())
    .filter(Boolean);

  // Parse each part
  const parsed = parts.map(splitRangeRef);

  // Find common sheet (only if ALL parts with a sheet agree)
  const sheetsFound = [
    ...new Set(parsed.map((p) => p.sheet).filter(Boolean)),
  ];
  const commonSheet = sheetsFound.length === 1 ? sheetsFound[0] : "";

  // Build display addresses — strip the common sheet prefix
  const addresses = commonSheet
    ? parsed.map((p) => p.address)
    : parts; // keep originals if sheets differ

  // Truncate
  if (addresses.length <= maxShow) {
    return { sheet: commonSheet, display: addresses.join(", ") };
  }
  const shown = addresses.slice(0, maxShow).join(", ");
  const more = addresses.length - maxShow;
  return {
    sheet: commonSheet,
    display: html`${shown}
      <span class="pi-params__more">+${more} more</span>`,
  };
}

/** Format a cell value for preview. */
function fmtCell(v: unknown): string {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "string") return v.length > 18 ? v.substring(0, 18) + "…" : v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  return JSON.stringify(v);
}

/**
 * Render a mini data-preview table for write_cells values.
 * Shows up to 3 rows × 6 columns, with truncation indicators.
 */
function renderDataPreview(values: unknown[][]): TemplateResult {
  const MAX_ROWS = 3;
  const MAX_COLS = 6;
  const totalRows = values.length;
  const totalCols = Math.max(...values.map((r) => (Array.isArray(r) ? r.length : 0)));
  const showRows = Math.min(totalRows, MAX_ROWS);
  const showCols = Math.min(totalCols, MAX_COLS);
  const moreRows = totalRows - showRows;
  const moreCols = totalCols - showCols;

  return html`
    <table class="pi-data-preview">
      ${values.slice(0, showRows).map(
        (row) => html`
          <tr>
            ${(Array.isArray(row) ? row : [row]).slice(0, showCols).map(
              (cell) => html`<td>${fmtCell(cell)}</td>`,
            )}
            ${moreCols > 0 ? html`<td class="pi-data-preview__fade">…</td>` : nothing}
          </tr>
        `,
      )}
      ${moreRows > 0
        ? html`<tr>
            <td
              colspan=${showCols + (moreCols > 0 ? 1 : 0)}
              class="pi-data-preview__fade"
            >
              …${moreRows} more row${moreRows !== 1 ? "s" : ""}
            </td>
          </tr>`
        : nothing}
    </table>
  `;
}

/** Render a monospaced formula snippet. */
function formulaSnippet(formula: string): TemplateResult {
  return html`<code class="pi-params__code">${formula}</code>`;
}

function plural(n: number, word: string): string {
  return `${n} ${word}${n !== 1 ? "s" : ""}`;
}

/* ── Layout ─────────────────────────────────────────────────── */

function renderParamList(items: ParamItem[]): TemplateResult {
  return html`
    <div class="pi-params">
      ${items.map(
        (item) => html`
          <div class="pi-params__row">
            <span class="pi-params__label">${item.label}</span>
            <span class="pi-params__value">${item.value}</span>
          </div>
        `,
      )}
    </div>
  `;
}

/* ── Per-tool humanizers ────────────────────────────────────── */

function humanizeFormatCells(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  // Range (with sheet grouping)
  if (p.range) {
    const rd = formatRangeForDisplay(str(p.range));
    if (rd.sheet) items.push({ label: "Sheet", value: rd.sheet });
    items.push({ label: "Range", value: cellRefs(str(p.range), Infinity) });
  }

  // Named styles
  if (p.style) {
    const names = Array.isArray(p.style)
      ? (p.style as unknown[]).map(str)
      : [str(p.style)];
    items.push({ label: "Style", value: names.join(" + ") });
  }

  // Font properties — grouped into one row
  const fontParts: Array<TemplateResult | string> = [];
  if (p.font_color) fontParts.push(colorChip(str(p.font_color)));
  if (p.bold === true) fontParts.push("bold");
  if (p.italic === true) fontParts.push("italic");
  if (p.underline === true) fontParts.push("underline");
  if (p.font_size) fontParts.push(str(p.font_size) + "pt");
  if (p.font_name) fontParts.push(str(p.font_name));
  if (fontParts.length > 0) {
    items.push({ label: "Font", value: joinParts(fontParts) });
  }

  // Fill
  if (p.fill_color) {
    items.push({ label: "Fill", value: colorChip(str(p.fill_color)) });
  }

  // Number format
  if (p.number_format) {
    const nf = str(p.number_format);
    const dp = num(p.number_format_dp);
    const sym = p.currency_symbol ? str(p.currency_symbol) : "";
    let display = nf;
    if (dp !== undefined) display += ` (${String(dp)}dp)`;
    if (sym) display += " " + sym;
    items.push({ label: "Format", value: display });
  }

  // Alignment
  const alignParts: string[] = [];
  if (p.horizontal_alignment) alignParts.push(str(p.horizontal_alignment));
  if (p.vertical_alignment) alignParts.push("v: " + str(p.vertical_alignment));
  if (p.wrap_text === true) alignParts.push("wrap");
  if (alignParts.length > 0) {
    items.push({ label: "Align", value: alignParts.join(", ") });
  }

  // Dimensions
  const cw = num(p.column_width);
  if (cw !== undefined) {
    items.push({ label: "Width", value: String(cw) + " chars" });
  }
  const rh = num(p.row_height);
  if (rh !== undefined) {
    items.push({ label: "Height", value: String(rh) + "pt" });
  }
  if (p.auto_fit === true) {
    items.push({ label: "Auto-fit", value: "yes" });
  }

  // Borders
  const edgeLabels: string[] = [];
  if (p.border_top) edgeLabels.push("top " + str(p.border_top));
  if (p.border_bottom) edgeLabels.push("bottom " + str(p.border_bottom));
  if (p.border_left) edgeLabels.push("left " + str(p.border_left));
  if (p.border_right) edgeLabels.push("right " + str(p.border_right));
  if (edgeLabels.length > 0) {
    items.push({ label: "Borders", value: edgeLabels.join(", ") });
  } else if (p.borders) {
    items.push({ label: "Borders", value: str(p.borders) + " (all edges)" });
  }

  // Merge
  if (p.merge === true) items.push({ label: "Merge", value: "yes" });
  if (p.merge === false) items.push({ label: "Merge", value: "unmerge" });

  return items;
}

function humanizeWriteCells(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  if (p.start_cell) {
    items.push({ label: "Start", value: cellRef(str(p.start_cell)) });
  }

  const rawValues = p.values;
  if (Array.isArray(rawValues) && rawValues.length > 0) {
    const values = rawValues as unknown[][];
    const rows = values.length;
    const cols = Math.max(...values.map((r) => (Array.isArray(r) ? r.length : 0)));
    items.push({
      label: "Size",
      value: plural(rows, "row") + " × " + plural(cols, "column"),
    });
    items.push({ label: "Data", value: renderDataPreview(values) });
  }

  if (p.allow_overwrite === true) {
    items.push({ label: "Overwrite", value: "allowed" });
  }

  return items;
}

function humanizeReadRange(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  if (p.range) {
    const rd = formatRangeForDisplay(str(p.range));
    if (rd.sheet) items.push({ label: "Sheet", value: rd.sheet });
    items.push({ label: "Range", value: cellRefs(str(p.range), Infinity) });
  }
  if (p.mode && p.mode !== "compact") {
    items.push({ label: "Mode", value: str(p.mode) });
  }

  return items;
}

function humanizeFillFormula(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  if (p.range) {
    const rd = formatRangeForDisplay(str(p.range));
    if (rd.sheet) items.push({ label: "Sheet", value: rd.sheet });
    items.push({ label: "Range", value: cellRefs(str(p.range), Infinity) });
  }
  if (p.formula) {
    items.push({ label: "Formula", value: formulaSnippet(str(p.formula)) });
  }
  if (p.allow_overwrite === true) {
    items.push({ label: "Overwrite", value: "allowed" });
  }

  return items;
}

function humanizeSearchWorkbook(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  if (p.query) {
    items.push({ label: "Query", value: '"' + str(p.query) + '"' });
  }
  if (p.search_formulas === true) {
    items.push({ label: "Search in", value: "formulas" });
  }
  if (p.use_regex === true) {
    items.push({ label: "Regex", value: "yes" });
  }
  if (p.sheet) {
    items.push({ label: "Sheet", value: str(p.sheet) });
  }
  const ctxRows = num(p.context_rows);
  if (ctxRows !== undefined && ctxRows > 0) {
    items.push({
      label: "Context",
      value: String(ctxRows) + " rows around each match",
    });
  }
  const maxRes = num(p.max_results);
  if (maxRes !== undefined && maxRes !== 20) {
    items.push({ label: "Limit", value: String(maxRes) + " results" });
  }

  return items;
}

function humanizeModifyStructure(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];
  const action = str(p.action);
  const count = num(p.count) ?? 1;
  const pos = num(p.position);

  switch (action) {
    case "insert_rows":
      items.push({
        label: "Action",
        value: "Insert " + plural(count, "row") + (pos !== undefined ? " at row " + String(pos) : ""),
      });
      break;
    case "delete_rows":
      items.push({
        label: "Action",
        value: "Delete " + plural(count, "row") + (pos !== undefined ? " from row " + String(pos) : ""),
      });
      break;
    case "insert_columns":
      items.push({
        label: "Action",
        value: "Insert " + plural(count, "column") + (pos !== undefined ? " at column " + String(pos) : ""),
      });
      break;
    case "delete_columns":
      items.push({
        label: "Action",
        value: "Delete " + plural(count, "column") + (pos !== undefined ? " from column " + String(pos) : ""),
      });
      break;
    case "add_sheet": {
      const name = p.new_name ? str(p.new_name) : p.name ? str(p.name) : "";
      items.push({
        label: "Action",
        value: name ? 'Add sheet "' + name + '"' : "Add sheet",
      });
      break;
    }
    case "delete_sheet":
      items.push({ label: "Action", value: "Delete sheet" });
      break;
    case "rename_sheet": {
      const newName = p.new_name ? str(p.new_name) : "";
      items.push({
        label: "Action",
        value: newName ? 'Rename → "' + newName + '"' : "Rename sheet",
      });
      break;
    }
    case "duplicate_sheet": {
      const targetName = p.new_name ? str(p.new_name) : "";
      items.push({
        label: "Action",
        value: targetName ? 'Duplicate as "' + targetName + '"' : "Duplicate sheet",
      });
      break;
    }
    case "hide_sheet":
      items.push({ label: "Action", value: "Hide sheet" });
      break;
    case "unhide_sheet":
      items.push({ label: "Action", value: "Show sheet" });
      break;
    default:
      items.push({ label: "Action", value: action.replace(/_/g, " ") });
  }

  if (p.sheet) {
    items.push({ label: "Sheet", value: str(p.sheet) });
  }

  return items;
}

function humanizeConditionalFormat(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  // Action
  if (p.action === "clear") {
    items.push({ label: "Action", value: "Clear all rules" });
  } else {
    items.push({ label: "Action", value: "Add rule" });
  }

  // Range (with sheet grouping)
  if (p.range) {
    const rd = formatRangeForDisplay(str(p.range));
    if (rd.sheet) items.push({ label: "Sheet", value: rd.sheet });
    items.push({ label: "Range", value: cellRefs(str(p.range), Infinity) });
  }

  // Rule details
  if (p.type === "formula" && p.formula) {
    items.push({ label: "Rule", value: formulaSnippet(str(p.formula)) });
  } else if (p.type === "cell_value" && p.operator) {
    const op = humanizeOperator(str(p.operator));
    const val = p.value !== undefined ? " " + str(p.value) : "";
    const val2 = p.value2 !== undefined ? " and " + str(p.value2) : "";
    items.push({ label: "Rule", value: op + val + val2 });
  }

  // Format
  const fmtParts: Array<TemplateResult | string> = [];
  if (p.fill_color) fmtParts.push(html`fill ${colorChip(str(p.fill_color))}`);
  if (p.font_color) fmtParts.push(html`font ${colorChip(str(p.font_color))}`);
  if (p.bold === true) fmtParts.push("bold");
  if (p.italic === true) fmtParts.push("italic");
  if (p.underline === true) fmtParts.push("underline");
  if (fmtParts.length > 0) {
    items.push({ label: "Format", value: joinParts(fmtParts) });
  }

  return items;
}

function humanizeTraceDependencies(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  if (p.cell) {
    items.push({ label: "Cell", value: cellRef(str(p.cell)) });
  }
  const depth = num(p.depth);
  if (depth !== undefined && depth !== 2) {
    items.push({
      label: "Depth",
      value: String(depth) + " level" + (depth !== 1 ? "s" : ""),
    });
  }

  return items;
}

function humanizeComments(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  if (p.action) {
    items.push({ label: "Action", value: str(p.action) });
  }

  if (p.range) {
    const rd = formatRangeForDisplay(str(p.range));
    if (rd.sheet) items.push({ label: "Sheet", value: rd.sheet });
    items.push({ label: "Range", value: cellRefs(str(p.range), Infinity) });
  }

  if (p.content) {
    items.push({ label: "Content", value: formulaSnippet(str(p.content)) });
  }

  return items;
}

function humanizeViewSettings(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];
  const action = str(p.action);
  const count = num(p.count);

  switch (action) {
    case "get":
      items.push({ label: "Action", value: "Get current settings" });
      break;
    case "show_gridlines":
      items.push({ label: "Action", value: "Show gridlines" });
      break;
    case "hide_gridlines":
      items.push({ label: "Action", value: "Hide gridlines" });
      break;
    case "show_headings":
      items.push({ label: "Action", value: "Show headings" });
      break;
    case "hide_headings":
      items.push({ label: "Action", value: "Hide headings" });
      break;
    case "freeze_rows":
      items.push({
        label: "Action",
        value: count !== undefined ? "Freeze top " + plural(count, "row") : "Freeze rows",
      });
      break;
    case "freeze_columns":
      items.push({
        label: "Action",
        value:
          count !== undefined
            ? "Freeze first " + plural(count, "column")
            : "Freeze columns",
      });
      break;
    case "freeze_at":
      items.push({
        label: "Action",
        value: p.range ? "Freeze panes at " + str(p.range) : "Freeze panes",
      });
      break;
    case "unfreeze":
      items.push({ label: "Action", value: "Unfreeze panes" });
      break;
    case "set_tab_color":
      items.push({
        label: "Action",
        value: p.color ? html`Tab color ${colorChip(str(p.color))}` : "Clear tab color",
      });
      break;
    case "hide_sheet":
      items.push({ label: "Action", value: "Hide sheet" });
      break;
    case "show_sheet":
      items.push({ label: "Action", value: "Show sheet" });
      break;
    case "very_hide_sheet":
      items.push({ label: "Action", value: "Very hide sheet (VeryHidden)" });
      break;
    case "set_standard_width": {
      const width = num(p.width);
      items.push({
        label: "Action",
        value: width !== undefined
          ? "Set standard width to " + String(width)
          : "Set standard width",
      });
      break;
    }
    case "activate":
      items.push({ label: "Action", value: "Activate sheet" });
      break;
    default:
      items.push({ label: "Action", value: action.replace(/_/g, " ") });
  }

  if (p.sheet) {
    items.push({ label: "Sheet", value: str(p.sheet) });
  }

  return items;
}

function humanizeGetWorkbookOverview(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  if (p.sheet) {
    items.push({ label: "Sheet", value: str(p.sheet) });
  } else {
    items.push({ label: "Scope", value: "Full workbook" });
  }

  return items;
}

function humanizeInstructions(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  if (p.level) {
    items.push({ label: "Scope", value: str(p.level) });
  }

  if (p.action) {
    items.push({ label: "Action", value: str(p.action) });
  }

  if (p.content) {
    const text = str(p.content);
    const compact = text.length > 120 ? `${text.slice(0, 117)}…` : text;
    items.push({ label: "Content", value: compact });
  }

  return items;
}

function humanizeConventions(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];
  const action = str(p.action);

  if (action) {
    items.push({ label: "Action", value: action });
  }

  if (action === "set") {
    const fields: Array<{ key: string; label: string }> = [
      { key: "currency_symbol", label: "Currency" },
      { key: "negative_style", label: "Negatives" },
      { key: "zero_style", label: "Zeros" },
      { key: "thousands_separator", label: "Thousands sep" },
      { key: "accounting_padding", label: "Accounting padding" },
      { key: "number_dp", label: "Number dp" },
      { key: "currency_dp", label: "Currency dp" },
      { key: "percent_dp", label: "Percent dp" },
      { key: "ratio_dp", label: "Ratio dp" },
    ];

    for (const { key, label } of fields) {
      const val = p[key];
      if (val !== undefined && val !== null) {
        const display = typeof val === "string" || typeof val === "number" || typeof val === "boolean"
          ? String(val)
          : JSON.stringify(val);
        items.push({ label, value: display });
      }
    }
  }

  return items;
}

function humanizeWorkbookHistory(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];
  const action = str(p.action || "list");

  items.push({ label: "Action", value: action });

  if (p.snapshot_id) {
    items.push({ label: "Checkpoint", value: str(p.snapshot_id) });
  }

  const limit = num(p.limit);
  if (limit !== undefined) {
    items.push({ label: "Limit", value: String(limit) });
  }

  return items;
}

function humanizeWebSearch(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  if (p.query) {
    items.push({ label: "Query", value: `\"${str(p.query)}\"` });
  }

  if (p.recency) {
    items.push({ label: "Recency", value: str(p.recency) });
  }

  if (p.site) {
    if (Array.isArray(p.site)) {
      const sites = p.site.map((site) => str(site)).filter((site) => site.length > 0);
      items.push({ label: "Sites", value: sites.join(", ") });
    } else {
      items.push({ label: "Site", value: str(p.site) });
    }
  }

  const maxResults = num(p.max_results);
  if (maxResults !== undefined) {
    items.push({ label: "Limit", value: `${maxResults} results` });
  }

  return items;
}

function humanizeMcp(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  if (p.tool) {
    items.push({ label: "Mode", value: "Call tool" });
    items.push({ label: "Tool", value: str(p.tool) });
  } else if (p.connect) {
    items.push({ label: "Mode", value: "Connect" });
    items.push({ label: "Server", value: str(p.connect) });
  } else if (p.describe) {
    items.push({ label: "Mode", value: "Describe tool" });
    items.push({ label: "Tool", value: str(p.describe) });
  } else if (p.search) {
    items.push({ label: "Mode", value: "Search tools" });
    items.push({ label: "Query", value: str(p.search) });
  } else if (p.server) {
    items.push({ label: "Mode", value: "List server tools" });
    items.push({ label: "Server", value: str(p.server) });
  } else {
    items.push({ label: "Mode", value: "Status" });
  }

  if (p.args) {
    const argsText = str(p.args);
    const compact = argsText.length > 120 ? `${argsText.slice(0, 117)}…` : argsText;
    items.push({ label: "Args", value: compact });
  }

  return items;
}

function humanizeFiles(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];
  const action = str(p.action);

  if (action) {
    items.push({ label: "Action", value: action });
  }

  if (p.path) {
    items.push({ label: "Path", value: str(p.path) });
  }

  if (p.mode) {
    items.push({ label: "Read mode", value: str(p.mode) });
  }

  if (p.encoding) {
    items.push({ label: "Encoding", value: str(p.encoding) });
  }

  if (p.mime_type) {
    items.push({ label: "MIME", value: str(p.mime_type) });
  }

  const maxChars = num(p.max_chars);
  if (maxChars !== undefined) {
    items.push({ label: "Max chars", value: String(maxChars) });
  }

  if (p.content !== undefined) {
    const content = str(p.content);
    const compact = content.length > 120 ? `${content.slice(0, 117)}…` : content;
    items.push({ label: "Content", value: compact });
  }

  return items;
}

function humanizePythonTransformRange(p: Record<string, unknown>): ParamItem[] {
  const items: ParamItem[] = [];

  if (p.range) {
    items.push({ label: "Input range", value: cellRefs(str(p.range), Infinity) });
  }

  if (p.output_start_cell) {
    items.push({ label: "Output start", value: cellRefs(str(p.output_start_cell), Infinity) });
  }

  const allowOverwrite = p.allow_overwrite;
  if (typeof allowOverwrite === "boolean") {
    items.push({ label: "Allow overwrite", value: allowOverwrite ? "Yes" : "No" });
  }

  const timeoutMs = num(p.timeout_ms);
  if (timeoutMs !== undefined) {
    items.push({ label: "Timeout", value: `${timeoutMs} ms` });
  }

  if (p.code) {
    const source = str(p.code);
    const lines = source.split(/\r?\n/u).length;
    const oneLine = source.replace(/\s+/gu, " ").trim();
    const compact = oneLine.length > 140 ? `${oneLine.slice(0, 137)}…` : oneLine;
    items.push({ label: "Python", value: compact.length > 0 ? compact : "(empty)" });
    if (lines > 1) {
      items.push({ label: "Code lines", value: String(lines) });
    }
  }

  return items;
}

/* ── Shared helpers ─────────────────────────────────────────── */

/** Join an array of mixed text/TemplateResult with comma separators. */
function joinParts(parts: Array<TemplateResult | string>): TemplateResult {
  return html`${parts.map(
    (part, i) => html`${i > 0 ? ", " : ""}${part}`,
  )}`;
}

/** Convert a cell_value operator to plain English. */
function humanizeOperator(op: string): string {
  const MAP: Record<string, string> = {
    Between: "between",
    NotBetween: "not between",
    EqualTo: "equal to",
    NotEqualTo: "not equal to",
    GreaterThan: "greater than",
    LessThan: "less than",
    GreaterThanOrEqual: "≥",
    LessThanOrEqual: "≤",
  };
  return MAP[op] ?? op;
}

/* ── Registry ───────────────────────────────────────────────── */

type HumanizerFn = (p: Record<string, unknown>) => ParamItem[];

const CORE_HUMANIZERS = {
  format_cells: humanizeFormatCells,
  write_cells: humanizeWriteCells,
  read_range: humanizeReadRange,
  fill_formula: humanizeFillFormula,
  search_workbook: humanizeSearchWorkbook,
  modify_structure: humanizeModifyStructure,
  conditional_format: humanizeConditionalFormat,
  trace_dependencies: humanizeTraceDependencies,
  view_settings: humanizeViewSettings,
  get_workbook_overview: humanizeGetWorkbookOverview,
  comments: humanizeComments,
  instructions: humanizeInstructions,
  conventions: humanizeConventions,
  workbook_history: humanizeWorkbookHistory,
} satisfies Record<CoreToolName, HumanizerFn>;

const EXTRA_HUMANIZERS: Record<string, HumanizerFn> = {
  web_search: humanizeWebSearch,
  mcp: humanizeMcp,
  files: humanizeFiles,
  python_transform_range: humanizePythonTransformRange,
};

const HUMANIZERS: Record<string, HumanizerFn> = {
  ...CORE_HUMANIZERS,
  ...EXTRA_HUMANIZERS,
};

/* ── Public API ─────────────────────────────────────────────── */

/**
 * Convert tool parameters to a human-readable Lit template.
 * Returns `null` for unknown tools (caller falls back to JSON).
 */
export function humanizeToolInput(
  toolName: string,
  params: unknown,
): TemplateResult | null {
  const fn = HUMANIZERS[toolName];
  if (!fn) return null;

  const p = safe(params);
  const items = fn(p);
  if (!items || items.length === 0) return null;

  return renderParamList(items);
}
