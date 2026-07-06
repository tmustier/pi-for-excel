/**
 * Visual dependency tree renderer for trace_dependencies results.
 *
 * Replaces the ASCII-art markdown tree with a structured HTML tree
 * featuring clickable cell references, styled formulas, and proper
 * tree connectors. The text output sent to the agent is unchanged.
 *
 * Same-sheet children show just the cell ref (e.g. "G44") instead
 * of the full qualified address — reduces repetitive noise.
 *
 * Values are formatted according to their Excel number format
 * (%, currency, commas). Negative percentages/currencies use
 * accounting-style parentheses: (1.8%), ($500.00).
 */

import { html, nothing, type TemplateResult } from "lit";
import { cellRef, cellRefDisplay } from "./cell-link.js";
import type { DepNodeDetail, TraceDependenciesMode } from "../tools/tool-details.js";
import { isExcelError } from "../utils/format.js";

/* ── Number format application ──────────────────────────────── */

/** Count decimal-place digits (0 or #) after the dot in a format section. */
function countDecimalDigits(section: string): number {
  const dot = section.indexOf(".");
  if (dot < 0) return 0;
  let n = 0;
  for (let i = dot + 1; i < section.length; i++) {
    const ch = section[i];
    if (ch === "0" || ch === "#") n++;
    else break;
  }
  return n;
}

/** Detect a currency symbol in the format string. */
function extractCurrency(section: string): string | null {
  for (const sym of ["$", "£", "€", "¥", "₹", "CHF", "kr"]) {
    if (section.includes(sym)) return sym;
  }
  // Locale-prefixed: [$€-407], [$$-409], etc.
  const locale = section.match(/\[\$([^\]-]+)/);
  if (locale) {
    const localeSymbol = locale[1];
    if (localeSymbol) {
      const sym = localeSymbol.replace(/-.*/, "").trim();
      if (sym.length > 0) return sym;
    }
  }
  return null;
}

/**
 * Apply an Excel number format to a numeric value.
 * Handles %, currency, commas. Negatives use accounting parens
 * for % and currency formats.
 */
function applyNumberFormat(value: number, fmt: string): string {
  // Use the positive section (first section before ";")
  const section = fmt.split(";")[0] ?? fmt;

  // ── Percentage ────────────────────────────────────────
  if (section.includes("%")) {
    const pct = value * 100;
    const dp = countDecimalDigits(section);
    const str = Math.abs(pct).toFixed(dp) + "%";
    return pct < 0 ? `(${str})` : str;
  }

  // ── Currency / accounting ─────────────────────────────
  const curr = extractCurrency(section);
  if (curr) {
    const dp = countDecimalDigits(section);
    const str =
      curr +
      Math.abs(value).toLocaleString("en-US", {
        minimumFractionDigits: dp,
        maximumFractionDigits: dp,
      });
    return value < 0 ? `(${str})` : str;
  }

  // ── Comma grouping (no currency) ─────────────────────
  if (section.includes(",")) {
    const dp = countDecimalDigits(section);
    return value.toLocaleString("en-US", {
      minimumFractionDigits: dp,
      maximumFractionDigits: dp,
    });
  }

  // ── Fixed decimal (0.00, 0.0, …) ─────────────────────
  if (section.includes(".")) {
    const dp = countDecimalDigits(section);
    if (dp > 0) return value.toFixed(dp);
  }

  // ── Fallback → smart formatting ───────────────────────
  return formatNumber(value);
}

/* ── Smart fallback formatting (no Excel format available) ─── */

const intFmt = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });
const lgFmt = new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 });
const mdFmt = new Intl.NumberFormat("en-US", { maximumFractionDigits: 4 });

function formatNumber(n: number): string {
  if (Number.isInteger(n)) return intFmt.format(n);
  const abs = Math.abs(n);
  if (abs >= 100) return lgFmt.format(n);
  if (abs >= 0.01) return mdFmt.format(n);
  return n.toPrecision(3);
}

/* ── Value display ──────────────────────────────────────────── */

/** Format a cell value for display, using Excel number format when available. */
function fmtValue(value: DynamicValue, numberFormat?: string): string {
  if (value === "" || value === null || value === undefined) return "";
  if (typeof value === "number") {
    return numberFormat && numberFormat !== "General"
      ? applyNumberFormat(value, numberFormat)
      : formatNumber(value);
  }
  if (typeof value === "string") return value;
  if (typeof value === "boolean") return String(value);
  return JSON.stringify(value);
}

/* ── Address helpers ────────────────────────────────────────── */

/** Extract sheet name from a qualified address, stripping quotes. */
function sheetOf(address: string): string {
  const bang = address.indexOf("!");
  return bang >= 0 ? address.substring(0, bang).replace(/^'|'$/g, "") : "";
}

/** Extract just the cell/range part after the sheet prefix. */
function cellPartOf(address: string): string {
  const bang = address.indexOf("!");
  return bang >= 0 ? address.substring(bang + 1) : address;
}

/* ── Tree rendering ─────────────────────────────────────────── */

/**
 * Render a single tree node and its children recursively.
 *
 * `parentSheet` is the sheet of the direct parent node. When a child
 * is on the same sheet, we show just the cell ref (e.g. "G44") instead
 * of the full "n-Operations!G44" to reduce noise.
 */
function childLabel(mode: TraceDependenciesMode, count: number): string {
  const singular = mode === "dependents" ? "dependent" : "precedent";
  return `${count} ${singular}${count === 1 ? "" : "s"}`;
}

function renderNodeHeader(
  node: DepNodeDetail,
  isRoot: boolean,
  parentSheet: string,
  mode: TraceDependenciesMode,
): TemplateResult {
  const valStr = fmtValue(node.value, node.numberFormat);
  const hasChildren = node.precedents.length > 0;
  const nodeSheet = sheetOf(node.address);
  const sameSheet = !isRoot && nodeSheet !== "" && nodeSheet === parentSheet;

  // Show abbreviated cell ref for same-sheet, full address for cross-sheet / root.
  const addrDisplay = sameSheet
    ? cellRefDisplay(cellPartOf(node.address), node.address)
    : cellRef(node.address);

  return html`
    <div class="pi-dep-node__row">
      ${valStr
        ? html`<span class="pi-dep-node__val ${isExcelError(node.value) ? "pi-dep-node__val--err" : ""}">${valStr}</span>`
        : nothing}
      <span class="pi-dep-node__addr">${addrDisplay}</span>
      ${hasChildren
        ? html`<span class="pi-dep-node__meta">${childLabel(mode, node.precedents.length)}</span>`
        : nothing}
    </div>
    ${node.formula
      ? html`<code class="pi-dep-node__formula" title=${node.formula}>${node.formula}</code>`
      : nothing}
  `;
}

function renderNode(
  node: DepNodeDetail,
  isRoot: boolean,
  parentSheet: string,
  depth: number,
  mode: TraceDependenciesMode,
): TemplateResult {
  const hasFormula = !!node.formula;
  const hasChildren = node.precedents.length > 0;
  const isLeaf = !hasFormula && !hasChildren;
  const nodeSheet = sheetOf(node.address);

  const children = hasChildren
    ? html`<div class="pi-dep-node__children">
        ${node.precedents.map((child) => renderNode(child, false, nodeSheet, depth + 1, mode))}
      </div>`
    : nothing;

  if (!hasChildren) {
    return html`
      <div class="pi-dep-node ${isRoot ? "pi-dep-node--root" : ""} ${isLeaf ? "pi-dep-node--leaf" : ""}">
        ${renderNodeHeader(node, isRoot, parentSheet, mode)}
      </div>
    `;
  }

  if (isRoot) {
    return html`
      <div class="pi-dep-node pi-dep-node--root">
        ${renderNodeHeader(node, true, parentSheet, mode)}
        ${children}
      </div>
    `;
  }

  const openByDefault = depth <= 1;

  return html`
    <details class="pi-dep-node pi-dep-node--branch" ?open=${openByDefault}>
      <summary class="pi-dep-node__summary">
        ${renderNodeHeader(node, false, parentSheet, mode)}
      </summary>
      ${children}
    </details>
  `;
}

/**
 * Render a dependency tree as a visual HTML tree.
 * Used by the trace_dependencies tool result renderer.
 */
export function renderDepTree(
  root: DepNodeDetail,
  mode: TraceDependenciesMode = "precedents",
): TemplateResult {
  return html`<div class="pi-dep-tree">${renderNode(root, true, "", 0, mode)}</div>`;
}
