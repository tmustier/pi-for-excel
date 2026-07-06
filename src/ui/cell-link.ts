/**
 * Clickable cell references — navigate to a cell/range in Excel.
 *
 * Uses native selection (activate sheet + select range) which
 * scrolls the viewport and highlights with Excel's blue selection
 * chrome. No fill mutation → no undo pollution, no CF conflicts,
 * no race conditions on rapid clicks.
 */

import { html, type TemplateResult } from "lit";
import { excelRun, parseRangeRef } from "../excel/helpers.js";

/* ── Debounce guard ─────────────────────────────────────────── */

/** Minimum ms between navigation actions. */
const DEBOUNCE_MS = 300;
let lastNavTime = 0;

/* ── Excel navigation ───────────────────────────────────────── */

/**
 * Extract the first sub-range from a comma-separated address.
 *   "D18:H18, D28:H28, D31:H31" → "D18:H18"
 * Keeps any sheet qualifier intact.
 */
function firstSubRange(address: string): string {
  const parsed = parseRangeRef(address);
  const first = (parsed.address.split(",")[0] ?? parsed.address).trim();
  return parsed.sheet ? `${parsed.sheet}!${first}` : first;
}

/**
 * Navigate Excel to the given address and select it.
 *
 * For multi-ranges (comma-separated), navigates to the first
 * sub-range — you can't scroll to disjoint areas simultaneously.
 *
 * Activates the target sheet (switching tabs if needed), then
 * selects the range which scrolls the viewport and applies
 * Excel's native blue selection highlight.
 */
async function navigateToRange(address: string): Promise<void> {
  const now = Date.now();
  if (now - lastNavTime < DEBOUNCE_MS) return;
  lastNavTime = now;

  const target = firstSubRange(address);
  const parsed = parseRangeRef(target);

  await excelRun(async (ctx) => {
    const ws = parsed.sheet
      ? ctx.workbook.worksheets.getItem(parsed.sheet)
      : ctx.workbook.worksheets.getActiveWorksheet();

    // Activate sheet first (switches tab, ensures correct context)
    ws.activate();
    await ctx.sync();

    // Select range (scrolls viewport + native highlight)
    const range = ws.getRange(parsed.address);
    range.select();
    await ctx.sync();
  });
}

/* ── Lit template helper ────────────────────────────────────── */

/**
 * Render a cell/range address as a clickable link that navigates
 * Excel to that location.
 *
 * Usage in Lit templates:
 *   html`Range: ${cellRef("Sheet1!A1:B10")}`
 */
export function cellRef(address: string): TemplateResult {
  return html`<a
    class="pi-cell-ref"
    href="#"
    @click=${(e: Event) => {
      e.preventDefault();
      e.stopPropagation();
      void navigateToRange(address);
    }}
  >${address}</a>`;
}

/**
 * Render a range display value (possibly a TemplateResult with
 * "+N more" suffix) as a clickable cell ref. Falls back to plain
 * text for non-string values (e.g. already-rendered templates).
 */
export function cellRefDisplay(
  display: TemplateResult | string,
  fullAddress: string,
): TemplateResult {
  return html`<a
    class="pi-cell-ref"
    href="#"
    @click=${(e: Event) => {
      e.preventDefault();
      e.stopPropagation();
      void navigateToRange(fullAddress);
    }}
  >${display}</a>`;
}

/**
 * Render a potentially comma-separated range string as individually
 * clickable links. Factors out a common sheet prefix and truncates
 * after `maxShow` items.
 *
 *   "Summary!C10, Summary!C12, Summary!C14"
 *   → [C10] , [C12] , [C14]   (each clickable, sheet shown separately)
 *
 *   "C10, C12"
 *   → [C10] , [C12]
 *
 * For a single range (no commas), delegates to `cellRef()`.
 */
export function cellRefs(address: string, maxShow = 8): TemplateResult {
  const parts = address.split(/\s*,\s*/).filter(Boolean);

  if (parts.length <= 1) return cellRef(address);

  // Find common sheet prefix
  const parsed = parts.map((p) => parseRangeRef(p));
  const sheets = [...new Set(parsed.map((p) => p.sheet).filter(Boolean))];
  const commonSheet = sheets.length === 1 ? sheets[0] : undefined;

  const shown = parts.slice(0, maxShow);
  const more = parts.length - maxShow;

  return html`${shown.map((part, i) => {
    const p = parseRangeRef(part);
    // Display: strip shared sheet prefix for readability
    const display = commonSheet ? p.address : part;
    // Navigation: ensure each part has the full sheet-qualified address
    const nav = p.sheet
      ? part
      : commonSheet
        ? `${commonSheet}!${p.address}`
        : part;
    return html`${i > 0 ? ", " : ""}${cellRefDisplay(display, nav)}`;
  })}${more > 0 ? html`, <span class="pi-params__more">+${more} more</span>` : ""}`;
}
