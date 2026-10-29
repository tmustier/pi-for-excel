import { html, type TemplateResult } from "lit";
import { excelRun, parseRangeRef } from "../excel/helpers.js";
import { t } from "../language/index.js";
import { showToast } from "./toast.js";

const DEBOUNCE_MS = 300;
let lastNavTime = 0;

function firstSubRange(address: string): string {
  const parsed = parseRangeRef(address);
  const first = (parsed.address.split(",")[0] ?? parsed.address).trim();
  return parsed.sheet ? `${parsed.sheet}!${first}` : first;
}

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

    ws.activate();
    await ctx.sync();

    const range = ws.getRange(parsed.address);
    range.select();
    await ctx.sync();
  });
}

export async function followCellLink(address: string): Promise<void> {
  try {
    await navigateToRange(address);
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    showToast(t("cell-link.failed", { address, reason }), { variant: "error" });
  }
}

export function cellRef(address: string): TemplateResult {
  return html`<a
    class="pi-cell-ref"
    href="#"
    @click=${(e: Event) => {
      e.preventDefault();
      e.stopPropagation();
      void followCellLink(address);
    }}
  >${address}</a>`;
}

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
      void followCellLink(fullAddress);
    }}
  >${display}</a>`;
}

export function cellRefs(address: string, maxShow = 8): TemplateResult {
  const parts = address.split(/\s*,\s*/).filter(Boolean);

  if (parts.length <= 1) return cellRef(address);

  const parsed = parts.map((p) => parseRangeRef(p));
  const sheets = [...new Set(parsed.map((p) => p.sheet).filter(Boolean))];
  const commonSheet = sheets.length === 1 ? sheets[0] : undefined;

  const shown = parts.slice(0, maxShow);
  const more = parts.length - maxShow;

  return html`${shown.map((part, i) => {
    const p = parseRangeRef(part);
    const display = commonSheet ? p.address : part;
    const nav = p.sheet
      ? part
      : commonSheet
        ? `${commonSheet}!${p.address}`
        : part;
    return html`${i > 0 ? ", " : ""}${cellRefDisplay(display, nav)}`;
  })}${more > 0 ? html`, <span class="pi-params__more">+${more} more</span>` : ""}`;
}
