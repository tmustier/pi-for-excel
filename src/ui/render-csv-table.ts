/**
 * Visual CSV table renderer for read_range CSV results.
 *
 * Replaces the syntax-highlighted code block with a proper HTML table
 * featuring Excel-style row/column headers and a "Copy CSV" button.
 * The raw CSV text sent to the agent is unchanged.
 */

import { html, nothing, type TemplateResult } from "lit";
import { colToLetter } from "../excel/helpers.js";
import { t } from "../language/index.js";
import type { ReadRangeCsvDetails } from "../tools/tool-details.js";
import { isExcelError } from "../utils/format.js";

/* ── Value formatting ───────────────────────────────────────── */

function fmtCell(v: DynamicValue): string {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "string") return v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  return JSON.stringify(v);
}

/* ── Copy-to-clipboard ──────────────────────────────────────── */

async function copyToClipboard(csv: string, btn: HTMLButtonElement): Promise<void> {
  try {
    await navigator.clipboard.writeText(csv);
    const orig = btn.textContent;
    btn.textContent = t("render-csv-table.copied");
    btn.classList.add("pi-csv-table__copy--done");
    setTimeout(() => {
      btn.textContent = orig;
      btn.classList.remove("pi-csv-table__copy--done");
    }, 1500);
  } catch {
    // Fallback: select-all in a temp textarea
    const ta = document.createElement("textarea");
    ta.value = csv;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
  }
}

/* ── Table rendering ────────────────────────────────────────── */

/**
 * Render a CSV result as an Excel-style table with row/column headers
 * and a copy button.
 */
export function renderCsvTable(details: ReadRangeCsvDetails): TemplateResult {
  const { startCol, startRow, values, csv } = details;
  if (values.length === 0) return html`<div class="pi-tool-card__plain-text">(empty)</div>`;

  const numCols = Math.max(...values.map((r) => (Array.isArray(r) ? r.length : 0)));

  return html`
    <div class="pi-csv-table">
      <div class="pi-csv-table__toolbar">
        <button
          class="pi-csv-table__copy"
          @click=${(e: Event) => {
            const btn = e.currentTarget;
            if (btn instanceof HTMLButtonElement) void copyToClipboard(csv, btn);
          }}
        >Copy CSV</button>
      </div>
      <div class="pi-csv-table__scroll">
        <table class="pi-csv-table__grid">
          <thead>
            <tr>
              <th class="pi-csv-table__corner"></th>
              ${Array.from({ length: numCols }, (_, i) =>
                html`<th class="pi-csv-table__col-hdr">${colToLetter(startCol + i)}</th>`,
              )}
            </tr>
          </thead>
          <tbody>
            ${values.map((row, r) => html`
              <tr>
                <th class="pi-csv-table__row-hdr">${startRow + r}</th>
                ${Array.from({ length: numCols }, (_, c) => {
                  const val = Array.isArray(row) ? row[c] : undefined;
                  const display = fmtCell(val);
                  const isNum = typeof val === "number";
                  const isErr = isExcelError(val);
                  return html`<td class="${isNum ? "pi-csv-table__num" : ""} ${isErr ? "pi-csv-table__err" : ""}">${display || nothing}</td>`;
                })}
              </tr>
            `)}
          </tbody>
        </table>
      </div>
    </div>
  `;
}
