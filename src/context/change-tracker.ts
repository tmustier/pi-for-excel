/**
 * Change tracker — edits made in the workbook by someone other than Pi.
 *
 * One workbook-level listener per taskpane (`worksheets.onChanged`, so sheets
 * added after the taskpane started are covered) appends to a bounded log with
 * sequence numbers. Each chat runtime keeps its own cursor into the log, so
 * every tab hears about every edit once, independently of the others.
 *
 * Pi's own writes (`triggerSource: "ThisLocalAddin"`) are not logged: the
 * agent already knows about them from its tool results.
 *
 * Structural events (sheet added, deleted, renamed, moved, shown/hidden, and
 * row/column inserts and deletes) also call `onStructureChange` so the
 * workbook blueprint can be rebuilt.
 */

import { excelRun } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

interface LoggedChange {
  seq: number;
  sheet: string;
  address: string;
  /** Office.js `changeType` when it is not a plain cell edit. */
  structural: string | null;
  remote: boolean;
  /** "before → after" for a single-cell edit when Excel reports it. */
  values: string | null;
}

/** Log entries kept in memory; older entries are summarised as dropped. */
const MAX_LOG_ENTRIES = 500;
const MAX_SHEETS_LISTED = 10;
const MAX_ADDRESSES_PER_SHEET = 20;
const MAX_VALUE_CHARS = 40;

const STRUCTURAL_CHANGE_LABELS: Readonly<Record<string, string>> = {
  RowInserted: "rows inserted",
  RowDeleted: "rows deleted",
  ColumnInserted: "columns inserted",
  ColumnDeleted: "columns deleted",
  CellInserted: "cells inserted",
  CellDeleted: "cells deleted",
};

/** A cell value as Office.js reports it in change details. */
type CellValue = string | number | boolean | null | undefined;

function formatValue(value: CellValue): string {
  if (value === null || value === undefined || value === "") return "(empty)";
  const text = typeof value === "string" ? JSON.stringify(value) : String(value);
  return text.length > MAX_VALUE_CHARS ? `${text.slice(0, MAX_VALUE_CHARS - 1)}…` : text;
}

function describeValues(details: Excel.ChangedEventDetail | undefined): string | null {
  if (!details) return null;
  // Office.js types these as `any`: keep only scalar cell values.
  const rawBefore: unknown = details.valueBefore;
  const rawAfter: unknown = details.valueAfter;
  const valueBefore: CellValue = typeof rawBefore === "string" || typeof rawBefore === "number" || typeof rawBefore === "boolean" ? rawBefore : undefined;
  const valueAfter: CellValue = typeof rawAfter === "string" || typeof rawAfter === "number" || typeof rawAfter === "boolean" ? rawAfter : undefined;
  if (valueBefore === undefined && valueAfter === undefined) return null;
  return `${formatValue(valueBefore)} → ${formatValue(valueAfter)}`;
}

export interface ChangeSummary {
  /** Agent-facing text, or null when nothing changed since the cursor. */
  text: string | null;
  /** Cursor to pass next time. */
  cursor: number;
}

export class ChangeTracker {
  private log: LoggedChange[] = [];
  private seq = 0;
  private started = false;
  private readonly sheetNames = new Map<string, string>();
  private readonly onStructureChange: () => void;

  constructor(options: { onStructureChange?: () => void } = {}) {
    this.onStructureChange = options.onStructureChange ?? (() => {});
  }

  /** Start listening on the whole workbook. Safe to call more than once. */
  async start(): Promise<void> {
    if (this.started) return;
    this.started = true;

    try {
      await excelRun(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/id,items/name");
        await context.sync();
        for (const sheet of sheets.items) this.sheetNames.set(sheet.id, sheet.name);

        sheets.onChanged.add((event) => this.handleChanged(event));
        sheets.onAdded.add((event) => this.handleSheetAdded(event.worksheetId));
        sheets.onDeleted.add(() => this.structureChanged());
        await context.sync();
      });
    } catch (error) {
      this.started = false;
      console.warn("[change-tracker] Failed to register workbook listeners:", getErrorMessage(error));
      return;
    }

    // ExcelApi 1.17: renames, moves and visibility changes. Optional.
    try {
      await excelRun(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.onNameChanged.add((event) => {
          this.sheetNames.set(event.worksheetId, event.nameAfter);
          return this.structureChanged();
        });
        sheets.onMoved.add(() => this.structureChanged());
        sheets.onVisibilityChanged.add(() => this.structureChanged());
        await context.sync();
      });
    } catch (error) {
      console.warn("[change-tracker] Sheet rename/move/visibility events unavailable:", getErrorMessage(error));
    }
  }

  /** Current end of the log. A new runtime starts here. */
  cursor(): number {
    return this.seq;
  }

  /** Describe edits logged after `cursor`, for the agent. */
  describeSince(cursor: number): ChangeSummary {
    const entries = this.log.filter((entry) => entry.seq > cursor);
    if (entries.length === 0) return { text: null, cursor: this.seq };

    const firstKept = this.log[0]?.seq ?? this.seq + 1;
    const dropped = firstKept > cursor + 1;

    // Latest entry per sheet + address wins; keep first-seen order.
    const bySheet = new Map<string, Map<string, LoggedChange>>();
    for (const entry of entries) {
      const sheetEntries = bySheet.get(entry.sheet) ?? new Map<string, LoggedChange>();
      sheetEntries.delete(entry.address);
      sheetEntries.set(entry.address, entry);
      bySheet.set(entry.sheet, sheetEntries);
    }

    const lines: string[] = [
      "Edits made in Excel since your previous request, by the user or a collaborator (Pi's own changes are not listed):",
    ];
    if (dropped) lines.push("- (earlier edits were not kept)");

    let sheetCount = 0;
    for (const [sheet, sheetEntries] of bySheet) {
      sheetCount += 1;
      if (sheetCount > MAX_SHEETS_LISTED) {
        lines.push(`- … and ${bySheet.size - MAX_SHEETS_LISTED} more sheet(s)`);
        break;
      }
      const items = [...sheetEntries.values()].map((entry) => {
        const notes: string[] = [];
        if (entry.structural) notes.push(entry.structural);
        if (entry.values) notes.push(entry.values);
        if (entry.remote) notes.push("by a collaborator");
        return notes.length > 0 ? `${entry.address} (${notes.join(", ")})` : entry.address;
      });
      const shown = items.slice(0, MAX_ADDRESSES_PER_SHEET);
      const more = items.length - shown.length;
      lines.push(`- ${sheet}: ${shown.join(", ")}${more > 0 ? `, … (+${more} more)` : ""}`);
    }

    return { text: lines.join("\n"), cursor: this.seq };
  }

  private handleChanged(event: Excel.WorksheetChangedEventArgs): Promise<void> {
    const structural = STRUCTURAL_CHANGE_LABELS[event.changeType] ?? null;
    if (structural) this.onStructureChange();
    if (event.triggerSource === "ThisLocalAddin") return Promise.resolve();

    this.seq += 1;
    this.log.push({
      seq: this.seq,
      sheet: this.sheetNames.get(event.worksheetId) ?? event.worksheetId,
      address: event.address || "unknown",
      structural,
      remote: event.source === "Remote",
      values: structural ? null : describeValues(event.details),
    });
    if (this.log.length > MAX_LOG_ENTRIES) this.log.splice(0, this.log.length - MAX_LOG_ENTRIES);
    return Promise.resolve();
  }

  private async handleSheetAdded(worksheetId: string): Promise<void> {
    this.onStructureChange();
    try {
      await excelRun(async (context) => {
        const sheet = context.workbook.worksheets.getItem(worksheetId);
        sheet.load("name");
        await context.sync();
        this.sheetNames.set(worksheetId, sheet.name);
      });
    } catch (error) {
      console.warn("[change-tracker] Could not read the added sheet's name:", getErrorMessage(error));
    }
  }

  /** Event handlers must return a promise. */
  private structureChanged(): Promise<void> {
    this.onStructureChange();
    return Promise.resolve();
  }
}
