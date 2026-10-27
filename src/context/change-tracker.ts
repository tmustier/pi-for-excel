/**
 * Change tracker — edits made in the workbook by someone other than Pi.
 *
 * One workbook-level listener per taskpane (`worksheets.onChanged`, so sheets
 * added after the taskpane started are covered) appends to a bounded log with
 * sequence numbers. Each chat runtime keeps its own cursor into the log, so
 * every tab hears about every edit once, independently of the others.
 *
 * Who made an edit. Every write Pi makes (any tab's tool call, or a backup
 * restore the user started) runs under the workbook lock and is recorded here
 * as a write window, so at most one is running at a time.
 * - `source: "Remote"` is a co-author (someone else editing the shared
 *   workbook); reported as theirs.
 * - `triggerSource: "ThisLocalAddin"` is Pi. It is matched to the write window
 *   it arrived in by time (and by range when several match): this runtime's own writes are left out (the agent
 *   knows about them); a restore is reported as "restored from a backup",
 *   another tab's write as "by Pi in another tab", and anything unmatched
 *   (an extension, say) as "by Pi outside this chat".
 * - Once a co-author has joined the workbook, Pi's writes arrive as
 *   `"Unknown"`, like the user's (and stay so after the co-author leaves). A
 *   local "Unknown" edit is matched the same way, but only when it also
 *   touches the range that write recorded (any cell when unknown); otherwise
 *   it is the user's. Matching runs when the edits are described, by which
 *   time the write is recorded.
 *
 * Structural events (sheet added, deleted, renamed, moved, shown/hidden, and
 * row/column inserts and deletes, local or remote) also call
 * `onStructureChange` so the workbook blueprint is rebuilt.
 */

import { excelRun, letterToCol, parseRangeRef } from "../excel/helpers.js";
import type { WorkbookWriteWindow } from "../tools/with-workbook-coordinator.js";
import { getErrorMessage } from "../utils/errors.js";

interface LoggedChange {
  seq: number;
  at: number;
  sheet: string;
  address: string;
  /** Office.js `changeType` when it is not a plain cell edit. */
  structural: string | null;
  remote: boolean;
  /** `triggerSource: "ThisLocalAddin"`: certainly made by Pi. */
  addin: boolean;
  /** "before → after" for a single-cell edit when Excel reports it. */
  values: string | null;
}

/** Log entries kept in memory; older entries are summarised as dropped. */
const MAX_LOG_ENTRIES = 500;
const MAX_SHEETS_LISTED = 10;
const MAX_ADDRESSES_PER_SHEET = 20;
const MAX_VALUE_CHARS = 40;
/**
 * Change events for Pi's write can arrive after the call returned. Edits
 * known to come from the add-in get a longer grace: no user edit can be
 * mistaken for them.
 */
const PI_WRITE_EVENT_GRACE_MS = 3_000;
const ADDIN_EVENT_GRACE_MS = 15_000;
const MAX_WRITE_WINDOWS = 200;

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
  const before = formatValue(valueBefore);
  const after = formatValue(valueAfter);
  // Remote edits often report the new value as both.
  return before === after ? `now ${after}` : `${before} → ${after}`;
}

interface Area {
  sheet: string | undefined;
  rows: [number, number];
  cols: [number, number];
}

const UNBOUNDED: [number, number] = [0, Number.MAX_SAFE_INTEGER];

/** "Sheet!A1:B2", "A:C", "3:5", comma-separated lists → areas; null if unparsable. */
function parseAreas(address: string): Area[] | null {
  const areas: Area[] = [];
  for (const part of address.split(",")) {
    const { sheet, address: local } = parseRangeRef(part.trim());
    const [from = "", to = from] = local.replace(/\$/g, "").toUpperCase().split(":");
    const a = /^([A-Z]*)(\d*)$/.exec(from);
    const b = /^([A-Z]*)(\d*)$/.exec(to);
    if (!a || !b || (a[1] === "" && a[2] === "")) return null;
    const rows: [number, number] = a[2] && b[2] ? [Number(a[2]), Number(b[2])] : UNBOUNDED;
    const cols: [number, number] = a[1] && b[1] ? [letterToCol(a[1]), letterToCol(b[1])] : UNBOUNDED;
    areas.push({ sheet, rows, cols });
  }
  return areas;
}

function overlaps(change: LoggedChange, written: readonly string[]): boolean {
  const changed = parseAreas(change.address);
  if (!changed) return true;
  for (const address of written) {
    const areas = parseAreas(address);
    if (!areas) return true;
    for (const w of areas) {
      if (w.sheet !== undefined && w.sheet !== change.sheet) continue;
      for (const c of changed) {
        if (c.rows[0] <= w.rows[1] && w.rows[0] <= c.rows[1] && c.cols[0] <= w.cols[1] && w.cols[0] <= c.cols[1]) return true;
      }
    }
  }
  return false;
}

/** Who made a write Pi did: a chat runtime (tab) or a backup restore by the user. */
export type PiWriteOrigin = { kind: "runtime"; runtimeId: string } | { kind: "restore" };

/** A write window, optionally with the range it wrote when known up front. */
export type PiWriteWindow = Pick<WorkbookWriteWindow, "toolCallId" | "startedAt" | "settledAt"> & { address?: string };

interface PiWrite {
  window: PiWriteWindow;
  origin: PiWriteOrigin;
}

type Author = "user" | "coAuthor" | "restore" | "otherTab" | "addin";

const GROUP_HEADINGS: ReadonlyArray<[Author, string]> = [
  ["user", "By the user:"],
  ["restore", "Restored from a backup by the user (undoing earlier changes by Pi; cells hold their earlier values again):"],
  ["otherTab", "By Pi in another chat tab:"],
  ["addin", "By Pi outside this chat (an extension or another tab):"],
  ["coAuthor", "By a co-author (someone else editing this workbook at the same time):"],
];

export interface ChangeSummary {
  /** Agent-facing text, or null when nothing changed since the cursor. */
  text: string | null;
  /** Cursor to pass next time. */
  cursor: number;
}

export interface ChangeTrackerOptions {
  onStructureChange?: () => void;
  /**
   * Ranges each of Pi's tool calls wrote, by tool call id (from the change
   * audit). A call with no entry is treated as touching any cell.
   */
  getPiWriteAddresses?: (toolCallIds: readonly string[]) => Promise<ReadonlyMap<string, readonly string[]>>;
}

export class ChangeTracker {
  private log: LoggedChange[] = [];
  private seq = 0;
  private started = false;
  private readonly sheetNames = new Map<string, string>();
  private readonly writes: PiWrite[] = [];
  private readonly onStructureChange: () => void;
  private readonly getPiWriteAddresses: NonNullable<ChangeTrackerOptions["getPiWriteAddresses"]>;

  constructor(options: ChangeTrackerOptions = {}) {
    this.onStructureChange = options.onStructureChange ?? (() => {});
    this.getPiWriteAddresses = options.getPiWriteAddresses ?? (() => Promise.resolve(new Map()));
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

  /**
   * Record a write Pi is making: a tool call in any tab, or a restore. The
   * caller sets `settledAt` (and may set `address`) on the same object later.
   */
  notePiWrite(window: PiWriteWindow, origin: PiWriteOrigin): void {
    this.writes.push({ window, origin });
    if (this.writes.length > MAX_WRITE_WINDOWS) this.writes.shift();
  }

  /** Current end of the log. A new runtime starts here. */
  cursor(): number {
    return this.seq;
  }

  /**
   * Describe edits logged after `cursor` that `runtimeId` did not make itself,
   * for that runtime's agent.
   */
  async describeSince(cursor: number, runtimeId: string): Promise<ChangeSummary> {
    const end = this.seq;
    const pending = this.log.filter((entry) => entry.seq > cursor && entry.seq <= end);
    if (pending.length === 0) return { text: null, cursor: end };

    const firstKept = this.log[0]?.seq ?? end + 1;
    const dropped = firstKept > cursor + 1;
    // One line per cell: its latest change says who left it in its current state.
    const all = await this.attribute(pending, runtimeId);
    const key = ({ entry }: { entry: LoggedChange }) => `${entry.sheet}\u0000${entry.address}`;
    const latest = new Map(all.map((item, index) => [key(item), index]));
    const attributed = all.filter((item, index) => latest.get(key(item)) === index);
    if (attributed.length === 0 && !dropped) return { text: null, cursor: end };

    const lines = ["Edits made in Excel since your previous request (your own changes are not listed):"];
    if (dropped) lines.push("- (earlier edits were not kept)");
    for (const [author, heading] of GROUP_HEADINGS) {
      const group = this.describeGroup(attributed.filter((item) => item.author === author).map((item) => item.entry));
      if (group.length > 0) lines.push(heading, ...group);
    }
    const shiftedSheets = [...new Set(attributed
      .filter((item) => item.author === "coAuthor" && item.entry.structural)
      .map((item) => item.entry.sheet))];
    if (shiftedSheets.length > 0) {
      lines.push(
        `A co-author inserted or deleted cells on ${shiftedSheets.join(", ")}: cells there may have moved, so addresses you read or wrote earlier may now point elsewhere. Read them again before relying on them.`,
      );
    }
    return { text: lines.join("\n"), cursor: end };
  }

  /** Who made each entry; entries `runtimeId` made itself are left out. */
  private async attribute(
    entries: readonly LoggedChange[],
    runtimeId: string,
  ): Promise<Array<{ entry: LoggedChange; author: Author }>> {
    // Writes are serialized and recorded in start order, so the latest one
    // that covers the entry's time is the one that was running.
    const writesAt = (entry: LoggedChange): PiWrite[] => {
      const grace = entry.addin ? ADDIN_EVENT_GRACE_MS : PI_WRITE_EVENT_GRACE_MS;
      return this.writes
        .filter(({ window }) => entry.at >= window.startedAt && entry.at <= window.settledAt + grace)
        .reverse();
    };

    // Ranges each candidate write touched: an "Unknown" local edit is Pi's
    // only if it touches one; for an add-in edit they pick between writes.
    const unknownIds = [...new Set(entries
      .filter((entry) => !entry.remote)
      .flatMap((entry) => writesAt(entry).filter(({ window }) => window.address === undefined).map(({ window }) => window.toolCallId)))];
    let written: ReadonlyMap<string, readonly string[]> = new Map();
    if (unknownIds.length > 0) {
      try {
        written = await this.getPiWriteAddresses(unknownIds);
      } catch (error) {
        console.warn("[change-tracker] Could not read Pi's written ranges:", getErrorMessage(error));
      }
    }
    const touches = (entry: LoggedChange, { window }: PiWrite): boolean => {
      const addresses = window.address !== undefined ? [window.address] : written.get(window.toolCallId);
      return !addresses || addresses.length === 0 || overlaps(entry, addresses);
    };

    const result: Array<{ entry: LoggedChange; author: Author }> = [];
    for (const entry of entries) {
      if (entry.remote) {
        result.push({ entry, author: "coAuthor" });
        continue;
      }
      const candidates = writesAt(entry);
      const write = candidates.find((candidate) => touches(entry, candidate)) ?? (entry.addin ? candidates[0] : undefined);
      if (!write) {
        result.push({ entry, author: entry.addin ? "addin" : "user" });
        continue;
      }
      if (write.origin.kind === "restore") result.push({ entry, author: "restore" });
      else if (write.origin.runtimeId !== runtimeId) result.push({ entry, author: "otherTab" });
    }
    return result;
  }

  private describeGroup(entries: readonly LoggedChange[]): string[] {
    // Latest entry per sheet + address wins; keep first-seen order.
    const bySheet = new Map<string, Map<string, LoggedChange>>();
    for (const entry of entries) {
      const sheetEntries = bySheet.get(entry.sheet) ?? new Map<string, LoggedChange>();
      sheetEntries.delete(entry.address);
      sheetEntries.set(entry.address, entry);
      bySheet.set(entry.sheet, sheetEntries);
    }

    const lines: string[] = [];
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
        return notes.length > 0 ? `${entry.address} (${notes.join(", ")})` : entry.address;
      });
      const shown = items.slice(0, MAX_ADDRESSES_PER_SHEET);
      const more = items.length - shown.length;
      lines.push(`- ${sheet}: ${shown.join(", ")}${more > 0 ? `, … (+${more} more)` : ""}`);
    }
    return lines;
  }

  private handleChanged(event: Excel.WorksheetChangedEventArgs): Promise<void> {
    const structural = STRUCTURAL_CHANGE_LABELS[event.changeType] ?? null;
    if (structural) this.onStructureChange();

    this.seq += 1;
    this.log.push({
      seq: this.seq,
      at: Date.now(),
      sheet: this.sheetNames.get(event.worksheetId) ?? event.worksheetId,
      address: event.address || "unknown",
      structural,
      remote: event.source === "Remote",
      addin: event.source !== "Remote" && event.triggerSource === "ThisLocalAddin",
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
