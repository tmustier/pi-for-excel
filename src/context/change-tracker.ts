/**
 * Change tracker — accumulates user edits between agent messages.
 *
 * Registers Worksheet.onChanged event handlers and collects a log of
 * cell changes. The log is flushed and injected into context on each
 * user message.
 */

import { excelRun } from "../excel/helpers.js";

interface CellChange {
  sheet: string;
  address: string;
  timestamp: number;
}

/** Maximum changes to accumulate before auto-truncating */
const MAX_CHANGES = 50;

export class ChangeTracker {
  private changes: CellChange[] = [];
  private registered = false;
  private sheetIdToName = new Map<string, string>();

  /** Start listening for changes on all visible sheets. */
  async start(): Promise<void> {
    if (this.registered) return;

    try {
      await excelRun(async (context: any) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/id,items/name,items/visibility");
        await context.sync();

        this.sheetIdToName.clear();
        for (const sheet of sheets.items) {
          this.sheetIdToName.set(sheet.id, sheet.name);
        }

        for (const sheet of sheets.items) {
          if (sheet.visibility !== "Visible") continue;

          sheet.onChanged.add((event: any) => {
            const sheetName = this.sheetIdToName.get(event.worksheetId) || event.worksheetId || "unknown";
            this.changes.push({
              sheet: sheetName,
              address: event.address || "unknown",
              timestamp: Date.now(),
            });

            // Auto-truncate to prevent unbounded memory use
            if (this.changes.length > MAX_CHANGES) {
              this.changes = this.changes.slice(-MAX_CHANGES);
            }
          });
        }
        await context.sync();
      });
      this.registered = true;
      console.log("[change-tracker] Listening for changes");
    } catch (e: any) {
      console.warn("[change-tracker] Failed to register:", e.message);
    }
  }

  /** Flush accumulated changes and return them as a context string. Returns null if no changes. */
  flush(): string | null {
    if (this.changes.length === 0) return null;

    // Deduplicate by sheet+address (keep last change per cell)
    const byCell = new Map<string, CellChange>();
    for (const c of this.changes) {
      byCell.set(`${c.sheet}!${c.address}`, c);
    }

    const unique = [...byCell.values()];
    this.changes = [];

    // Group by sheet
    const bySheet = new Map<string, string[]>();
    for (const c of unique) {
      const existing = bySheet.get(c.sheet) || [];
      existing.push(c.address);
      bySheet.set(c.sheet, existing);
    }

    const parts: string[] = [];
    for (const [sheet, addresses] of bySheet) {
      const display = addresses.length > 10
        ? addresses.slice(0, 10).join(", ") + `, … (+${addresses.length - 10} more)`
        : addresses.join(", ");
      parts.push(`${sheet}: ${display}`);
    }

    return `**User changes since last message:** ${parts.join("; ")}`;
  }

  /** Check if there are pending changes. */
  get hasPendingChanges(): boolean {
    return this.changes.length > 0;
  }
}
