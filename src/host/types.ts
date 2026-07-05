import type { WorkbookContext } from "../workbook/context.js";

export type SpreadsheetHostKind = "office" | "wps" | "browser";

export interface SpreadsheetHostReadyInfo {
  kind: SpreadsheetHostKind;
  host?: string;
  platform?: string;
}

export interface SpreadsheetHost {
  kind: SpreadsheetHostKind;
  ready: () => Promise<SpreadsheetHostReadyInfo>;
  getWorkbookContext: () => Promise<WorkbookContext>;
}
