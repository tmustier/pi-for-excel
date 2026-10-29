import { getCurrentSpreadsheetHost } from "../host/current.js";
import type { WorkbookContext } from "../host/workbook-context.js";

export type { WorkbookContext } from "../host/workbook-context.js";
export { formatWorkbookLabel } from "../host/workbook-context.js";

export async function getWorkbookContext(): Promise<WorkbookContext> {
  return getCurrentSpreadsheetHost().getWorkbookContext();
}
