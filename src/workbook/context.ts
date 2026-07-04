/**
 * Workbook context primitives.
 *
 * Compatibility wrapper around the current SpreadsheetHost. The hashing and
 * label helpers live in `src/host/workbook-context.ts` so Office and future WPS
 * adapters share the same privacy-preserving identity semantics.
 */

import { getCurrentSpreadsheetHost } from "../host/current.js";
import type { WorkbookContext } from "../host/workbook-context.js";

export type { WorkbookContext } from "../host/workbook-context.js";
export { formatWorkbookLabel } from "../host/workbook-context.js";

/**
 * Best-effort workbook context.
 *
 * IMPORTANT: callers should persist only `workbookId` (the hash), never a raw
 * host document URL/path.
 */
export async function getWorkbookContext(): Promise<WorkbookContext> {
  return getCurrentSpreadsheetHost().getWorkbookContext();
}
