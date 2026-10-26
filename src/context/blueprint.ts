/**
 * Workbook blueprint — structural overview sent to the model as an
 * auto-context note (see `src/taskpane/context-injection.ts`).
 *
 * Re-uses buildOverview() from the get_workbook_overview tool, cached per
 * workbook. The cache is invalidated by structural changes: Pi's own
 * structure writes (workbook coordinator) and the user's (change tracker
 * sheet/row/column events). The injector compares the blueprint with the last
 * one stored in the transcript and sends a new note only when it differs.
 */

import { buildOverview } from "../tools/get-workbook-overview.js";

const UNKNOWN_WORKBOOK_KEY = "__workbook_unknown__";

const blueprintCacheByWorkbook = new Map<string, string>();

function normalizeWorkbookKey(workbookId: string | null): string {
  return workbookId ?? UNKNOWN_WORKBOOK_KEY;
}

/**
 * Get the workbook blueprint (cached per workbook identity). Never waits for
 * the user to leave cell-edit mode: it runs before a model request.
 */
export async function getBlueprint(workbookId: string | null): Promise<string> {
  const workbookKey = normalizeWorkbookKey(workbookId);
  const cached = blueprintCacheByWorkbook.get(workbookKey);
  if (cached !== undefined) return cached;

  const blueprint = await buildOverview({ waitForCellEdit: false });
  blueprintCacheByWorkbook.set(workbookKey, blueprint);
  return blueprint;
}

/** Invalidate one workbook's cached blueprint (null: every workbook). */
export function invalidateBlueprint(workbookId: string | null): void {
  if (workbookId === null) {
    blueprintCacheByWorkbook.clear();
    return;
  }
  blueprintCacheByWorkbook.delete(normalizeWorkbookKey(workbookId));
}
