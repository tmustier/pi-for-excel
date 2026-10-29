import { buildOverview } from "../tools/get-workbook-overview.js";

const UNKNOWN_WORKBOOK_KEY = "__workbook_unknown__";

const blueprintCacheByWorkbook = new Map<string, string>();

function normalizeWorkbookKey(workbookId: string | null): string {
  return workbookId ?? UNKNOWN_WORKBOOK_KEY;
}

export async function getBlueprint(workbookId: string | null): Promise<string> {
  const workbookKey = normalizeWorkbookKey(workbookId);
  const cached = blueprintCacheByWorkbook.get(workbookKey);
  if (cached !== undefined) return cached;

  const blueprint = await buildOverview({ waitForCellEdit: false });
  blueprintCacheByWorkbook.set(workbookKey, blueprint);
  return blueprint;
}

export function invalidateBlueprint(workbookId: string | null): void {
  if (workbookId === null) {
    blueprintCacheByWorkbook.clear();
    return;
  }
  blueprintCacheByWorkbook.delete(normalizeWorkbookKey(workbookId));
}
