function isTaskpaneTabLayoutPayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Persistence for open session tabs per workbook.
 *
 * Stores tab layout in SettingsStore so reloading the taskpane restores
 * the same set of session tabs, their "Chat N" numbers and the active tab.
 * A session appears in the layout at most once: one session, one tab.
 */

import { Type } from "typebox";
import { Value } from "typebox/value";

import type { SettingsStore } from "../storage/local/settings-store.js";


const WORKBOOK_TAB_LAYOUT_PREFIX = "workbook.tabLayout.v1.";
const GLOBAL_WORKBOOK_LAYOUT_KEY = "__global__";

export interface WorkbookTabLayout {
  sessionIds: string[];
  activeSessionId: string | null;
  /** Default "Chat N" number per open session. Absent in layouts saved before numbers were kept. */
  tabNumbers?: Record<string, number>;
  /** Next number a new tab gets, so numbers of closed tabs are not reused after a reload. */
  nextTabNumber?: number;
}

const tabNumbersSchema = Type.Record(Type.String(), Type.Integer({ minimum: 1 }));
const tabNumberSchema = Type.Integer({ minimum: 1 });

function normalizeSessionId(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function isSessionIdList(value: unknown): value is unknown[] {
  return Array.isArray(value);
}

export function workbookTabLayoutKey(workbookId: string | null): string {
  const suffix = workbookId && workbookId.trim().length > 0
    ? workbookId
    : GLOBAL_WORKBOOK_LAYOUT_KEY;

  return `${WORKBOOK_TAB_LAYOUT_PREFIX}${suffix}`;
}

export function normalizeWorkbookTabLayout(layout: WorkbookTabLayout): WorkbookTabLayout {
  const sessionIds: string[] = [];

  for (const rawSessionId of layout.sessionIds) {
    const normalized = normalizeSessionId(rawSessionId);
    // Older builds could list one session twice; it restores as one tab.
    if (!normalized || sessionIds.includes(normalized)) continue;
    sessionIds.push(normalized);
  }

  const normalizedActiveSessionId = normalizeSessionId(layout.activeSessionId);
  const activeSessionId = normalizedActiveSessionId && sessionIds.includes(normalizedActiveSessionId)
    ? normalizedActiveSessionId
    : (sessionIds[0] ?? null);

  const normalized: WorkbookTabLayout = { sessionIds, activeSessionId };

  if (layout.tabNumbers) {
    const tabNumbers: Record<string, number> = {};
    const usedNumbers = new Set<number>();
    for (const sessionId of sessionIds) {
      const tabNumber = layout.tabNumbers[sessionId];
      if (tabNumber === undefined || usedNumbers.has(tabNumber)) continue;
      usedNumbers.add(tabNumber);
      tabNumbers[sessionId] = tabNumber;
    }
    normalized.tabNumbers = tabNumbers;
  }

  if (layout.nextTabNumber !== undefined) {
    normalized.nextTabNumber = layout.nextTabNumber;
  }

  return normalized;
}

export function parseWorkbookTabLayout(value: unknown): WorkbookTabLayout | null {
  if (!isTaskpaneTabLayoutPayloadShape(value)) return null;

  const rawSessionIds = value.sessionIds;
  if (!isSessionIdList(rawSessionIds)) return null;

  const parsed: WorkbookTabLayout = {
    sessionIds: rawSessionIds
      .map((sessionId) => normalizeSessionId(sessionId))
      .filter((sessionId): sessionId is string => sessionId !== null),
    activeSessionId: normalizeSessionId(value.activeSessionId),
  };

  // Malformed numbers fall back to fresh numbering rather than dropping the layout.
  if (Value.Check(tabNumbersSchema, value.tabNumbers)) {
    parsed.tabNumbers = value.tabNumbers;
  }
  if (Value.Check(tabNumberSchema, value.nextTabNumber)) {
    parsed.nextTabNumber = value.nextTabNumber;
  }

  const normalized = normalizeWorkbookTabLayout(parsed);
  if (normalized.sessionIds.length === 0) return null;

  return normalized;
}

export async function loadWorkbookTabLayout(
  settings: SettingsStore,
  workbookId: string | null,
): Promise<WorkbookTabLayout | null> {
  const stored = await settings.get(workbookTabLayoutKey(workbookId));
  return parseWorkbookTabLayout(stored);
}

export async function saveWorkbookTabLayout(
  settings: SettingsStore,
  workbookId: string | null,
  layout: WorkbookTabLayout,
): Promise<void> {
  const normalized = normalizeWorkbookTabLayout(layout);
  const key = workbookTabLayoutKey(workbookId);

  if (normalized.sessionIds.length === 0) {
    await settings.delete(key);
    return;
  }

  await settings.set(key, normalized);
}
