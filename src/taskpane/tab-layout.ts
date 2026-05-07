/**
 * Persistence for open session tabs per workbook.
 *
 * Stores tab layout in SettingsStore so reloading the taskpane restores
 * the same set of session tabs and active tab.
 */

import type { SettingsStore } from "@earendil-works/pi-web-ui";

import { isRecord } from "../utils/type-guards.js";

const WORKBOOK_TAB_LAYOUT_PREFIX = "workbook.tabLayout.v1.";
const GLOBAL_WORKBOOK_LAYOUT_KEY = "__global__";

export interface WorkbookTabLayout {
  sessionIds: string[];
  activeSessionId: string | null;
}

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
    if (!normalized) continue;
    sessionIds.push(normalized);
  }

  const normalizedActiveSessionId = normalizeSessionId(layout.activeSessionId);
  const activeSessionId = normalizedActiveSessionId && sessionIds.includes(normalizedActiveSessionId)
    ? normalizedActiveSessionId
    : (sessionIds[0] ?? null);

  return {
    sessionIds,
    activeSessionId,
  };
}

export function parseWorkbookTabLayout(value: unknown): WorkbookTabLayout | null {
  if (!isRecord(value)) return null;

  const rawSessionIds = value.sessionIds;
  if (!isSessionIdList(rawSessionIds)) return null;

  const parsed: WorkbookTabLayout = {
    sessionIds: rawSessionIds
      .map((sessionId) => normalizeSessionId(sessionId))
      .filter((sessionId): sessionId is string => sessionId !== null),
    activeSessionId: normalizeSessionId(value.activeSessionId),
  };

  const normalized = normalizeWorkbookTabLayout(parsed);
  if (normalized.sessionIds.length === 0) return null;

  return normalized;
}

export async function loadWorkbookTabLayout(
  settings: SettingsStore,
  workbookId: string | null,
): Promise<WorkbookTabLayout | null> {
  try {
    const stored = await settings.get(workbookTabLayoutKey(workbookId));
    return parseWorkbookTabLayout(stored);
  } catch {
    return null;
  }
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
