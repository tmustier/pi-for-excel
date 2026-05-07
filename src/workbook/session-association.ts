/**
 * Session ↔ workbook association.
 *
 * `SessionsStore` metadata schema is fixed (pi-web-ui), so we store the mapping
 * in `SettingsStore` using a dedicated key prefix.
 */

import type { SettingsStore } from "@earendil-works/pi-web-ui";

const SESSION_WORKBOOK_PREFIX = "session.workbook.v1.";
const WORKBOOK_LATEST_SESSION_PREFIX = "workbook.latestSession.v1.";

export function sessionWorkbookKey(sessionId: string): string {
  return `${SESSION_WORKBOOK_PREFIX}${sessionId}`;
}

export function workbookLatestSessionKey(workbookId: string): string {
  return `${WORKBOOK_LATEST_SESSION_PREFIX}${workbookId}`;
}

export async function getSessionWorkbookId(
  settings: SettingsStore,
  sessionId: string,
): Promise<string | null> {
  const v = await settings.get(sessionWorkbookKey(sessionId));
  return typeof v === "string" && v.trim().length > 0 ? v : null;
}

/**
 * Link a session to a workbook.
 *
 * Does not overwrite an existing link (so resuming an old session in a different
 * workbook won't accidentally "move" it).
 */
export async function linkSessionToWorkbook(
  settings: SettingsStore,
  sessionId: string,
  workbookId: string,
): Promise<void> {
  const key = sessionWorkbookKey(sessionId);
  const existing = await settings.get(key);
  if (typeof existing === "string" && existing.trim().length > 0) return;
  await settings.set(key, workbookId);
}

export async function setLatestSessionForWorkbook(
  settings: SettingsStore,
  workbookId: string,
  sessionId: string,
): Promise<void> {
  await settings.set(workbookLatestSessionKey(workbookId), sessionId);
}

export async function getLatestSessionForWorkbook(
  settings: SettingsStore,
  workbookId: string,
): Promise<string | null> {
  const v = await settings.get(workbookLatestSessionKey(workbookId));
  return typeof v === "string" && v.trim().length > 0 ? v : null;
}

export interface SessionWorkbookPartition {
  matchingSessionIds: string[];
  unlinkedSessionIds: string[];
  foreignSessionIds: string[];
}

/**
 * Partition session ids by workbook association.
 *
 * - `matchingSessionIds`: linked to this workbook
 * - `unlinkedSessionIds`: legacy/no mapping
 * - `foreignSessionIds`: linked to another workbook
 */
export async function partitionSessionIdsByWorkbook(
  settings: SettingsStore,
  sessionIds: string[],
  workbookId: string,
): Promise<SessionWorkbookPartition> {
  const matchingSessionIds: string[] = [];
  const unlinkedSessionIds: string[] = [];
  const foreignSessionIds: string[] = [];

  for (const sessionId of sessionIds) {
    const linkedWorkbookId = await getSessionWorkbookId(settings, sessionId);
    if (!linkedWorkbookId) {
      unlinkedSessionIds.push(sessionId);
      continue;
    }

    if (linkedWorkbookId === workbookId) {
      matchingSessionIds.push(sessionId);
      continue;
    }

    foreignSessionIds.push(sessionId);
  }

  return {
    matchingSessionIds,
    unlinkedSessionIds,
    foreignSessionIds,
  };
}
