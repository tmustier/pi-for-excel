import type { SettingsWriter } from "../storage/local/settings-store.js";

const SESSION_WORKBOOK_PREFIX = "session.workbook.v1.";
const WORKBOOK_LATEST_SESSION_PREFIX = "workbook.latestSession.v1.";

export function sessionWorkbookKey(sessionId: string): string {
  return `${SESSION_WORKBOOK_PREFIX}${sessionId}`;
}

export function workbookLatestSessionKey(workbookId: string): string {
  return `${WORKBOOK_LATEST_SESSION_PREFIX}${workbookId}`;
}

export async function getSessionWorkbookId(
  settings: SettingsWriter,
  sessionId: string,
): Promise<string | null> {
  const value = await settings.get(sessionWorkbookKey(sessionId));
  return typeof value === "string" && value.trim().length > 0 ? value : null;
}

export async function linkSessionToWorkbook(
  settings: SettingsWriter,
  sessionId: string,
  workbookId: string,
): Promise<void> {
  const key = sessionWorkbookKey(sessionId);
  const existing = await settings.get(key);
  if (typeof existing === "string" && existing.trim().length > 0) return;
  await settings.set(key, workbookId);
}

export async function moveSessionToWorkbook(
  settings: SettingsWriter,
  sessionId: string,
  from: string | null,
  to: string,
): Promise<void> {
  const existing = await getSessionWorkbookId(settings, sessionId);
  if (existing !== null && existing !== from) return;
  await settings.set(sessionWorkbookKey(sessionId), to);
}

export async function setLatestSessionForWorkbook(
  settings: SettingsWriter,
  workbookId: string,
  sessionId: string,
): Promise<void> {
  await settings.set(workbookLatestSessionKey(workbookId), sessionId);
}

export async function getLatestSessionForWorkbook(
  settings: SettingsWriter,
  workbookId: string,
): Promise<string | null> {
  const value = await settings.get(workbookLatestSessionKey(workbookId));
  return typeof value === "string" && value.trim().length > 0 ? value : null;
}

export interface SessionWorkbookPartition {
  matchingSessionIds: string[];
  unlinkedSessionIds: string[];
  foreignSessionIds: string[];
}

export async function partitionSessionIdsByWorkbook(
  settings: SettingsWriter,
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
