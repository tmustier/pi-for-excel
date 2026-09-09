/**
 * Session ↔ workbook association.
 *
 * `SessionsStore` metadata schema is fixed (pi-web-ui), so we store the mapping
 * in `SettingsStore` using a dedicated key prefix.
 */

export interface SessionAssociationSettingsStore {
  get(key: string): Promise<DynamicValue>;
  set(key: string, value: DynamicValue): Promise<void>;
}

const SESSION_WORKBOOK_PREFIX = "session.workbook.v1.";
const WORKBOOK_LATEST_SESSION_PREFIX = "workbook.latestSession.v1.";

export function sessionWorkbookKey(sessionId: string): string {
  return `${SESSION_WORKBOOK_PREFIX}${sessionId}`;
}

export function workbookLatestSessionKey(workbookId: string): string {
  return `${WORKBOOK_LATEST_SESSION_PREFIX}${workbookId}`;
}

export async function getSessionWorkbookId(
  settings: SessionAssociationSettingsStore,
  sessionId: string,
): Promise<string | null> {
  const value = await settings.get(sessionWorkbookKey(sessionId));
  return typeof value === "string" && value.trim().length > 0 ? value : null;
}

/**
 * Link a session to a workbook.
 *
 * Does not overwrite an existing link (so resuming an old session in a different
 * workbook won't accidentally "move" it).
 */
export async function linkSessionToWorkbook(
  settings: SessionAssociationSettingsStore,
  sessionId: string,
  workbookId: string,
): Promise<void> {
  const key = sessionWorkbookKey(sessionId);
  const existing = await settings.get(key);
  if (typeof existing === "string" && existing.trim().length > 0) return;
  await settings.set(key, workbookId);
}

export async function setLatestSessionForWorkbook(
  settings: SessionAssociationSettingsStore,
  workbookId: string,
  sessionId: string,
): Promise<void> {
  await settings.set(workbookLatestSessionKey(workbookId), sessionId);
}

export async function getLatestSessionForWorkbook(
  settings: SessionAssociationSettingsStore,
  workbookId: string,
): Promise<string | null> {
  try {
    const value = await settings.get(workbookLatestSessionKey(workbookId));
    return typeof value === "string" && value.trim().length > 0 ? value : null;
  } catch {
    return null;
  }
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
  settings: SessionAssociationSettingsStore,
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
