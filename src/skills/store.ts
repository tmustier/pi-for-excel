/**
 * Skills persistence helpers.
 *
 * Skills can be enabled in two scopes:
 * - session: only for one chat tab/session
 * - workbook: applies to all sessions for the active workbook
 */

export type SkillScope = "session" | "workbook";

export interface SkillSettingsStore {
  get(key: string): Promise<unknown>;
  set(key: string, value: unknown): Promise<void>;
}

const SESSION_SKILLS_PREFIX = "skills.session.v1.";
const WORKBOOK_SKILLS_PREFIX = "skills.workbook.v1.";

export const EXTERNAL_TOOLS_ENABLED_SETTING_KEY = "external.tools.enabled";

export function sessionSkillsKey(sessionId: string): string {
  return `${SESSION_SKILLS_PREFIX}${sessionId}`;
}

export function workbookSkillsKey(workbookId: string): string {
  return `${WORKBOOK_SKILLS_PREFIX}${workbookId}`;
}

function normalizeStringArray(value: unknown): string[] {
  if (Array.isArray(value)) {
    const out: string[] = [];
    for (const item of value) {
      if (typeof item !== "string") continue;
      const trimmed = item.trim();
      if (trimmed.length === 0) continue;
      out.push(trimmed);
    }
    return out;
  }

  if (typeof value === "string") {
    return value
      .split(",")
      .map((token) => token.trim())
      .filter((token) => token.length > 0);
  }

  return [];
}

export function normalizeSkillIds(raw: unknown, knownSkillIds: readonly string[]): string[] {
  const known = new Set<string>(knownSkillIds);
  const requested = normalizeStringArray(raw);

  const enabledSet = new Set<string>();
  for (const skillId of requested) {
    if (!known.has(skillId)) continue;
    enabledSet.add(skillId);
  }

  const ordered: string[] = [];
  for (const skillId of knownSkillIds) {
    if (enabledSet.has(skillId)) {
      ordered.push(skillId);
    }
  }

  return ordered;
}

async function getScopeSkillIds(
  settings: SkillSettingsStore,
  scope: SkillScope,
  identifier: string,
  knownSkillIds: readonly string[],
): Promise<string[]> {
  const key = scope === "session"
    ? sessionSkillsKey(identifier)
    : workbookSkillsKey(identifier);
  const raw = await settings.get(key);
  return normalizeSkillIds(raw, knownSkillIds);
}

async function setScopeSkillIds(
  settings: SkillSettingsStore,
  scope: SkillScope,
  identifier: string,
  skillIds: readonly string[],
  knownSkillIds: readonly string[],
): Promise<void> {
  const key = scope === "session"
    ? sessionSkillsKey(identifier)
    : workbookSkillsKey(identifier);

  const normalized = normalizeSkillIds(skillIds, knownSkillIds);
  await settings.set(key, normalized);
}

export async function getSessionSkillIds(
  settings: SkillSettingsStore,
  sessionId: string,
  knownSkillIds: readonly string[],
): Promise<string[]> {
  return getScopeSkillIds(settings, "session", sessionId, knownSkillIds);
}

export async function setSessionSkillIds(
  settings: SkillSettingsStore,
  sessionId: string,
  skillIds: readonly string[],
  knownSkillIds: readonly string[],
): Promise<void> {
  await setScopeSkillIds(settings, "session", sessionId, skillIds, knownSkillIds);
}

export async function getWorkbookSkillIds(
  settings: SkillSettingsStore,
  workbookId: string,
  knownSkillIds: readonly string[],
): Promise<string[]> {
  return getScopeSkillIds(settings, "workbook", workbookId, knownSkillIds);
}

export async function setWorkbookSkillIds(
  settings: SkillSettingsStore,
  workbookId: string,
  skillIds: readonly string[],
  knownSkillIds: readonly string[],
): Promise<void> {
  await setScopeSkillIds(settings, "workbook", workbookId, skillIds, knownSkillIds);
}

export async function setSkillEnabledInScope(args: {
  settings: SkillSettingsStore;
  scope: SkillScope;
  identifier: string;
  skillId: string;
  enabled: boolean;
  knownSkillIds: readonly string[];
}): Promise<void> {
  const { settings, scope, identifier, skillId, enabled, knownSkillIds } = args;
  const existing = await getScopeSkillIds(settings, scope, identifier, knownSkillIds);

  const nextSet = new Set<string>(existing);
  if (enabled) {
    nextSet.add(skillId);
  } else {
    nextSet.delete(skillId);
  }

  const nextIds = Array.from(nextSet);
  await setScopeSkillIds(settings, scope, identifier, nextIds, knownSkillIds);
}

export async function resolveConfiguredSkillIds(args: {
  settings: SkillSettingsStore;
  sessionId: string;
  workbookId: string | null;
  knownSkillIds: readonly string[];
}): Promise<string[]> {
  const { settings, sessionId, workbookId, knownSkillIds } = args;

  const sessionSkillIds = await getSessionSkillIds(settings, sessionId, knownSkillIds);
  const workbookSkillIds = workbookId
    ? await getWorkbookSkillIds(settings, workbookId, knownSkillIds)
    : [];

  const enabledSet = new Set<string>([...workbookSkillIds, ...sessionSkillIds]);

  const ordered: string[] = [];
  for (const skillId of knownSkillIds) {
    if (enabledSet.has(skillId)) {
      ordered.push(skillId);
    }
  }

  return ordered;
}

function parseStoredBoolean(value: unknown): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    return normalized === "1" || normalized === "true" || normalized === "yes";
  }
  return false;
}

export async function getExternalToolsEnabled(settings: SkillSettingsStore): Promise<boolean> {
  const raw = await settings.get(EXTERNAL_TOOLS_ENABLED_SETTING_KEY);
  return parseStoredBoolean(raw);
}

export async function setExternalToolsEnabled(
  settings: SkillSettingsStore,
  enabled: boolean,
): Promise<void> {
  await settings.set(EXTERNAL_TOOLS_ENABLED_SETTING_KEY, enabled ? "1" : "0");
}
