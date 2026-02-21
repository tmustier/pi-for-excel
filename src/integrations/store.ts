/**
 * Integrations persistence helpers.
 *
 * Integrations can be enabled in two scopes:
 * - session: only for one chat tab/session
 * - workbook: applies to all sessions for the active workbook
 *
 * Workbook scope that has never been configured (null in storage) inherits
 * catalog defaults (e.g. web_search is enabled by default). Session scope
 * stays explicit by default, with an opt-in runtime fallback when workbook
 * identity is unavailable.
 */

import { getDefaultEnabledIntegrationIds } from "./catalog.js";

export type IntegrationScope = "session" | "workbook";

export interface IntegrationSettingsStore {
  get(key: string): Promise<unknown>;
  set(key: string, value: unknown): Promise<void>;
}

export interface SessionIntegrationIdsOptions {
  /**
   * When true and the session scope has never been configured, return
   * catalog defaults instead of an empty set.
   */
  applyDefaultsWhenUnconfigured?: boolean;
}

const SESSION_INTEGRATIONS_PREFIX = "integrations.session.v1.";
const WORKBOOK_INTEGRATIONS_PREFIX = "integrations.workbook.v1.";

export const EXTERNAL_TOOLS_ENABLED_SETTING_KEY = "external.tools.enabled";

export function sessionIntegrationsKey(sessionId: string): string {
  return `${SESSION_INTEGRATIONS_PREFIX}${sessionId}`;
}

export function workbookIntegrationsKey(workbookId: string): string {
  return `${WORKBOOK_INTEGRATIONS_PREFIX}${workbookId}`;
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

export function normalizeIntegrationIds(raw: unknown, knownIntegrationIds: readonly string[]): string[] {
  const known = new Set<string>(knownIntegrationIds);
  const requested = normalizeStringArray(raw);

  const enabledSet = new Set<string>();
  for (const integrationId of requested) {
    if (!known.has(integrationId)) continue;
    enabledSet.add(integrationId);
  }

  const ordered: string[] = [];
  for (const integrationId of knownIntegrationIds) {
    if (enabledSet.has(integrationId)) {
      ordered.push(integrationId);
    }
  }

  return ordered;
}

async function getScopeIntegrationIds(
  settings: IntegrationSettingsStore,
  scope: IntegrationScope,
  identifier: string,
  knownIntegrationIds: readonly string[],
  applyDefaultsWhenUnconfigured: boolean,
): Promise<string[]> {
  const key = scope === "session"
    ? sessionIntegrationsKey(identifier)
    : workbookIntegrationsKey(identifier);
  const raw = await settings.get(key);

  // Never configured workbook scope inherits defaults (web search).
  // Session scope stays explicit unless the caller opts into fallback defaults.
  if (raw == null) {
    if (scope === "workbook" || applyDefaultsWhenUnconfigured) {
      return normalizeIntegrationIds(getDefaultEnabledIntegrationIds(), knownIntegrationIds);
    }
    return [];
  }

  return normalizeIntegrationIds(raw, knownIntegrationIds);
}

async function setScopeIntegrationIds(
  settings: IntegrationSettingsStore,
  scope: IntegrationScope,
  identifier: string,
  integrationIds: readonly string[],
  knownIntegrationIds: readonly string[],
): Promise<void> {
  const key = scope === "session"
    ? sessionIntegrationsKey(identifier)
    : workbookIntegrationsKey(identifier);

  const normalized = normalizeIntegrationIds(integrationIds, knownIntegrationIds);
  await settings.set(key, normalized);
}

export async function getSessionIntegrationIds(
  settings: IntegrationSettingsStore,
  sessionId: string,
  knownIntegrationIds: readonly string[],
  options?: SessionIntegrationIdsOptions,
): Promise<string[]> {
  return getScopeIntegrationIds(
    settings,
    "session",
    sessionId,
    knownIntegrationIds,
    options?.applyDefaultsWhenUnconfigured === true,
  );
}

export async function setSessionIntegrationIds(
  settings: IntegrationSettingsStore,
  sessionId: string,
  integrationIds: readonly string[],
  knownIntegrationIds: readonly string[],
): Promise<void> {
  await setScopeIntegrationIds(settings, "session", sessionId, integrationIds, knownIntegrationIds);
}

export async function getWorkbookIntegrationIds(
  settings: IntegrationSettingsStore,
  workbookId: string,
  knownIntegrationIds: readonly string[],
): Promise<string[]> {
  return getScopeIntegrationIds(settings, "workbook", workbookId, knownIntegrationIds, false);
}

export async function setWorkbookIntegrationIds(
  settings: IntegrationSettingsStore,
  workbookId: string,
  integrationIds: readonly string[],
  knownIntegrationIds: readonly string[],
): Promise<void> {
  await setScopeIntegrationIds(settings, "workbook", workbookId, integrationIds, knownIntegrationIds);
}

export async function setIntegrationEnabledInScope(args: {
  settings: IntegrationSettingsStore;
  scope: IntegrationScope;
  identifier: string;
  integrationId: string;
  enabled: boolean;
  knownIntegrationIds: readonly string[];
}): Promise<void> {
  const { settings, scope, identifier, integrationId, enabled, knownIntegrationIds } = args;
  const existing = await getScopeIntegrationIds(settings, scope, identifier, knownIntegrationIds, false);

  const nextSet = new Set<string>(existing);
  if (enabled) {
    nextSet.add(integrationId);
  } else {
    nextSet.delete(integrationId);
  }

  const nextIds = Array.from(nextSet);
  await setScopeIntegrationIds(settings, scope, identifier, nextIds, knownIntegrationIds);
}

export async function resolveConfiguredIntegrationIds(args: {
  settings: IntegrationSettingsStore;
  sessionId: string;
  workbookId: string | null;
  knownIntegrationIds: readonly string[];
}): Promise<string[]> {
  const { settings, sessionId, workbookId, knownIntegrationIds } = args;

  const sessionIntegrationIds = await getSessionIntegrationIds(
    settings,
    sessionId,
    knownIntegrationIds,
    { applyDefaultsWhenUnconfigured: workbookId === null },
  );
  const workbookIntegrationIds = workbookId
    ? await getWorkbookIntegrationIds(settings, workbookId, knownIntegrationIds)
    : [];

  const enabledSet = new Set<string>([...workbookIntegrationIds, ...sessionIntegrationIds]);

  const ordered: string[] = [];
  for (const integrationId of knownIntegrationIds) {
    if (enabledSet.has(integrationId)) {
      ordered.push(integrationId);
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

export async function getExternalToolsEnabled(settings: IntegrationSettingsStore): Promise<boolean> {
  const raw = await settings.get(EXTERNAL_TOOLS_ENABLED_SETTING_KEY);
  // Default ON so web search is available once a provider API key is configured.
  if (raw == null) return true;
  return parseStoredBoolean(raw);
}

export async function setExternalToolsEnabled(
  settings: IntegrationSettingsStore,
  enabled: boolean,
): Promise<void> {
  await settings.set(EXTERNAL_TOOLS_ENABLED_SETTING_KEY, enabled ? "1" : "0");
}
