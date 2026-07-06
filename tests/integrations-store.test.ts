import assert from "node:assert/strict";
import { test } from "node:test";

import {
  getExternalToolsEnabled,
  getSessionIntegrationIds,
  getWorkbookIntegrationIds,
  resolveConfiguredIntegrationIds,
  setExternalToolsEnabled,
  setSessionIntegrationIds,
  setIntegrationEnabledInScope,
  setWorkbookIntegrationIds,
} from "../src/integrations/store.ts";

const KNOWN_INTEGRATIONS = ["web_search", "mcp_tools"] as const;

class MemorySettingsStore {
  private readonly values = new Map<string, DynamicValue>();

  get(key: string): Promise<DynamicValue> {
    return Promise.resolve(this.values.has(key) ? this.values.get(key) ?? null : null);
  }

  set(key: string, value: DynamicValue): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }
}

void test("resolves session + workbook integrations in catalog order", async () => {
  const settings = new MemorySettingsStore();

  await setSessionIntegrationIds(settings, "session-1", ["mcp_tools"], KNOWN_INTEGRATIONS);
  await setWorkbookIntegrationIds(settings, "workbook-1", ["web_search"], KNOWN_INTEGRATIONS);

  const resolved = await resolveConfiguredIntegrationIds({
    settings,
    sessionId: "session-1",
    workbookId: "workbook-1",
    knownIntegrationIds: KNOWN_INTEGRATIONS,
  });

  assert.deepEqual(resolved, ["web_search", "mcp_tools"]);
});

void test("setIntegrationEnabledInScope toggles session/workbook flags", async () => {
  const settings = new MemorySettingsStore();

  await setIntegrationEnabledInScope({
    settings,
    scope: "session",
    identifier: "session-2",
    integrationId: "web_search",
    enabled: true,
    knownIntegrationIds: KNOWN_INTEGRATIONS,
  });

  await setIntegrationEnabledInScope({
    settings,
    scope: "workbook",
    identifier: "workbook-2",
    integrationId: "mcp_tools",
    enabled: true,
    knownIntegrationIds: KNOWN_INTEGRATIONS,
  });

  assert.deepEqual(
    await getSessionIntegrationIds(settings, "session-2", KNOWN_INTEGRATIONS),
    ["web_search"],
  );
  assert.deepEqual(
    await getWorkbookIntegrationIds(settings, "workbook-2", KNOWN_INTEGRATIONS),
    ["web_search", "mcp_tools"],
  );

  await setIntegrationEnabledInScope({
    settings,
    scope: "session",
    identifier: "session-2",
    integrationId: "web_search",
    enabled: false,
    knownIntegrationIds: KNOWN_INTEGRATIONS,
  });

  assert.deepEqual(
    await getSessionIntegrationIds(settings, "session-2", KNOWN_INTEGRATIONS),
    [],
  );
});

void test("external tools gate defaults on and can be disabled", async () => {
  const settings = new MemorySettingsStore();

  assert.equal(await getExternalToolsEnabled(settings), true);

  await setExternalToolsEnabled(settings, false);
  assert.equal(await getExternalToolsEnabled(settings), false);

  await setExternalToolsEnabled(settings, true);
  assert.equal(await getExternalToolsEnabled(settings), true);
});

void test("unconfigured session scope is explicit empty", async () => {
  const settings = new MemorySettingsStore();

  const ids = await getSessionIntegrationIds(settings, "new-session", KNOWN_INTEGRATIONS);
  assert.deepEqual(ids, []);
});

void test("session scope can opt into defaults when unconfigured", async () => {
  const settings = new MemorySettingsStore();

  const ids = await getSessionIntegrationIds(settings, "new-session", KNOWN_INTEGRATIONS, {
    applyDefaultsWhenUnconfigured: true,
  });
  assert.deepEqual(ids, ["web_search"]);
});

void test("explicitly cleared session scope returns empty", async () => {
  const settings = new MemorySettingsStore();

  // Set then clear web_search → stores [] explicitly
  await setIntegrationEnabledInScope({
    settings,
    scope: "session",
    identifier: "session-x",
    integrationId: "web_search",
    enabled: true,
    knownIntegrationIds: KNOWN_INTEGRATIONS,
  });
  await setIntegrationEnabledInScope({
    settings,
    scope: "session",
    identifier: "session-x",
    integrationId: "web_search",
    enabled: false,
    knownIntegrationIds: KNOWN_INTEGRATIONS,
  });

  const ids = await getSessionIntegrationIds(settings, "session-x", KNOWN_INTEGRATIONS);
  assert.deepEqual(ids, []);

  const fallbackIds = await getSessionIntegrationIds(settings, "session-x", KNOWN_INTEGRATIONS, {
    applyDefaultsWhenUnconfigured: true,
  });
  assert.deepEqual(fallbackIds, []);
});

void test("unconfigured workbook scope returns default-enabled integrations", async () => {
  const settings = new MemorySettingsStore();

  const ids = await getWorkbookIntegrationIds(settings, "new-workbook", KNOWN_INTEGRATIONS);
  assert.deepEqual(ids, ["web_search"]);
});

void test("resolveConfiguredIntegrationIds includes defaults for fresh session+workbook", async () => {
  const settings = new MemorySettingsStore();

  const ids = await resolveConfiguredIntegrationIds({
    settings,
    sessionId: "fresh-session",
    workbookId: "fresh-workbook",
    knownIntegrationIds: KNOWN_INTEGRATIONS,
  });

  assert.deepEqual(ids, ["web_search"]);
});

void test("resolveConfiguredIntegrationIds includes defaults when workbook identity is unavailable", async () => {
  const settings = new MemorySettingsStore();

  const ids = await resolveConfiguredIntegrationIds({
    settings,
    sessionId: "fresh-session-no-workbook",
    workbookId: null,
    knownIntegrationIds: KNOWN_INTEGRATIONS,
  });

  assert.deepEqual(ids, ["web_search"]);
});

void test("workbook-level disable persists across new sessions", async () => {
  const settings = new MemorySettingsStore();

  await setIntegrationEnabledInScope({
    settings,
    scope: "workbook",
    identifier: "workbook-off",
    integrationId: "web_search",
    enabled: false,
    knownIntegrationIds: KNOWN_INTEGRATIONS,
  });

  assert.deepEqual(
    await getWorkbookIntegrationIds(settings, "workbook-off", KNOWN_INTEGRATIONS),
    [],
  );

  const resolved = await resolveConfiguredIntegrationIds({
    settings,
    sessionId: "brand-new-session",
    workbookId: "workbook-off",
    knownIntegrationIds: KNOWN_INTEGRATIONS,
  });

  assert.deepEqual(resolved, []);
});
