import assert from "node:assert/strict";
import { test } from "node:test";

import {
  getExternalToolsEnabled,
  getSessionSkillIds,
  getWorkbookSkillIds,
  resolveConfiguredSkillIds,
  setExternalToolsEnabled,
  setSessionSkillIds,
  setSkillEnabledInScope,
  setWorkbookSkillIds,
} from "../src/skills/store.ts";

const KNOWN_SKILLS = ["web_search", "mcp_tools"] as const;

class MemorySettingsStore {
  private readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.has(key) ? this.values.get(key) ?? null : null);
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }
}

void test("resolves session + workbook skills in catalog order", async () => {
  const settings = new MemorySettingsStore();

  await setSessionSkillIds(settings, "session-1", ["mcp_tools"], KNOWN_SKILLS);
  await setWorkbookSkillIds(settings, "workbook-1", ["web_search"], KNOWN_SKILLS);

  const resolved = await resolveConfiguredSkillIds({
    settings,
    sessionId: "session-1",
    workbookId: "workbook-1",
    knownSkillIds: KNOWN_SKILLS,
  });

  assert.deepEqual(resolved, ["web_search", "mcp_tools"]);
});

void test("setSkillEnabledInScope toggles session/workbook flags", async () => {
  const settings = new MemorySettingsStore();

  await setSkillEnabledInScope({
    settings,
    scope: "session",
    identifier: "session-2",
    skillId: "web_search",
    enabled: true,
    knownSkillIds: KNOWN_SKILLS,
  });

  await setSkillEnabledInScope({
    settings,
    scope: "workbook",
    identifier: "workbook-2",
    skillId: "mcp_tools",
    enabled: true,
    knownSkillIds: KNOWN_SKILLS,
  });

  assert.deepEqual(
    await getSessionSkillIds(settings, "session-2", KNOWN_SKILLS),
    ["web_search"],
  );
  assert.deepEqual(
    await getWorkbookSkillIds(settings, "workbook-2", KNOWN_SKILLS),
    ["mcp_tools"],
  );

  await setSkillEnabledInScope({
    settings,
    scope: "session",
    identifier: "session-2",
    skillId: "web_search",
    enabled: false,
    knownSkillIds: KNOWN_SKILLS,
  });

  assert.deepEqual(
    await getSessionSkillIds(settings, "session-2", KNOWN_SKILLS),
    [],
  );
});

void test("external tools gate defaults off and can be enabled", async () => {
  const settings = new MemorySettingsStore();

  assert.equal(await getExternalToolsEnabled(settings), false);

  await setExternalToolsEnabled(settings, true);
  assert.equal(await getExternalToolsEnabled(settings), true);

  await setExternalToolsEnabled(settings, false);
  assert.equal(await getExternalToolsEnabled(settings), false);
});
