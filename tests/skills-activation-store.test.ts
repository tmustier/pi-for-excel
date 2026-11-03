import assert from "node:assert/strict";
import { test } from "node:test";
import { loadDisabledSkillNamesFromSettings, setSkillEnabledInSettings } from "../src/skills/activation-store.ts";

void test("concurrent skill toggles preserve both disabled skills", async () => {
  const values = new Map<string, unknown>();
  const settings = {
    get: (key: string): Promise<unknown> => Promise.resolve(values.get(key)),
    set: (key: string, value: unknown): Promise<void> => { values.set(key, value); return Promise.resolve(); },
  };
  await Promise.all(["alpha", "beta"].map((name) => setSkillEnabledInSettings({ settings, name, enabled: false })));
  assert.deepEqual([...await loadDisabledSkillNamesFromSettings(settings)], ["alpha", "beta"]);
});
