import assert from "node:assert/strict";
import { test } from "node:test";

import {
  EXTERNAL_AGENT_SKILLS_STORAGE_KEY,
  loadExternalAgentSkillsFromSettings,
  removeExternalAgentSkillFromSettings,
  upsertExternalAgentSkillInSettings,
} from "../src/skills/external-store.ts";

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

void test("loadExternalAgentSkillsFromSettings loads valid external skills", async () => {
  const settings = new MemorySettingsStore();

  await settings.set(EXTERNAL_AGENT_SKILLS_STORAGE_KEY, {
    version: 1,
    items: [
      {
        location: "/tmp/skills/custom-skill/SKILL.md",
        markdown: `---
name: custom-skill
description: External custom skill.
---

# Custom Skill
`,
      },
    ],
  });

  const skills = await loadExternalAgentSkillsFromSettings(settings);
  assert.equal(skills.length, 1);
  assert.equal(skills[0].name, "custom-skill");
  assert.equal(skills[0].sourceKind, "external");
  assert.equal(skills[0].location, "/tmp/skills/custom-skill/SKILL.md");
});

void test("loadExternalAgentSkillsFromSettings ignores invalid payloads", async () => {
  const settings = new MemorySettingsStore();

  await settings.set(EXTERNAL_AGENT_SKILLS_STORAGE_KEY, {
    version: 1,
    items: [
      {
        location: "/tmp/skills/invalid/SKILL.md",
        markdown: "# Missing frontmatter",
      },
      {
        location: 12,
        markdown: "---\nname: invalid\ndescription: invalid\n---",
      },
    ],
  });

  const skills = await loadExternalAgentSkillsFromSettings(settings);
  assert.deepEqual(skills, []);
});

void test("loadExternalAgentSkillsFromSettings returns empty for unknown version", async () => {
  const settings = new MemorySettingsStore();

  void settings.set(EXTERNAL_AGENT_SKILLS_STORAGE_KEY, {
    version: 2,
    items: [],
  });

  const skills = await loadExternalAgentSkillsFromSettings(settings);
  assert.deepEqual(skills, []);
});

void test("upsertExternalAgentSkillInSettings installs and overwrites by skill name", async () => {
  const settings = new MemorySettingsStore();

  await upsertExternalAgentSkillInSettings({
    settings,
    markdown: `---
name: custom-skill
description: First description.
---

# First
`,
  });

  await upsertExternalAgentSkillInSettings({
    settings,
    markdown: `---
name: custom-skill
description: Updated description.
---

# Updated
`,
  });

  const skills = await loadExternalAgentSkillsFromSettings(settings);
  assert.equal(skills.length, 1);
  assert.equal(skills[0].name, "custom-skill");
  assert.equal(skills[0].description, "Updated description.");
  assert.equal(skills[0].location, "skills/external/custom-skill/SKILL.md");
});

void test("upsertExternalAgentSkillInSettings enforces expectedName when provided", async () => {
  const settings = new MemorySettingsStore();

  await assert.rejects(
    () => upsertExternalAgentSkillInSettings({
      settings,
      expectedName: "different-name",
      markdown: `---
name: custom-skill
description: External custom skill.
---

# Custom
`,
    }),
    /Skill name mismatch: expected "different-name" but markdown declares "custom-skill"/,
  );
});

void test("removeExternalAgentSkillFromSettings removes by name and reports whether removed", async () => {
  const settings = new MemorySettingsStore();

  await upsertExternalAgentSkillInSettings({
    settings,
    markdown: `---
name: custom-skill
description: External custom skill.
---

# Custom
`,
  });

  const removed = await removeExternalAgentSkillFromSettings({
    settings,
    name: "custom-skill",
  });
  assert.equal(removed, true);

  const skillsAfterRemove = await loadExternalAgentSkillsFromSettings(settings);
  assert.deepEqual(skillsAfterRemove, []);

  const removedMissing = await removeExternalAgentSkillFromSettings({
    settings,
    name: "missing-skill",
  });
  assert.equal(removedMissing, false);
});
