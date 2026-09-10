import assert from "node:assert/strict";
import { test } from "node:test";

import type {
  WorkspaceFileEntry,
  WorkspaceFileReadResult,
} from "../src/files/types.ts";
import type {
  WorkspaceMutationOptions,
  WorkspaceReadOptions,
} from "../src/files/workspace.ts";
import type { AgentSkillDefinition } from "../src/skills/types.ts";
import {
  SKILL_ACTIVATION_STORAGE_KEY,
  filterAgentSkillsByEnabledState,
  loadDisabledSkillNamesFromSettings,
  setSkillEnabledInSettings,
} from "../src/skills/activation-store.ts";
import {
  loadDiscoverableAgentSkillsFromWorkspace,
  loadExternalAgentSkillsFromWorkspace,
  loadWorkspaceAgentSkillsFromWorkspace,
  removeExternalAgentSkillFromWorkspace,
  upsertExternalAgentSkillInWorkspace,
} from "../src/skills/external-store.ts";

interface MemoryWorkspaceFile {
  text: string;
  modifiedAt: number;
}

class MemoryExternalSkillWorkspace {
  private readonly files = new Map<string, MemoryWorkspaceFile>();

  seedTextFile(path: string, text: string): void {
    this.files.set(path, {
      text,
      modifiedAt: Date.now(),
    });
  }

  listPaths(): string[] {
    return Array.from(this.files.keys()).sort((left, right) => left.localeCompare(right));
  }

  listFiles(): Promise<WorkspaceFileEntry[]> {
    const entries = Array.from(this.files.entries())
      .sort((left, right) => left[0].localeCompare(right[0]))
      .map(([path, file]) => this.buildEntry(path, file));

    return Promise.resolve(entries);
  }

  readFile(path: string, options: WorkspaceReadOptions = {}): Promise<WorkspaceFileReadResult> {
    const file = this.files.get(path);
    if (!file) {
      throw new Error(`File not found: ${path}`);
    }

    const maxChars = options.maxChars;
    const limitedText = typeof maxChars === "number"
      ? file.text.slice(0, Math.max(0, maxChars))
      : file.text;
    const truncated = typeof maxChars === "number" && file.text.length > Math.max(0, maxChars);

    return Promise.resolve({
      ...this.buildEntry(path, file),
      text: limitedText,
      truncated,
    });
  }

  writeTextFile(
    path: string,
    text: string,
    _mimeTypeHint?: string,
    _options: WorkspaceMutationOptions = {},
  ): Promise<void> {
    this.files.set(path, {
      text,
      modifiedAt: Date.now(),
    });

    return Promise.resolve();
  }

  deleteFile(path: string, _options: WorkspaceMutationOptions = {}): Promise<void> {
    if (!this.files.delete(path)) {
      throw new Error(`File not found: ${path}`);
    }

    return Promise.resolve();
  }

  private buildEntry(path: string, file: MemoryWorkspaceFile): WorkspaceFileEntry {
    const name = path.split("/").at(-1) ?? path;

    return {
      path,
      name,
      size: file.text.length,
      modifiedAt: file.modifiedAt,
      mimeType: "text/markdown",
      kind: "text",
      sourceKind: "workspace",
      readOnly: false,
    };
  }
}

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

void test("loadExternalAgentSkillsFromWorkspace loads valid external skills", async () => {
  const workspace = new MemoryExternalSkillWorkspace();

  workspace.seedTextFile(
    "skills/external/custom-skill/SKILL.md",
    `---
name: custom-skill
description: External custom skill.
---

# Custom Skill
`,
  );

  const skills = await loadExternalAgentSkillsFromWorkspace(workspace);
  assert.equal(skills.length, 1);
  assert.equal(skills[0].name, "custom-skill");
  assert.equal(skills[0].sourceKind, "external");
  assert.equal(skills[0].location, "skills/external/custom-skill/SKILL.md");
});

void test("loadExternalAgentSkillsFromWorkspace ignores invalid or non-canonical files", async () => {
  const workspace = new MemoryExternalSkillWorkspace();

  workspace.seedTextFile("skills/external/invalid/SKILL.md", "# Missing frontmatter");
  workspace.seedTextFile(
    "skills/external/not-a-skill/README.md",
    `---
name: not-a-skill
description: Wrong filename.
---
`,
  );
  workspace.seedTextFile(
    "skills/external/nested/custom/SKILL.md",
    `---
name: nested
description: Wrong path depth.
---
`,
  );

  const skills = await loadExternalAgentSkillsFromWorkspace(workspace);
  assert.deepEqual(skills, []);
});

void test("loadWorkspaceAgentSkillsFromWorkspace discovers skills/<name>/SKILL.md", async () => {
  const workspace = new MemoryExternalSkillWorkspace();

  workspace.seedTextFile(
    "skills/workspace-direct/SKILL.md",
    `---
name: workspace-direct
description: Workspace-discovered skill.
---

# Workspace Skill
`,
  );
  workspace.seedTextFile(
    "skills/workspace-direct/README.md",
    `---
name: readme
description: Wrong filename.
---
`,
  );
  workspace.seedTextFile(
    "skills/external/managed/SKILL.md",
    `---
name: managed
description: Managed external skill.
---
`,
  );

  const skills = await loadWorkspaceAgentSkillsFromWorkspace(workspace);
  assert.equal(skills.length, 1);
  assert.equal(skills[0].name, "workspace-direct");
  assert.equal(skills[0].location, "skills/workspace-direct/SKILL.md");
});

void test("loadDiscoverableAgentSkillsFromWorkspace prefers managed external over workspace-discovered", async () => {
  const workspace = new MemoryExternalSkillWorkspace();

  workspace.seedTextFile(
    "skills/external/shared-name/SKILL.md",
    `---
name: shared-skill
description: Managed external copy.
---

# Managed
`,
  );
  workspace.seedTextFile(
    "skills/workspace-copy/SKILL.md",
    `---
name: shared-skill
description: Workspace copy.
---

# Workspace
`,
  );
  workspace.seedTextFile(
    "skills/workspace-only/SKILL.md",
    `---
name: workspace-only
description: Workspace-only skill.
---

# Workspace only
`,
  );

  const skills = await loadDiscoverableAgentSkillsFromWorkspace(workspace);

  assert.deepEqual(skills.map((skill) => skill.name), ["shared-skill", "workspace-only"]);
  const shared = skills.find((skill) => skill.name === "shared-skill");
  assert.ok(shared);
  assert.equal(shared?.location, "skills/external/shared-name/SKILL.md");
});

void test("upsertExternalAgentSkillInWorkspace installs and overwrites by skill name", async () => {
  const workspace = new MemoryExternalSkillWorkspace();

  workspace.seedTextFile(
    "skills/external/legacy-copy/SKILL.md",
    `---
name: custom-skill
description: Legacy duplicate.
---

# Legacy
`,
  );

  await upsertExternalAgentSkillInWorkspace({
    workspace,
    markdown: `---
name: custom-skill
description: Updated description.
---

# Updated
`,
  });

  const skills = await loadExternalAgentSkillsFromWorkspace(workspace);
  assert.equal(skills.length, 1);
  assert.equal(skills[0].name, "custom-skill");
  assert.equal(skills[0].description, "Updated description.");
  assert.equal(skills[0].location, "skills/external/custom-skill/SKILL.md");
  assert.deepEqual(workspace.listPaths(), ["skills/external/custom-skill/SKILL.md"]);
});

void test("upsertExternalAgentSkillInWorkspace enforces expectedName when provided", async () => {
  const workspace = new MemoryExternalSkillWorkspace();

  await assert.rejects(
    () => upsertExternalAgentSkillInWorkspace({
      workspace,
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

void test("removeExternalAgentSkillFromWorkspace removes by name and reports whether removed", async () => {
  const workspace = new MemoryExternalSkillWorkspace();

  workspace.seedTextFile(
    "skills/external/custom-skill/SKILL.md",
    `---
name: custom-skill
description: Canonical copy.
---

# Canonical
`,
  );
  workspace.seedTextFile(
    "skills/external/legacy-copy/SKILL.md",
    `---
name: custom-skill
description: Duplicate copy.
---

# Duplicate
`,
  );

  const removed = await removeExternalAgentSkillFromWorkspace({
    workspace,
    name: "custom-skill",
  });
  assert.equal(removed, true);

  const skillsAfterRemove = await loadExternalAgentSkillsFromWorkspace(workspace);
  assert.deepEqual(skillsAfterRemove, []);
  assert.deepEqual(workspace.listPaths(), []);

  const removedMissing = await removeExternalAgentSkillFromWorkspace({
    workspace,
    name: "missing-skill",
  });
  assert.equal(removedMissing, false);
});

void test("loadDisabledSkillNamesFromSettings normalizes and deduplicates names", async () => {
  const settings = new MemorySettingsStore();

  await settings.set(SKILL_ACTIVATION_STORAGE_KEY, {
    version: 1,
    disabledNames: [" Web-Search ", "custom-skill", "web-search", ""],
  });

  const disabled = await loadDisabledSkillNamesFromSettings(settings);
  assert.deepEqual(Array.from(disabled).sort(), ["custom-skill", "web-search"]);
});

void test("setSkillEnabledInSettings disables and re-enables by name", async () => {
  const settings = new MemorySettingsStore();

  const disabled = await setSkillEnabledInSettings({
    settings,
    name: "Web-Search",
    enabled: false,
  });

  assert.equal(disabled.changed, true);
  assert.equal(disabled.enabled, false);
  assert.equal(disabled.name, "web-search");

  const disabledNames = await loadDisabledSkillNamesFromSettings(settings);
  assert.deepEqual(Array.from(disabledNames), ["web-search"]);

  const duplicateDisable = await setSkillEnabledInSettings({
    settings,
    name: "web-search",
    enabled: false,
  });
  assert.equal(duplicateDisable.changed, false);

  const enabled = await setSkillEnabledInSettings({
    settings,
    name: "web-search",
    enabled: true,
  });

  assert.equal(enabled.changed, true);
  assert.equal(enabled.enabled, true);

  const afterEnable = await loadDisabledSkillNamesFromSettings(settings);
  assert.deepEqual(Array.from(afterEnable), []);
});

void test("filterAgentSkillsByEnabledState excludes disabled skill names", () => {
  const bundledSkill: AgentSkillDefinition = {
    name: "web-search",
    description: "Bundled web search.",
    location: "skills/web-search/SKILL.md",
    sourceKind: "bundled",
    markdown: "# Web Search",
    body: "# Web Search",
  };

  const externalSkill: AgentSkillDefinition = {
    name: "custom-skill",
    description: "External skill.",
    location: "skills/external/custom-skill/SKILL.md",
    sourceKind: "external",
    markdown: "# Custom Skill",
    body: "# Custom Skill",
  };

  const filtered = filterAgentSkillsByEnabledState({
    skills: [bundledSkill, externalSkill],
    disabledSkillNames: new Set(["custom-skill"]),
  });

  assert.deepEqual(filtered.map((skill) => skill.name), ["web-search"]);
});
