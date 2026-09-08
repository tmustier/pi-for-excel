import assert from "node:assert/strict";
import { test } from "node:test";

import { parseSkillDocument } from "../src/skills/frontmatter.ts";

void test("parseSkillDocument reads required frontmatter fields", () => {
  const parsed = parseSkillDocument(`---
name: sample-skill
description: Helpful workflow for sample tasks.
compatibility: Requires sample runtime.
metadata:
  integration-id: sample
---

# Sample Skill

Use this skill when needed.
`);

  assert.ok(parsed);
  assert.equal(parsed.frontmatter.name, "sample-skill");
  assert.equal(parsed.frontmatter.description, "Helpful workflow for sample tasks.");
  assert.equal(parsed.frontmatter.compatibility, "Requires sample runtime.");
  assert.match(parsed.body, /# Sample Skill/);
});

void test("parseSkillDocument returns null when required fields are missing", () => {
  const parsed = parseSkillDocument(`---
name: missing-description
---

# Missing Description
`);

  assert.equal(parsed, null);
});
