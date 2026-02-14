import assert from "node:assert/strict";
import { test } from "node:test";

import { stripYamlFrontmatter } from "../src/ui/markdown-preprocess.ts";

void test("stripYamlFrontmatter removes leading YAML metadata block", () => {
  const input = `---
name: xlsx
description: Use this skill for spreadsheet workflows
metadata:
  integration-id: sample
---

## Important Requirements
Body content.
`;

  const output = stripYamlFrontmatter(input);
  assert.equal(output.trimStart().startsWith("## Important Requirements"), true);
  assert.equal(output.includes("name: xlsx"), false);
});

void test("stripYamlFrontmatter preserves markdown that starts with thematic breaks", () => {
  const input = `---
Section intro
---

## Real heading
Body content.
`;

  const output = stripYamlFrontmatter(input);
  assert.equal(output, input);
});

void test("stripYamlFrontmatter supports YAML list values", () => {
  const input = `---
tags:
  - excel
  - formulas
---
# Skill
`;

  const output = stripYamlFrontmatter(input);
  assert.equal(output, "# Skill\n");
});

void test("stripYamlFrontmatter ignores non-YAML front block content", () => {
  const input = `---
## Intro: details
---
# Heading
`;

  const output = stripYamlFrontmatter(input);
  assert.equal(output, input);
});

void test("stripYamlFrontmatter does not treat prose labels as metadata keys", () => {
  const input = `---
Note: this section is prose, not metadata
---
# Heading
`;

  const output = stripYamlFrontmatter(input);
  assert.equal(output, input);
});

void test("stripYamlFrontmatter supports canonical title-cased frontmatter keys", () => {
  const input = `---
Title: Spreadsheet Skill
Date: 2026-02-14
---
# Heading
`;

  const output = stripYamlFrontmatter(input);
  assert.equal(output, "# Heading\n");
});
