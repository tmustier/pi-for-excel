import assert from "node:assert/strict";
import { test } from "node:test";
import { readFile } from "node:fs/promises";

import {
  buildCoreToolPromptLines,
  CORE_TOOL_CAPABILITIES,
  TOOL_DISCLOSURE_BUNDLES,
  TOOL_DISCLOSURE_FULL_ACCESS_PATTERNS,
  TOOL_DISCLOSURE_TRIGGER_PATTERNS,
  TOOL_NAMES_WITH_HUMANIZER,
  TOOL_NAMES_WITH_RENDERER,
  TOOL_UI_METADATA,
  UI_TOOL_NAMES,
} from "../src/tools/capabilities.ts";
import { CORE_TOOL_NAMES } from "../src/tools/names.ts";
import { buildSystemPrompt } from "../src/prompt/system-prompt.ts";

void test("core capability metadata covers all core tools", () => {
  assert.equal(CORE_TOOL_CAPABILITIES.length, CORE_TOOL_NAMES.length);

  const capabilityNames = CORE_TOOL_CAPABILITIES.map((capability) => capability.name);
  assert.deepEqual(capabilityNames, [...CORE_TOOL_NAMES]);

  for (const capability of CORE_TOOL_CAPABILITIES) {
    assert.equal(capability.tier, "core");
    assert.ok(capability.promptDescription.length > 0);
  }
});

void test("system prompt core tool section is generated from capability metadata", () => {
  const prompt = buildSystemPrompt();
  const toolLines = buildCoreToolPromptLines();

  for (const line of toolLines.split("\n")) {
    assert.equal(prompt.includes(line), true);
  }
});

void test("UI tool registration derives from centralized UI metadata", async () => {
  const rendererSource = await readFile(new URL("../src/ui/tool-renderers.ts", import.meta.url), "utf8");
  assert.match(rendererSource, /import\s*\{\s*TOOL_NAMES_WITH_RENDERER/);
  assert.match(rendererSource, /CUSTOM_RENDERED_TOOL_NAMES:\s*readonly SupportedToolName\[\]\s*=\s*TOOL_NAMES_WITH_RENDERER/);

  const humanizerSource = await readFile(new URL("../src/ui/humanize-params.ts", import.meta.url), "utf8");
  assert.match(humanizerSource, /TOOL_NAMES_WITH_HUMANIZER/);
  assert.match(humanizerSource, /HUMANIZABLE_TOOL_NAME_SET/);

  const uniqueNames = new Set(UI_TOOL_NAMES);
  assert.equal(uniqueNames.size, UI_TOOL_NAMES.length);
  assert.ok(UI_TOOL_NAMES.includes("execute_office_js"));
  assert.ok(UI_TOOL_NAMES.includes("execute_wps_js"));
});

void test("UI metadata drives renderer and humanizer tool subsets", () => {
  assert.equal(TOOL_NAMES_WITH_RENDERER.length > 0, true);
  assert.equal(TOOL_NAMES_WITH_HUMANIZER.length > 0, true);

  for (const name of TOOL_NAMES_WITH_RENDERER) {
    assert.equal(TOOL_UI_METADATA[name].renderer, true);
  }

  for (const name of TOOL_NAMES_WITH_HUMANIZER) {
    assert.equal(TOOL_UI_METADATA[name].humanizer, true);
  }
});

void test("disclosure bundles keep shared core safety tools and add category-specific tools", () => {
  for (const [bundleName, names] of Object.entries(TOOL_DISCLOSURE_BUNDLES)) {
    if (bundleName === "full") continue;

    assert.equal(names.includes("skills"), true);
    assert.equal(names.includes("instructions"), true);
    assert.equal(names.includes("workbook_history"), true);
  }

  assert.equal(TOOL_DISCLOSURE_BUNDLES.analysis.includes("trace_dependencies"), true);
  assert.equal(TOOL_DISCLOSURE_BUNDLES.analysis.includes("explain_formula"), true);
  assert.equal(TOOL_DISCLOSURE_BUNDLES.formatting.includes("format_cells"), true);
  assert.equal(TOOL_DISCLOSURE_BUNDLES.formatting.includes("view_settings"), true);
  assert.equal(TOOL_DISCLOSURE_BUNDLES.structure.includes("modify_structure"), true);
  assert.equal(TOOL_DISCLOSURE_BUNDLES.comments.includes("comments"), true);
});

void test("disclosure trigger patterns cover full-access plus category intents", () => {
  assert.equal(TOOL_DISCLOSURE_FULL_ACCESS_PATTERNS.length > 0, true);
  assert.equal(TOOL_DISCLOSURE_TRIGGER_PATTERNS.comments.length > 0, true);
  assert.equal(TOOL_DISCLOSURE_TRIGGER_PATTERNS.analysis.length > 0, true);
  assert.equal(TOOL_DISCLOSURE_TRIGGER_PATTERNS.structure.length > 0, true);
  assert.equal(TOOL_DISCLOSURE_TRIGGER_PATTERNS.formatting.length > 0, true);
});
