import assert from "node:assert/strict";
import { test } from "node:test";
import {
  buildCoreToolPromptLines,
  CORE_TOOL_CAPABILITIES,
  filterToolsForDisclosureBundle,
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
import { humanizeToolInput } from "../src/ui/humanize-params.ts";
import { getToolRenderer } from "../src/ui/messages/tool-renderer-registry.ts";
import "../src/ui/tool-renderers.ts";

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

void test("every core tool has a specific renderer and humanized representative input", () => {
  const representativeInput = {
    action: "list",
    cell: "Sheet1!A1",
    content: "Review this calculation",
    formula: "=SUM(B2:B4)",
    level: "workbook",
    name: "financial-modeling",
    query: "revenue",
    range: "Sheet1!A1:B4",
    sheet: "Sheet1",
    source_range: "Sheet1!A1:B4",
    start_cell: "Sheet1!A1",
    values: [["Revenue", 42]],
  };

  for (const name of CORE_TOOL_NAMES) {
    const renderer = getToolRenderer(name);
    assert.ok(renderer, `Core tool ${name} must have a registered renderer`);
    assert.notEqual(
      humanizeToolInput(name, representativeInput),
      null,
      `Core tool ${name} must humanize representative input instead of showing generic JSON`,
    );
  }

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

void test("every disclosed tool bundle resolves to registered core tools", () => {
  const registeredTools = CORE_TOOL_NAMES.map((name) => ({ name }));
  const registeredNames = new Set(CORE_TOOL_NAMES);

  for (const bundleName of ["core", "analysis", "formatting", "structure", "comments", "full"] as const) {
    const names = TOOL_DISCLOSURE_BUNDLES[bundleName];
    for (const name of names) {
      assert.equal(registeredNames.has(name), true, `${bundleName} discloses unregistered tool ${name}`);
    }

    const disclosedNames = new Set<string>(names);
    assert.deepEqual(
      filterToolsForDisclosureBundle(registeredTools, bundleName),
      registeredTools.filter((tool) => disclosedNames.has(tool.name)),
    );
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
