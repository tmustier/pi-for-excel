import assert from "node:assert/strict";
import { test } from "node:test";
import { readFile } from "node:fs/promises";

import { Type } from "@sinclair/typebox";

import {
  isCoreToolUnsupportedOnWps,
  selectCoreToolForHost,
  type AnyHostSelectableTool,
} from "../src/tools/host-selection.ts";
import { CORE_TOOL_NAMES } from "../src/tools/names.ts";

function createFakeTool(name: string): AnyHostSelectableTool {
  return {
    name,
    label: name,
    description: `${name} description`,
    parameters: Type.Object({}),
    execute: () => Promise.resolve({
      content: [{ type: "text", text: `${name} ok` }],
      details: undefined,
    }),
  };
}

void test("registry creates core tools by mapping CORE_TOOL_NAMES", async () => {
  const registrySource = await readFile(new URL("../src/tools/registry.ts", import.meta.url), "utf8");

  assert.match(registrySource, /CORE_TOOL_NAMES\.map/);
  assert.match(registrySource, /selectCoreToolForHost/);
  assert.equal(CORE_TOOL_NAMES.includes("instructions"), true);
  assert.equal(CORE_TOOL_NAMES.includes("conventions"), true);
  assert.equal(CORE_TOOL_NAMES.includes("skills"), true);
});

void test("host selection keeps Office tool handlers and wraps WPS workbook tools", async () => {
  const officeTool = createFakeTool("read_range");
  const selectedOfficeTool = selectCoreToolForHost("read_range", officeTool, "office");
  assert.equal(selectedOfficeTool, officeTool);

  const wpsTool = selectCoreToolForHost("read_range", officeTool, "wps");
  assert.notEqual(wpsTool, officeTool);
  assert.equal(wpsTool.name, officeTool.name);
  assert.equal(wpsTool.label, officeTool.label);
  assert.equal(wpsTool.description, officeTool.description);

  await assert.rejects(
    async () => wpsTool.execute("tool-call-1", { range: "A1" }),
    /not yet supported on WPS Spreadsheets.*NEXSELL-370/u,
  );
});

void test("WPS leaves local settings/skills core tools available", () => {
  assert.equal(isCoreToolUnsupportedOnWps("read_range"), true);
  assert.equal(isCoreToolUnsupportedOnWps("workbook_history"), true);
  assert.equal(isCoreToolUnsupportedOnWps("instructions"), false);
  assert.equal(isCoreToolUnsupportedOnWps("conventions"), false);
  assert.equal(isCoreToolUnsupportedOnWps("skills"), false);
});
