import assert from "node:assert/strict";
import { test } from "node:test";

import { Type } from "@sinclair/typebox";

import {
  composeCoreToolsForHost,
  isCoreToolUnsupportedOnWps,
  selectOfficeCoupledToolForHost,
  type AnyHostSelectableTool,
} from "../src/tools/host-selection.ts";
import { UnsupportedHostToolError } from "../src/tools/unsupported-host-tool.ts";
import { CORE_TOOL_NAMES, type CoreToolName } from "../src/tools/names.ts";

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

function createFakeToolFactory(): {
  factory: (name: CoreToolName) => AnyHostSelectableTool;
  createdTools: Map<CoreToolName, AnyHostSelectableTool>;
} {
  const createdTools = new Map<CoreToolName, AnyHostSelectableTool>();
  return {
    factory: (name: CoreToolName) => {
      const tool = createFakeTool(name);
      createdTools.set(name, tool);
      return tool;
    },
    createdTools,
  };
}

void test("composeCoreToolsForHost keeps Office/browser handlers untouched", () => {
  for (const hostKind of ["office", "browser"] as const) {
    const { factory, createdTools } = createFakeToolFactory();
    const tools = composeCoreToolsForHost(factory, hostKind);

    for (const [index, name] of CORE_TOOL_NAMES.entries()) {
      assert.equal(tools[index], createdTools.get(name));
    }
  }
});

void test("Office-coupled non-core tools fail fast on WPS and pass through elsewhere", async () => {
  const officeCoupledTool = createFakeTool("execute_office_js");

  assert.equal(selectOfficeCoupledToolForHost(officeCoupledTool, "office"), officeCoupledTool);
  assert.equal(selectOfficeCoupledToolForHost(officeCoupledTool, "browser"), officeCoupledTool);

  const wpsTool = selectOfficeCoupledToolForHost(officeCoupledTool, "wps");
  assert.notEqual(wpsTool, officeCoupledTool);
  assert.equal(wpsTool.name, officeCoupledTool.name);

  await assert.rejects(
    async () => wpsTool.execute("tool-call-1", {}),
    (error: DynamicValue) => {
      assert.ok(error instanceof UnsupportedHostToolError);
      assert.equal(error.hostKind, "wps");
      assert.equal(error.toolName, "execute_office_js");
      return true;
    },
  );
});

void test("WPS leaves local settings/skills and Phase 2 override core tools available", () => {
  assert.equal(isCoreToolUnsupportedOnWps("read_range"), false);
  assert.equal(isCoreToolUnsupportedOnWps("write_cells"), false);
  assert.equal(isCoreToolUnsupportedOnWps("get_workbook_overview"), false);
  assert.equal(isCoreToolUnsupportedOnWps("workbook_history"), true);
  assert.equal(isCoreToolUnsupportedOnWps("instructions"), false);
  assert.equal(isCoreToolUnsupportedOnWps("conventions"), false);
  assert.equal(isCoreToolUnsupportedOnWps("skills"), false);
});
