import assert from "node:assert/strict";
import { test } from "node:test";

import { Type } from "@sinclair/typebox";

import {
  WPS_CORE_TOOL_EXECUTE_OVERRIDES,
  composeCoreToolsForHost,
  isCoreToolUnsupportedOnWps,
  selectCoreToolForHost,
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

void test("composeCoreToolsForHost returns core tools in CORE_TOOL_NAMES order for every host", () => {
  for (const hostKind of ["office", "wps", "browser"] as const) {
    const { factory } = createFakeToolFactory();
    const tools = composeCoreToolsForHost(factory, hostKind);
    assert.deepEqual(tools.map((tool) => tool.name), [...CORE_TOOL_NAMES]);
  }
});

void test("composeCoreToolsForHost keeps Office/browser handlers untouched", () => {
  for (const hostKind of ["office", "browser"] as const) {
    const { factory, createdTools } = createFakeToolFactory();
    const tools = composeCoreToolsForHost(factory, hostKind);

    for (const [index, name] of CORE_TOOL_NAMES.entries()) {
      assert.equal(tools[index], createdTools.get(name));
    }
  }
});

void test("composeCoreToolsForHost keeps metadata stable on WPS and fails fast with a typed error", async () => {
  const { factory, createdTools } = createFakeToolFactory();
  const wpsTools = composeCoreToolsForHost(factory, "wps");

  for (const [index, name] of CORE_TOOL_NAMES.entries()) {
    const wpsTool = wpsTools[index];
    const originalTool = createdTools.get(name);
    assert.ok(originalTool);

    assert.equal(wpsTool.name, originalTool.name);
    assert.equal(wpsTool.label, originalTool.label);
    assert.equal(wpsTool.description, originalTool.description);
    assert.deepEqual(wpsTool.parameters, originalTool.parameters);

    if (WPS_CORE_TOOL_EXECUTE_OVERRIDES[name]) {
      assert.notEqual(wpsTool, originalTool);
      assert.notEqual(wpsTool.execute, originalTool.execute);
    } else if (isCoreToolUnsupportedOnWps(name)) {
      assert.notEqual(wpsTool, originalTool);
      await assert.rejects(
        async () => wpsTool.execute("tool-call-1", {}),
        (error: unknown) => {
          assert.ok(error instanceof UnsupportedHostToolError);
          assert.equal(error.code, "unsupported_host_tool");
          assert.equal(error.hostKind, "wps");
          assert.equal(error.toolName, name);
          assert.match(error.message, /not yet supported on WPS Spreadsheets.*NEXSELL-370/u);
          return true;
        },
      );
    } else {
      assert.equal(wpsTool, originalTool);
    }
  }
});

void test("host selection keeps Office tool handlers and swaps only execute for WPS overrides", () => {
  const officeTool = createFakeTool("read_range");
  const selectedOfficeTool = selectCoreToolForHost("read_range", officeTool, "office");
  assert.equal(selectedOfficeTool, officeTool);

  const browserTool = selectCoreToolForHost("read_range", officeTool, "browser");
  assert.equal(browserTool, officeTool);

  const wpsTool = selectCoreToolForHost("read_range", officeTool, "wps");
  assert.notEqual(wpsTool, officeTool);
  assert.equal(wpsTool.name, officeTool.name);
  assert.equal(wpsTool.label, officeTool.label);
  assert.equal(wpsTool.description, officeTool.description);
  assert.equal(wpsTool.parameters, officeTool.parameters);
  assert.notEqual(wpsTool.execute, officeTool.execute);
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
    (error: unknown) => {
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
