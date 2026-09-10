import assert from "node:assert/strict";
import { test } from "node:test";

import { Type } from "typebox";

import {
  composeCoreToolsForHost,
  type AnyHostSelectableTool,
} from "../src/tools/host-selection.ts";
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

