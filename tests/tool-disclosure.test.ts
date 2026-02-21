import assert from "node:assert/strict";
import { test } from "node:test";

import type { Context, Tool } from "@mariozechner/pi-ai";
import { Type } from "@sinclair/typebox";

import { selectToolBundle } from "../src/context/tool-disclosure.ts";
import { CORE_TOOL_NAMES } from "../src/tools/names.ts";

function createTool(name: string): Tool {
  return {
    name,
    description: `${name} tool`,
    parameters: Type.Object({}),
  };
}

function createContext(args: {
  prompt?: string;
  tools?: readonly Tool[];
  includeAutoContextMessage?: boolean;
}): Context {
  const messages: Context["messages"] = [];

  if (args.prompt) {
    messages.push({
      role: "user",
      content: args.prompt,
      timestamp: 1,
    });
  }

  if (args.includeAutoContextMessage) {
    messages.push({
      role: "user",
      content: "[Auto-context] Workbook snapshot",
      timestamp: 2,
    });
  }

  return {
    messages,
    tools: args.tools ? [...args.tools] : undefined,
  };
}

function createCoreToolSet(): Tool[] {
  return CORE_TOOL_NAMES.map((name) => createTool(name));
}

void test("selectToolBundle returns none when no tools are present", () => {
  const result = selectToolBundle({ messages: [] });
  assert.equal(result.bundleId, "none");
  assert.equal(result.tools, undefined);
});

void test("selectToolBundle keeps full core tool visibility for cache stability", () => {
  const context = createContext({
    prompt: "Please format this table with borders and color.",
    tools: createCoreToolSet(),
    includeAutoContextMessage: true,
  });

  const result = selectToolBundle(context);

  assert.equal(result.bundleId, "full");
  assert.deepEqual(
    result.tools?.map((tool) => tool.name),
    [...CORE_TOOL_NAMES],
  );
});

void test("selectToolBundle keeps full tools when non-core tools are present", () => {
  const tools = [...createCoreToolSet(), createTool("web_search")];
  const context = createContext({
    prompt: "Please search the web for this data.",
    tools,
  });

  const result = selectToolBundle(context);

  assert.equal(result.bundleId, "full");
  assert.deepEqual(result.tools?.map((tool) => tool.name), tools.map((tool) => tool.name));
});
