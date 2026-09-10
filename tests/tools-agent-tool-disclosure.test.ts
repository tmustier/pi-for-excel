import assert from "node:assert/strict";
import { test } from "node:test";

import { createModels, fauxAssistantMessage, fauxProvider, type Context, type Tool } from "@earendil-works/pi-ai";
import { Type } from "typebox";

import { createOfficeStreamFn } from "../src/auth/stream-proxy.ts";
import { CORE_TOOL_NAMES } from "../src/tools/names.ts";

function createTool(name: string): Tool {
  return { name, description: `${name} tool`, parameters: Type.Object({}) };
}

async function toolsSent(context: Context): Promise<string[] | undefined> {
  const faux = fauxProvider();
  const models = createModels();
  models.setProvider(faux.provider);
  let sentTools: string[] | undefined;
  faux.setResponses([(request) => {
    sentTools = request.tools?.map((tool) => tool.name);
    return fauxAssistantMessage("done");
  }]);
  const hadDocument = Reflect.has(globalThis, "document");
  const previousDocument = Reflect.get(globalThis, "document");
  Reflect.set(globalThis, "document", new EventTarget());
  try {
    const stream = await createOfficeStreamFn(() => Promise.resolve(undefined), models)(
      faux.getModel(),
      context,
    );
    await stream.result();
    return sentTools;
  } finally {
    if (hadDocument) Reflect.set(globalThis, "document", previousDocument);
    else Reflect.deleteProperty(globalThis, "document");
  }
}

void test("agent request omits tools when none are configured", async () => {
  assert.equal(await toolsSent({ messages: [] }), undefined);
});

void test("agent request keeps every core tool visible", async () => {
  const tools = CORE_TOOL_NAMES.map((name) => createTool(name));
  assert.deepEqual(await toolsSent({ messages: [], tools }), [...CORE_TOOL_NAMES]);
});

void test("agent request keeps extension tools alongside core tools", async () => {
  const tools = [...CORE_TOOL_NAMES.map((name) => createTool(name)), createTool("web_search")];
  assert.deepEqual(await toolsSent({ messages: [], tools }), [...CORE_TOOL_NAMES, "web_search"]);
});
