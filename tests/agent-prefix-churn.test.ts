import assert from "node:assert/strict";
import { test } from "node:test";

import { Type } from "typebox";
import { createModels, fauxAssistantMessage, fauxProvider, type Context, type Model, type Api } from "@earendil-works/pi-ai";

import { createOfficeStreamFn, getPayloadSnapshots } from "../src/auth/stream-proxy.ts";
import { setDebugEnabled } from "../src/debug/debug.ts";

function context(systemPrompt: string, includeTool: boolean): Context {
  return {
    systemPrompt,
    messages: [],
    ...(includeTool ? { tools: [{ name: "search_docs", description: "Search docs", parameters: Type.Object({ query: Type.String() }) }] } : {}),
  };
}

void test("agent request diagnostics report stable prefixes and model, prompt, and tool changes per session", async () => {
  const priorDocument = Reflect.get(globalThis, "document") as unknown;
  const priorWindow = Reflect.get(globalThis, "window") as unknown;
  Reflect.set(globalThis, "document", new EventTarget());
  const values = new Map<string, string>();
  Reflect.set(globalThis, "window", { localStorage: {
    getItem: (key: string) => values.get(key) ?? null,
    setItem: (key: string, value: string) => values.set(key, value),
  } });

  try {
    setDebugEnabled(true);
    const faux = fauxProvider({ models: [
      { id: "prefix-a", contextWindow: 200_000, maxTokens: 4_096 },
      { id: "prefix-b", contextWindow: 200_000, maxTokens: 4_096 },
    ] });
    faux.setResponses(Array.from({ length: 31 }, () => fauxAssistantMessage("done")));
    const models = createModels();
    models.setProvider(faux.provider);
    const stream = createOfficeStreamFn(() => Promise.resolve(undefined), models);
    const modelA = faux.getModel("prefix-a");
    const modelB = faux.getModel("prefix-b");
    const send = async (model: Model<Api>, requestContext: Context, sessionId: string): Promise<void> => {
      const result = await stream(model, requestContext, { sessionId });
      await result.result();
    };
    await send(modelA, context("base", true), "r3b-prefix-other");
    for (let index = 0; index < 25; index += 1) {
      await send(modelA, context("base", true), `r3b-prefix-evict-${index}`);
    }
    await send(modelA, context("changed", false), "r3b-prefix-other");
    await send(modelA, context("base", true), "r3b-prefix-main");
    await send(modelA, context("base", true), "r3b-prefix-main");
    await send(modelB, context("changed", false), "r3b-prefix-main");

    const snapshots = getPayloadSnapshots();
    const main = snapshots.filter((snapshot) => snapshot.sessionId === "r3b-prefix-main");
    assert.deepEqual(main.map((snapshot) => snapshot.prefixChangeReasons), [[], [], ["model", "systemPrompt", "tools"]]);
    const other = snapshots.filter((snapshot) => snapshot.sessionId === "r3b-prefix-other");
    assert.deepEqual(other.map((snapshot) => snapshot.prefixChangeReasons), [[]]);
  } finally {
    setDebugEnabled(false);
    if (priorDocument === undefined) Reflect.deleteProperty(globalThis, "document"); else Reflect.set(globalThis, "document", priorDocument);
    if (priorWindow === undefined) Reflect.deleteProperty(globalThis, "window"); else Reflect.set(globalThis, "window", priorWindow);
  }
});
