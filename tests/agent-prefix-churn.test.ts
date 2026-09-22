import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentTool } from "@earendil-works/pi-agent-core";
import {
  createModels,
  fauxAssistantMessage,
  fauxProvider,
} from "@earendil-works/pi-ai";
import { Type } from "typebox";

import {
  createOfficeStreamFn,
  getLastContext,
  getPayloadSnapshots,
} from "../src/auth/stream-proxy.ts";
import { setDebugEnabled } from "../src/debug/debug.ts";

void test("Agent transcript diagnostics report stable prefixes and model, prompt, and tool changes", async () => {
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
    faux.setResponses([
      fauxAssistantMessage("first"),
      fauxAssistantMessage("second"),
      fauxAssistantMessage("third"),
    ]);
    const models = createModels();
    models.setProvider(faux.provider);

    const parameters = Type.Object({ query: Type.String() });
    const searchTool: AgentTool<typeof parameters, undefined> = {
      name: "search_docs",
      label: "Search docs",
      description: "Search docs",
      parameters,
      execute: () => Promise.resolve({
        content: [{ type: "text", text: "unused" }],
        details: undefined,
      }),
    };
    const sessionId = "agent-transcript-prefix-regression";
    const agent = new Agent({
      initialState: {
        model: faux.getModel("prefix-a"),
        messages: [{ role: "system", content: "base", timestamp: 0 }],
        tools: [searchTool],
      },
      streamFn: createOfficeStreamFn(() => Promise.resolve(undefined), models),
    });
    agent.sessionId = sessionId;

    await agent.prompt("first turn");
    await agent.prompt("stable turn");

    agent.state.model = faux.getModel("prefix-b");
    agent.state.tools = [];
    agent.state.messages = [
      ...agent.state.messages,
      { role: "system", content: "changed", timestamp: 1 },
    ];
    await agent.prompt("changed turn");

    const snapshots = getPayloadSnapshots().filter(
      (snapshot) => snapshot.sessionId === sessionId,
    );
    assert.deepEqual(
      snapshots.map((snapshot) => snapshot.prefixChangeReasons),
      [[], [], ["model", "systemPrompt", "tools"]],
    );
    assert.equal(snapshots[0]?.systemChars, "base".length);
    assert.equal(snapshots[0]?.toolCount, 1);
    assert.equal(snapshots[0]?.toolBundle, "full");
    assert.equal(snapshots[0]?.toolsIncluded, true);
    assert.equal(snapshots[2]?.systemChars, "base\n\nchanged".length);
    assert.equal(snapshots[2]?.toolCount, 0);
    assert.equal(snapshots[2]?.toolBundle, "none");
    assert.equal(snapshots[2]?.toolsIncluded, false);

    const captured = getLastContext(sessionId);
    assert.equal(captured?.systemPrompt, "base\n\nchanged");
    assert.deepEqual(captured?.tools, []);
  } finally {
    setDebugEnabled(false);
    if (priorDocument === undefined) Reflect.deleteProperty(globalThis, "document"); else Reflect.set(globalThis, "document", priorDocument);
    if (priorWindow === undefined) Reflect.deleteProperty(globalThis, "window"); else Reflect.set(globalThis, "window", priorWindow);
  }
});
