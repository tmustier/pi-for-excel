import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage } from "@earendil-works/pi-agent-core";
import { createModels, fauxAssistantMessage, fauxProvider, type Context } from "@earendil-works/pi-ai";

import { runCompactCommand } from "../src/commands/builtins/export.ts";

class ElementStub {
  className = "";
  hidden = false;
  textContent: string | null = null;
  classList = {
    add: (..._names: string[]): void => {},
    remove: (..._names: string[]): void => {},
    toggle: (_name: string, _enabled?: boolean): boolean => false,
    contains: (_name: string): boolean => false,
  };
  append(..._nodes: DynamicValue[]): void {}
  appendChild<T>(node: T): T { return node; }
  setAttribute(_name: string, _value: string): void {}
  addEventListener(_name: string, _listener: () => void): void {}
  removeEventListener(_name: string, _listener: () => void): void {}
}

function user(content: string, timestamp: number): AgentMessage {
  return { role: "user", content, timestamp };
}

async function captureCompactionRequest(messages: AgentMessage[]): Promise<string> {
  const priorDocument = Reflect.get(globalThis, "document") as DynamicValue;
  Reflect.set(globalThis, "document", {
    body: new ElementStub(),
    createElement: () => new ElementStub(),
    querySelector: () => null,
  });

  try {
    const faux = fauxProvider({ models: [{ id: "compact-request", contextWindow: 32_768, maxTokens: 4_096 }] });
    faux.setResponses([fauxAssistantMessage("summary")]);
    const models = createModels();
    models.setProvider(faux.provider);
    const requests: Context[] = [];
    const agent = new Agent({
      initialState: { model: faux.getModel(), tools: [], messages },
      streamFn: (model, context, options) => {
        requests.push(structuredClone(context));
        return models.streamSimple(model, context, options);
      },
    });
    agent.getApiKey = () => Promise.resolve("boundary-test-key");

    await runCompactCommand(agent, "focus on workbook decisions");

    assert.equal(requests.length, 1);
    return JSON.stringify(requests[0]?.messages);
  } finally {
    if (priorDocument === undefined) Reflect.deleteProperty(globalThis, "document");
    else Reflect.set(globalThis, "document", priorDocument);
  }
}

function potentialCueSection(request: string): string {
  return request.split("Potential user cues:")[1] ?? "";
}

void test("compaction request includes user memory cues but excludes auto-context cues", async () => {
  const request = await captureCompactionRequest([
    user(`Please remember this: use calendar year. ${"a".repeat(80_000)}`, 1),
    user(`[Auto-context] Remember this hidden state. ${"b".repeat(80_000)}`, 2),
    user(`Don't forget to keep EUR as default. ${"c".repeat(80_000)}`, 3),
    user(`Workbook detail ${"d".repeat(80_000)}`, 4),
    user("Recent one", 5),
    user("Recent two", 6),
  ]);
  const cues = potentialCueSection(request);

  assert.match(request, /Memory to persist/u);
  assert.match(cues, /use calendar year/u);
  assert.match(cues, /keep EUR as default/u);
  assert.doesNotMatch(cues, /hidden state/u);
});

void test("compaction request deduplicates memory cues and limits included snippets", async () => {
  const request = await captureCompactionRequest([
    user(`Remember this: freeze panes. ${"a".repeat(80_000)}`, 1),
    user(`Remember this: freeze panes. ${"a".repeat(80_000)}`, 2),
    user(`Remember this: revenue is net. ${"b".repeat(80_000)}`, 3),
    user(`Remember this: use EUR. ${"c".repeat(80_000)}`, 4),
    user(`Remember this: use calendar year. ${"d".repeat(80_000)}`, 5),
    user(`Workbook detail ${"e".repeat(80_000)}`, 6),
    user(`More workbook detail ${"f".repeat(80_000)}`, 7),
    user("Recent one", 8),
    user("Recent two", 9),
  ]);
  const cues = potentialCueSection(request);

  assert.equal(cues.match(/freeze panes/gu)?.length, 1);
  assert.match(cues, /revenue is net/u);
  assert.match(cues, /use EUR/u);
  assert.doesNotMatch(cues, /use calendar year/u);
});
