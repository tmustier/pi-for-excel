import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage } from "@earendil-works/pi-agent-core";
import { createModels, fauxAssistantMessage, fauxProvider, type Context } from "@earendil-works/pi-ai";

import { runCompactCommand } from "../src/commands/builtins/export.ts";

class ElementStub {
  id = "";
  className = "";
  hidden = false;
  textContent: string | null = null;
  type = "";
  classList = {
    add: (..._names: string[]): void => {},
    remove: (..._names: string[]): void => {},
    toggle: (_name: string, _enabled?: boolean): boolean => false,
    contains: (_name: string): boolean => false,
  };
  append(..._nodes: unknown[]): void {}
  appendChild<T>(node: T): T { return node; }
  setAttribute(_name: string, _value: string): void {}
  addEventListener(_name: string, _listener: () => void): void {}
  removeEventListener(_name: string, _listener: () => void): void {}
}

function user(content: string, timestamp: number): AgentMessage {
  return { role: "user", content, timestamp };
}

void test("compact sends user memory cues to the summarizing model", async () => {
  const priorDocument = Reflect.get(globalThis, "document") as unknown;
  const body = new ElementStub();
  Reflect.set(globalThis, "document", {
    body,
    createElement: () => new ElementStub(),
    querySelector: () => null,
  });

  try {
    const faux = fauxProvider({ models: [{ id: "compact-request", contextWindow: 32_768, maxTokens: 4_096 }] });
    faux.setResponses([fauxAssistantMessage("A concise summary")]);
    const models = createModels();
    models.setProvider(faux.provider);
    const requests: Context[] = [];
    const model = faux.getModel();
    const agent = new Agent({
      initialState: {
        model,
        tools: [],
        messages: [
          user(`Please remember this: report in EUR. ${"a".repeat(30_000)}`, 1),
          user(`Earlier workbook detail ${"b".repeat(30_000)}`, 2),
          user(`Another workbook detail ${"c".repeat(30_000)}`, 3),
          user(`More workbook detail ${"d".repeat(30_000)}`, 4),
          user("Recent instruction one", 5),
          user("Recent instruction two", 6),
        ],
      },
      streamFn: (requestModel, context, options) => {
        requests.push(structuredClone(context));
        return models.streamSimple(requestModel, context, options);
      },
    });
    agent.getApiKey = () => Promise.resolve("boundary-test-key");

    await runCompactCommand(agent, "focus on formulas");

    assert.equal(requests.length, 1);
    const requestText = JSON.stringify(requests[0]?.messages);
    assert.match(requestText, /focus on formulas/u);
    assert.match(requestText, /Memory to persist/u);
    assert.match(requestText, /report in EUR/u);
  } finally {
    if (priorDocument === undefined) Reflect.deleteProperty(globalThis, "document");
    else Reflect.set(globalThis, "document", priorDocument);
  }
});
