import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage } from "@earendil-works/pi-agent-core";
import { createModels, fauxAssistantMessage, fauxProvider, fauxToolCall, type Context } from "@earendil-works/pi-ai";

import { runCompactCommand } from "../src/commands/builtins/export.ts";
import { createConvertToLlm } from "../src/messages/convert-to-llm.ts";
import { estimateContextTokens } from "../src/utils/context-tokens.ts";

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

function stubDocument(): () => void {
  const priorDocument = Reflect.get(globalThis, "document") as unknown;
  Reflect.set(globalThis, "document", {
    body: new ElementStub(),
    createElement: () => new ElementStub(),
    querySelector: () => null,
  });
  return () => {
    if (priorDocument === undefined) Reflect.deleteProperty(globalThis, "document");
    else Reflect.set(globalThis, "document", priorDocument);
  };
}

void test("compact tells the summarizer an answered request is done and frees the space of its large kept result", async () => {
  const restoreDocument = stubDocument();
  try {
    const faux = fauxProvider({ models: [{ id: "compact-tail", contextWindow: 200_000, maxTokens: 4_096 }] });
    faux.setResponses([fauxAssistantMessage("Summary")]);
    const models = createModels();
    models.setProvider(faux.provider);
    const model = faux.getModel();
    const requests: Context[] = [];

    // Two answered reads with ~22k-token results, then a short exchange.
    const read = (n: number, range: string, chars: number, answer: string, at: number): AgentMessage[] => [
      user(`Read ${range}`, at),
      fauxAssistantMessage(fauxToolCall("read_range", { range }, { id: `call-${n}` }), { stopReason: "toolUse" }),
      {
        role: "toolResult",
        toolCallId: `call-${n}`,
        toolName: "read_range",
        content: [{ type: "text", text: `rows of ${range}: ${"9".repeat(chars)}` }],
        isError: false,
        timestamp: at + 1,
      },
      fauxAssistantMessage(answer),
    ];
    const agent = new Agent({
      initialState: {
        model,
        tools: [],
        messages: [
          ...read(1, "Calc!A1:D2500", 100_000, "First half read: total 1,234.", 1),
          ...read(2, "Calc!A2501:D5000", 88_000, "Second half read: total 5,678.", 10),
          user("Reply with just OK.", 20),
          fauxAssistantMessage("OK"),
        ],
      },
      convertToLlm: createConvertToLlm({ getContextWindow: () => 200_000 }),
      streamFn: (requestModel, context, options) => {
        requests.push(structuredClone(context));
        return models.streamSimple(requestModel, context, options);
      },
    });
    agent.getApiKey = () => Promise.resolve("boundary-test-key");
    const before = estimateContextTokens({ systemPrompt: "", messages: agent.state.messages }).totalTokens;

    await runCompactCommand(agent, "");

    // The second request was cut from its answer; the summarizer still learns it was answered.
    const prompt = JSON.stringify(requests[0]?.messages);
    assert.match(prompt, /\[User request, answered\]: Read Calc!A2501:D5000/u);
    assert.match(prompt, /final reply to the request above, given after these messages\]: Second half read: total 5,678\./u);
    assert.match(prompt, /never under In Progress or Next Steps/u);

    // The kept 22k-token result reaches the model as a preview; the tool card keeps it whole.
    const kept = agent.state.messages.find((message) => message.role === "toolResult" && message.toolCallId === "call-2");
    assert.ok(kept?.role === "toolResult" && kept.content.some((block) => block.type === "text" && block.text.length > 88_000));
    const llm = createConvertToLlm({ getContextWindow: () => 200_000 })(agent.state.messages);
    const sent = llm.find((message) => message.role === "toolResult" && message.toolCallId === "call-2");
    assert.ok(sent?.role === "toolResult");
    assert.match(sent.content.map((block) => (block.type === "text" ? block.text : "")).join(""), /^\[Compacted tool result\] read_range/u);

    // Most of the window is freed, and the context meter agrees.
    const after = estimateContextTokens({ systemPrompt: "", messages: agent.state.messages }).totalTokens;
    assert.ok(before > 45_000, `before: ${before}`);
    assert.ok(after < 2_000, `after: ${after}`);
  } finally {
    restoreDocument();
  }
});
