import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage } from "@earendil-works/pi-agent-core";
import {
  createModels,
  fauxAssistantMessage,
  fauxProvider,
  type Context,
} from "@earendil-works/pi-ai";

import { createActionQueue } from "../src/taskpane/action-queue.ts";

function user(text: string, timestamp: number): AgentMessage {
  return { role: "user", content: text, timestamp };
}

function createHarness(contextWindow: number, responses: ReturnType<typeof fauxAssistantMessage>[]) {
  const faux = fauxProvider({ models: [{ id: "request-test", contextWindow, maxTokens: 4_096 }] });
  const models = createModels();
  models.setProvider(faux.provider);
  faux.setResponses(responses);
  const requests: Context[] = [];
  const model = faux.getModel();
  const agent = new Agent({
    initialState: { model, messages: [], tools: [] },
    streamFn: (requestModel, context, options) => {
      requests.push(structuredClone(context));
      return models.streamSimple(requestModel, context, options);
    },
  });
  return { agent, requests };
}

function createQueue(agent: Agent, runCompact: () => Promise<void>) {
  return createActionQueue({
    agent,
    autoCompactEnabled: true,
    runCompact,
    sidebar: { setBusyIndicator: () => {} },
    queueDisplay: { setActionQueue: () => {} },
  });
}

async function waitForIdle(queue: ReturnType<typeof createQueue>): Promise<void> {
  const started = Date.now();
  while (queue.isBusy()) {
    if (Date.now() - started > 2_000) throw new Error("Timed out waiting for agent request");
    await new Promise<void>((resolve) => setTimeout(resolve, 5));
  }
}

void test("a 200k model compacts before sending a request above the 170k hard cap", async () => {
  const runtime = createHarness(200_000, [fauxAssistantMessage("done")]);
  runtime.agent.state.messages = [
    user("a".repeat(680_004), 1),
    user("second", 2),
    user("third", 3),
    user("fourth", 4),
  ];
  let compactRuns = 0;
  const queue = createQueue(runtime.agent, () => {
    compactRuns += 1;
    runtime.agent.state.messages = [user("compacted history", 5)];
    return Promise.resolve();
  });

  queue.enqueuePrompt("continue");
  await waitForIdle(queue);

  assert.equal(compactRuns, 1);
  assert.equal(runtime.requests.length, 1);
  assert.match(JSON.stringify(runtime.requests[0]?.messages), /compacted history/u);
  assert.doesNotMatch(JSON.stringify(runtime.requests[0]?.messages), /a{100}/u);
  queue.shutdown();
});

void test("an overflow is compacted and retried once, then the retry failure is surfaced", async () => {
  const overflow = fauxAssistantMessage("", {
    stopReason: "error",
    errorMessage: "Requested token count exceeds the model's maximum context length of 65536 tokens",
  });
  const runtime = createHarness(65_536, [overflow, overflow]);
  let compactRuns = 0;
  const queue = createQueue(runtime.agent, () => {
    compactRuns += 1;
    runtime.agent.state.messages = [
      user("compacted history", 10),
      {
        role: "toolResult",
        toolCallId: "call-1",
        toolName: "read_range",
        content: [{ type: "text", text: "kept result" }],
        isError: false,
        timestamp: 11,
      },
    ];
    return Promise.resolve();
  });

  queue.enqueuePrompt("analyze workbook");
  await waitForIdle(queue);

  assert.equal(compactRuns, 1);
  assert.equal(runtime.requests.length, 2);
  assert.match(JSON.stringify(runtime.requests[1]?.messages), /compacted history/u);
  const final = runtime.agent.state.messages.at(-1);
  assert.equal(final?.role, "assistant");
  if (final?.role !== "assistant") assert.fail("expected surfaced assistant failure");
  assert.equal(final.stopReason, "error");
  assert.match(final.errorMessage ?? "", /maximum context length/u);
  queue.shutdown();
});
