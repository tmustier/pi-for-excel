import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage } from "@earendil-works/pi-agent-core";
import { createModels, fauxAssistantMessage, fauxProvider, type Context } from "@earendil-works/pi-ai";

import { createActionQueue } from "../src/taskpane/action-queue.ts";

const LITELLM_OVERFLOW =
  "400 litellm.ContextWindowExceededError: ContextWindowExceededError: This model's maximum context length is 65536 tokens. However, your prompt contains 61441 input tokens.";

function user(text: string, timestamp: number): AgentMessage {
  return { role: "user", content: text, timestamp };
}

function createHarness(responses: ReturnType<typeof fauxAssistantMessage>[]) {
  const faux = fauxProvider({ models: [{ id: "deepseek-r1-32b", contextWindow: 65_536, maxTokens: 4_096 }] });
  faux.setResponses(responses);
  const models = createModels();
  models.setProvider(faux.provider);
  const requests: Context[] = [];
  const agent = new Agent({
    initialState: { model: faux.getModel(), messages: [], tools: [] },
    streamFn: (model, context, options) => {
      requests.push(structuredClone(context));
      return models.streamSimple(model, context, options);
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
    if (Date.now() - started > 2_000) throw new Error("Timed out waiting for request recovery");
    await new Promise<void>((resolve) => setTimeout(resolve, 5));
  }
}

void test("LiteLLM context overflow from the active model compacts and retries the agent request", async () => {
  const overflow = fauxAssistantMessage("", { stopReason: "error", errorMessage: LITELLM_OVERFLOW });
  const runtime = createHarness([overflow, fauxAssistantMessage("recovered")]);
  const queue = createQueue(runtime.agent, () => {
    runtime.agent.state.messages = [user("compacted workbook history", 10)];
    return Promise.resolve();
  });

  queue.enqueuePrompt("analyze workbook");
  await waitForIdle(queue);

  assert.equal(runtime.requests.length, 2);
  assert.match(JSON.stringify(runtime.requests[1]?.messages), /compacted workbook history/u);
  assert.equal(runtime.agent.state.messages.at(-1)?.role, "assistant");
  queue.shutdown();
});

void test("non-overflow model errors do not compact or retry the agent request", async () => {
  const runtime = createHarness([
    fauxAssistantMessage("", { stopReason: "error", errorMessage: "429 rate limit exceeded" }),
  ]);
  let compacted = false;
  const queue = createQueue(runtime.agent, () => {
    compacted = true;
    return Promise.resolve();
  });

  queue.enqueuePrompt("analyze workbook");
  await waitForIdle(queue);

  assert.equal(runtime.requests.length, 1);
  assert.equal(compacted, false);
  assert.match(JSON.stringify(runtime.agent.state.messages.at(-1)), /rate limit exceeded/u);
  queue.shutdown();
});

void test("a no-op compaction restores the LiteLLM failure without another request", async () => {
  const runtime = createHarness([
    fauxAssistantMessage("", { stopReason: "error", errorMessage: LITELLM_OVERFLOW }),
  ]);
  const queue = createQueue(runtime.agent, () => Promise.resolve());

  queue.enqueuePrompt("analyze workbook");
  await waitForIdle(queue);

  assert.equal(runtime.requests.length, 1);
  assert.match(JSON.stringify(runtime.agent.state.messages.at(-1)), /ContextWindowExceededError/u);
  queue.shutdown();
});
