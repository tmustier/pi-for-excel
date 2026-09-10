import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent } from "@earendil-works/pi-agent-core";
import { createModels, fauxAssistantMessage, fauxProvider, type Context } from "@earendil-works/pi-ai";

import { createActionQueue } from "../src/taskpane/action-queue.ts";

const OVERFLOW = "400 litellm.ContextWindowExceededError: This model's maximum context length is 65536 tokens";

function harness(response: ReturnType<typeof fauxAssistantMessage>, runCompact: () => Promise<void>) {
  const faux = fauxProvider({ models: [{ id: "active-model", contextWindow: 65_536, maxTokens: 4_096 }] });
  faux.setResponses([response]);
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
  const queue = createActionQueue({
    agent,
    autoCompactEnabled: true,
    runCompact,
    sidebar: { setBusyIndicator: () => {} },
    queueDisplay: { setActionQueue: () => {} },
  });
  return { agent, queue, requests };
}

async function run(runtime: ReturnType<typeof harness>): Promise<void> {
  runtime.queue.enqueuePrompt("analyze workbook");
  const started = Date.now();
  while (runtime.queue.isBusy()) {
    if (Date.now() - started > 2_000) throw new Error("Timed out waiting for agent request");
    await new Promise<void>((resolve) => setTimeout(resolve, 5));
  }
  runtime.queue.shutdown();
}

void test("successful response tails do not trigger overflow recovery", async () => {
  let successCompactions = 0;
  const successRuntime = harness(fauxAssistantMessage("done"), () => { successCompactions += 1; return Promise.resolve(); });
  await run(successRuntime);
  assert.equal(successRuntime.requests.length, 1);
  assert.equal(successCompactions, 0);
});

void test("a compaction exception preserves the original overflow failure", async () => {
  const runtime = harness(
    fauxAssistantMessage("", { stopReason: "error", errorMessage: OVERFLOW }),
    () => Promise.reject(new Error("summarizer exploded")),
  );
  await run(runtime);

  assert.equal(runtime.requests.length, 1);
  assert.match(JSON.stringify(runtime.agent.state.messages.at(-1)), /ContextWindowExceededError/u);
});
