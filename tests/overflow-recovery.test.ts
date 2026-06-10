import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage } from "@earendil-works/pi-agent-core";
import type { Api, AssistantMessage, Model, ToolResultMessage, Usage } from "@earendil-works/pi-ai";

import {
  findTrailingContextOverflowError,
  recoverFromContextOverflow,
} from "../src/compaction/overflow-recovery.ts";

const EMPTY_USAGE: Usage = {
  input: 0,
  output: 0,
  cacheRead: 0,
  cacheWrite: 0,
  totalTokens: 0,
  cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
};

// Real error shape from #558 (LiteLLM custom gateway, 65k model).
const LITELLM_OVERFLOW_ERROR =
  "400 litellm.ContextWindowExceededError: litellm.BadRequestError: ContextWindowExceededError: " +
  "Hosted_vllmException - This model's maximum context length is 65536 tokens. However, you requested " +
  "4096 output tokens and your prompt contains at least 61441 input tokens.";

function createModel(contextWindow: number, id = "deepseek-r1-32b"): Model<Api> {
  return {
    id,
    name: id,
    api: "openai-completions",
    provider: "custom-gateway",
    baseUrl: "https://gateway.example.invalid",
    reasoning: false,
    input: ["text"],
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
    contextWindow,
    maxTokens: 4096,
  };
}

function createUser(text: string, timestamp: number): AgentMessage {
  return { role: "user", content: text, timestamp };
}

function createToolResult(text: string, timestamp: number): ToolResultMessage {
  return {
    role: "toolResult",
    toolCallId: `call-${timestamp}`,
    toolName: "read_range",
    content: [{ type: "text", text }],
    isError: false,
    timestamp,
  };
}

function createOverflowError(model: Model<Api>, timestamp: number): AssistantMessage {
  return {
    role: "assistant",
    content: [{ type: "text", text: "" }],
    api: model.api,
    provider: model.provider,
    model: model.id,
    usage: EMPTY_USAGE,
    stopReason: "error",
    errorMessage: LITELLM_OVERFLOW_ERROR,
    timestamp,
  };
}

function createTestAgent(args: {
  model: Model<Api>;
  messages: AgentMessage[];
}): Agent & { continueCalls: number } {
  class RecordingAgent extends Agent {
    continueCalls = 0;

    override continue(): Promise<void> {
      this.continueCalls += 1;
      return Promise.resolve();
    }
  }

  return new RecordingAgent({
    initialState: {
      model: args.model,
      messages: args.messages,
      tools: [],
    },
  });
}

void test("findTrailingContextOverflowError matches LiteLLM overflow from the active model", () => {
  const model = createModel(65_536);
  const failure = createOverflowError(model, 3);

  const found = findTrailingContextOverflowError({
    messages: [createUser("analyze this data", 1), createToolResult("rows...", 2), failure],
    model,
  });

  assert.equal(found, failure);
});

void test("findTrailingContextOverflowError ignores failures from a different model", () => {
  const oldModel = createModel(65_536, "small-model");
  const newModel = createModel(200_000, "big-model");

  const found = findTrailingContextOverflowError({
    messages: [createUser("hi", 1), createOverflowError(oldModel, 2)],
    model: newModel,
  });

  assert.equal(found, null);
});

void test("findTrailingContextOverflowError ignores non-overflow errors and non-error tails", () => {
  const model = createModel(65_536);

  const rateLimited: AssistantMessage = {
    ...createOverflowError(model, 2),
    errorMessage: "429 rate limit exceeded, too many requests",
  };

  assert.equal(
    findTrailingContextOverflowError({
      messages: [createUser("hi", 1), rateLimited],
      model,
    }),
    null,
  );

  assert.equal(
    findTrailingContextOverflowError({
      messages: [createUser("hi", 1), createToolResult("data", 2)],
      model,
    }),
    null,
  );
});

void test("recoverFromContextOverflow drops the failure, compacts, and retries once", async () => {
  const model = createModel(65_536);
  const user = createUser("analyze this data", 1);
  const toolResult = createToolResult("rows...", 2);
  const agent = createTestAgent({
    model,
    messages: [user, toolResult, createOverflowError(model, 3)],
  });

  let compactRuns = 0;
  const recovered = await recoverFromContextOverflow({
    agent,
    runCompact: () => {
      compactRuns += 1;
      // Simulate compaction rewriting history (new array identity, kept tail).
      agent.state.messages = [createUser("compaction summary", 4), toolResult];
      return Promise.resolve();
    },
  });

  assert.equal(recovered, true);
  assert.equal(compactRuns, 1);
  assert.equal(agent.continueCalls, 1);
  assert.equal(
    agent.state.messages.some((m) => m.role === "assistant"),
    false,
  );
});

void test("recoverFromContextOverflow restores the failure when compaction is a no-op", async () => {
  const model = createModel(65_536);
  const failure = createOverflowError(model, 3);
  const agent = createTestAgent({
    model,
    messages: [createUser("hi", 1), createToolResult("rows...", 2), failure],
  });

  const recovered = await recoverFromContextOverflow({
    agent,
    runCompact: () => {
      // Compaction failed / nothing to compact: messages left untouched.
      return Promise.resolve();
    },
  });

  assert.equal(recovered, false);
  assert.equal(agent.continueCalls, 0);

  const last = agent.state.messages[agent.state.messages.length - 1];
  if (last?.role !== "assistant") throw new Error("expected trailing assistant failure");
  assert.equal(last.errorMessage, LITELLM_OVERFLOW_ERROR);
});

void test("recoverFromContextOverflow restores the failure when compaction throws", async () => {
  const model = createModel(65_536);
  const failure = createOverflowError(model, 3);
  const agent = createTestAgent({
    model,
    messages: [createUser("hi", 1), createToolResult("rows...", 2), failure],
  });

  const recovered = await recoverFromContextOverflow({
    agent,
    runCompact: () => Promise.reject(new Error("summarizer exploded")),
  });

  assert.equal(recovered, false);
  assert.equal(agent.continueCalls, 0);

  const last = agent.state.messages[agent.state.messages.length - 1];
  if (last?.role !== "assistant") throw new Error("expected trailing assistant failure");
  assert.equal(last.errorMessage, LITELLM_OVERFLOW_ERROR);
});

void test("recoverFromContextOverflow does nothing without a trailing overflow error", async () => {
  const model = createModel(65_536);
  const agent = createTestAgent({
    model,
    messages: [createUser("hi", 1), createToolResult("rows...", 2)],
  });

  let compactRuns = 0;
  const recovered = await recoverFromContextOverflow({
    agent,
    runCompact: () => {
      compactRuns += 1;
      return Promise.resolve();
    },
  });

  assert.equal(recovered, false);
  assert.equal(compactRuns, 0);
  assert.equal(agent.continueCalls, 0);
});
