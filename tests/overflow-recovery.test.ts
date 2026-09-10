import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { Api, AssistantMessage, Model, ToolResultMessage, Usage } from "@earendil-works/pi-ai";

import { findTrailingContextOverflowError } from "../src/compaction/overflow-recovery.ts";

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

