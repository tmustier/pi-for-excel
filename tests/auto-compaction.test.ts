import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage } from "@earendil-works/pi-agent-core";
import type { Api, Model, ToolResultMessage } from "@earendil-works/pi-ai";

import {
  maybeAutoCompactBeforeContinuation,
  shouldAutoCompactForProjectedTokens,
} from "../src/compaction/auto-compaction.ts";
import {
  buildCompactionMemoryFocusInstruction,
  collectCompactionMemoryCues,
  mergeCompactionAdditionalFocus,
} from "../src/compaction/memory-nudge.ts";

function createUserMessage(text: string, timestamp: number): AgentMessage {
  return {
    role: "user",
    content: text,
    timestamp,
  };
}

function createAssistantMessage(text: string, timestamp: number): AgentMessage {
  return {
    role: "assistant",
    content: [{ type: "text", text }],
    timestamp,
  };
}

function createCompleteAssistantMessage(text: string, timestamp: number): AgentMessage {
  return {
    role: "assistant",
    content: [{ type: "text", text }],
    api: "openai-completions",
    provider: "test",
    model: "test-model",
    usage: {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
      totalTokens: 0,
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
    },
    stopReason: "toolUse",
    timestamp,
  };
}

function createToolResultMessage(text: string, timestamp: number): ToolResultMessage {
  return {
    role: "toolResult",
    toolCallId: `call-${timestamp}`,
    toolName: "read_range",
    content: [{ type: "text", text }],
    isError: false,
    timestamp,
  };
}

function createModel(contextWindow: number): Model<Api> {
  return {
    id: "test-model",
    name: "Test Model",
    api: "openai-completions",
    provider: "test",
    baseUrl: "https://example.invalid",
    reasoning: false,
    input: ["text"],
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
    contextWindow,
    maxTokens: 4096,
  };
}

function createAgent(args: { contextWindow: number; messages: AgentMessage[] }): Agent {
  return new Agent({
    initialState: {
      model: createModel(args.contextWindow),
      messages: args.messages,
      tools: [],
    },
  });
}

// 32k window → hard trigger at 16,384 tokens (65,536 chars).
function createToolLoopMessages(toolResultChars: number): AgentMessage[] {
  return [
    createUserMessage("analyze this data", 1),
    createCompleteAssistantMessage("reading…", 2),
    createToolResultMessage("small output", 3),
    createCompleteAssistantMessage("reading more…", 4),
    createToolResultMessage("y".repeat(toolResultChars), 5),
  ];
}

void test("mid-turn check compacts and returns a replacement loop context", async () => {
  const agent = createAgent({
    contextWindow: 32_768,
    messages: createToolLoopMessages(80_000),
  });

  const keptTail = agent.state.messages.slice(-2);
  let compactRuns = 0;

  const update = await maybeAutoCompactBeforeContinuation({
    agent,
    enabled: true,
    runCompact: () => {
      compactRuns += 1;
      agent.state.messages = [createUserMessage("compaction summary", 6), ...keptTail];
      return Promise.resolve();
    },
  });

  assert.equal(compactRuns, 1);
  assert.notEqual(update, undefined);
  assert.equal(update?.context?.messages.length, 3);
  assert.equal(update?.context?.messages[update.context.messages.length - 1]?.role, "toolResult");
});

void test("mid-turn check is a no-op when disabled or under threshold", async () => {
  let compactRuns = 0;
  const runCompact = () => {
    compactRuns += 1;
    return Promise.resolve();
  };

  const overBudget = createAgent({
    contextWindow: 32_768,
    messages: createToolLoopMessages(80_000),
  });
  assert.equal(
    await maybeAutoCompactBeforeContinuation({ agent: overBudget, enabled: false, runCompact }),
    undefined,
  );

  const underBudget = createAgent({
    contextWindow: 32_768,
    messages: createToolLoopMessages(1_000),
  });
  assert.equal(
    await maybeAutoCompactBeforeContinuation({ agent: underBudget, enabled: true, runCompact }),
    undefined,
  );

  assert.equal(compactRuns, 0);
});

void test("mid-turn check skips when the turn is not continuing (no trailing tool result)", async () => {
  let compactRuns = 0;
  const agent = createAgent({
    contextWindow: 32_768,
    messages: [
      ...createToolLoopMessages(80_000),
      createCompleteAssistantMessage("done", 6),
    ],
  });

  const update = await maybeAutoCompactBeforeContinuation({
    agent,
    enabled: true,
    runCompact: () => {
      compactRuns += 1;
      return Promise.resolve();
    },
  });

  assert.equal(update, undefined);
  assert.equal(compactRuns, 0);
});

void test("mid-turn check returns undefined when compaction does not rewrite history", async () => {
  const agent = createAgent({
    contextWindow: 32_768,
    messages: createToolLoopMessages(80_000),
  });

  const update = await maybeAutoCompactBeforeContinuation({
    agent,
    enabled: true,
    runCompact: () => {
      // compaction failed / nothing to do
      return Promise.resolve();
    },
  });

  assert.equal(update, undefined);
});

void test("does not trigger auto-compaction before hard threshold", () => {
  const shouldCompact = shouldAutoCompactForProjectedTokens({
    projectedTokens: 169_999,
    contextWindow: 200_000,
  });

  assert.equal(shouldCompact, false);
});

void test("triggers auto-compaction after hard threshold", () => {
  const shouldCompact = shouldAutoCompactForProjectedTokens({
    projectedTokens: 170_001,
    contextWindow: 200_000,
  });

  assert.equal(shouldCompact, true);
});

void test("small context windows still use reserve-based hard threshold", () => {
  const below = shouldAutoCompactForProjectedTokens({
    projectedTokens: 16_384,
    contextWindow: 32_768,
  });
  const above = shouldAutoCompactForProjectedTokens({
    projectedTokens: 16_385,
    contextWindow: 32_768,
  });

  assert.equal(below, false);
  assert.equal(above, true);
});

void test("collects memory cues from user messages and ignores auto-context", () => {
  const messages: AgentMessage[] = [
    createUserMessage("[Auto-context] Please remember this summary.", 1),
    createUserMessage("Please remember this: this workbook uses calendar year.", 2),
    createAssistantMessage("Got it.", 3),
    createUserMessage("Don't forget to keep EUR as the default currency.", 4),
  ];

  const summary = collectCompactionMemoryCues(messages);

  assert.equal(summary.cueCount, 2);
  assert.equal(summary.snippets.length, 2);
  assert.ok(summary.snippets.every((snippet) => !snippet.startsWith("[Auto-context]")));
  assert.match(summary.snippets[0] ?? "", /remember this/i);
  assert.match(summary.snippets[1] ?? "", /don['’]t forget/i);
});

void test("deduplicates snippets and respects snippet limits", () => {
  const messages: AgentMessage[] = [
    createUserMessage("Remember this: freeze panes on Summary.", 1),
    createUserMessage("Remember this: freeze panes on Summary.", 2),
    createUserMessage("Please save this for future reference: Revenue is net of refunds.", 3),
    createUserMessage("Please save this for future reference: Revenue is net of refunds.", 4),
  ];

  const summary = collectCompactionMemoryCues(messages, 1);

  assert.equal(summary.cueCount, 4);
  assert.equal(summary.snippets.length, 1);
});

void test("builds memory focus instructions only when cues are present", () => {
  const noCueInstruction = buildCompactionMemoryFocusInstruction({
    cueCount: 0,
    snippets: [],
  });
  assert.equal(noCueInstruction, null);

  const instruction = buildCompactionMemoryFocusInstruction({
    cueCount: 2,
    snippets: [
      "Remember this: workbook uses calendar year.",
      "Don't forget EUR defaults.",
    ],
  });

  assert.notEqual(instruction, null);
  assert.match(instruction ?? "", /instructions tool/i);
  assert.match(instruction ?? "", /notes\//i);
  assert.match(instruction ?? "", /Memory to persist/i);
  assert.match(instruction ?? "", /Potential user cues/i);
});

void test("merges compaction focus parts", () => {
  const merged = mergeCompactionAdditionalFocus(
    "focus on formulas",
    null,
    "capture durable memory",
  );

  assert.equal(merged, "focus on formulas\n\ncapture durable memory");
  assert.equal(mergeCompactionAdditionalFocus(" ", null, undefined), undefined);
});
