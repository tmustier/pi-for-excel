import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { Api, AssistantMessage, Model, ToolResultMessage, Usage, UserMessage } from "@earendil-works/pi-ai/compat";

import { decideRuntimeIdle } from "../src/taskpane/background-verify-idle.ts";
import {
  collectBuiltInModelCandidates,
  resolveBridgeModelSelection,
  type BridgeModelCandidate,
} from "../src/taskpane/background-verify-model.ts";
import {
  assistantTextSnippet,
  buildTranscriptExport,
  summarizeLastToolCall,
} from "../src/taskpane/background-verify-transcript.ts";

const EMPTY_USAGE: Usage = {
  input: 0,
  output: 0,
  cacheRead: 0,
  cacheWrite: 0,
  totalTokens: 0,
  cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
};

function nonReasoningModel(id = "legacy-completions"): Model<Api> {
  return {
    id,
    name: id,
    api: "openai-completions",
    provider: "custom-gateway",
    baseUrl: "https://gateway.example.invalid",
    reasoning: false,
    input: ["text"],
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
    contextWindow: 65_536,
    maxTokens: 4_096,
  };
}

function userMessage(text: string, timestamp: number): UserMessage {
  return { role: "user", content: text, timestamp };
}

function assistantMessage(args: {
  text?: string;
  toolCalls?: Array<{ id: string; name: string; arguments: Record<string, string | number> }>;
  usage?: Partial<Usage>;
  stopReason?: string;
  errorMessage?: string;
  timestamp: number;
}): AssistantMessage {
  const content: AssistantMessage["content"] = [];
  if (args.text !== undefined) content.push({ type: "text", text: args.text });
  for (const call of args.toolCalls ?? []) {
    content.push({ type: "toolCall", id: call.id, name: call.name, arguments: call.arguments });
  }
  return {
    role: "assistant",
    content,
    api: "openai-responses",
    provider: "openai-codex",
    model: "gpt-5.6-sol",
    usage: { ...EMPTY_USAGE, ...args.usage },
    stopReason: args.stopReason ?? "endTurn",
    timestamp: args.timestamp,
    ...(args.errorMessage !== undefined ? { errorMessage: args.errorMessage } : {}),
  };
}

function toolResult(args: {
  toolCallId: string;
  toolName: string;
  text: string;
  isError: boolean;
  timestamp: number;
}): ToolResultMessage {
  return {
    role: "toolResult",
    toolCallId: args.toolCallId,
    toolName: args.toolName,
    content: [{ type: "text", text: args.text }],
    isError: args.isError,
    timestamp: args.timestamp,
  };
}

// ── resolveBridgeModelSelection ──────────────────────────────────────────────

void test("resolveBridgeModelSelection rejects empty provider/modelId", () => {
  const result = resolveBridgeModelSelection({ candidates: [], provider: "  ", modelId: "gpt" });
  assert.equal(result.ok, false);
  if (!result.ok) assert.match(result.error, /non-empty provider and modelId/u);
});

void test("resolveBridgeModelSelection reports no match against empty registry", () => {
  const result = resolveBridgeModelSelection({
    candidates: [],
    provider: "openai-codex",
    modelId: "gpt-5.6-sol",
  });
  assert.equal(result.ok, false);
  if (!result.ok) {
    assert.match(result.error, /No registered model/u);
    assert.equal(result.availableModelCount, 0);
  }
});

void test("resolveBridgeModelSelection resolves gpt-5.6-sol with medium thinking (eval target)", () => {
  const candidates = collectBuiltInModelCandidates();
  const target = candidates.find((c) => c.provider === "openai-codex" && c.id === "gpt-5.6-sol");
  if (!target) {
    assert.fail("registry no longer exposes openai-codex/gpt-5.6-sol — refresh docs/model-updates.md");
    return;
  }

  const explicit = resolveBridgeModelSelection({
    candidates,
    provider: "openai-codex",
    modelId: "gpt-5.6-sol",
    requestedThinkingLevel: "medium",
  });
  assert.equal(explicit.ok, true);
  if (explicit.ok) {
    assert.equal(explicit.thinkingLevel, "medium");
    assert.equal(explicit.requestedThinkingLevel, "medium");
    assert.ok(explicit.supportedThinkingLevels.includes("medium"));
  }

  // No thinking requested → faithful production default for a reasoning model.
  const defaulted = resolveBridgeModelSelection({
    candidates,
    provider: "openai-codex",
    modelId: "gpt-5.6-sol",
  });
  assert.equal(defaulted.ok, true);
  if (defaulted.ok) {
    assert.equal(defaulted.thinkingLevel, "high");
    assert.equal(defaulted.requestedThinkingLevel, null);
  }
});

void test("resolveBridgeModelSelection rejects an unsupported thinking level with the supported set", () => {
  const candidates = collectBuiltInModelCandidates();
  const result = resolveBridgeModelSelection({
    candidates,
    provider: "openai-codex",
    modelId: "gpt-5.6-sol",
    requestedThinkingLevel: "ludicrous",
  });
  assert.equal(result.ok, false);
  if (!result.ok) {
    assert.match(result.error, /not supported/u);
    assert.ok(result.supportedThinkingLevels && result.supportedThinkingLevels.includes("medium"));
  }
});

void test("resolveBridgeModelSelection defaults non-reasoning models to off and rejects thinking", () => {
  const candidates: BridgeModelCandidate[] = [
    { provider: "custom-gateway", id: "legacy-completions", model: nonReasoningModel() },
  ];

  const defaulted = resolveBridgeModelSelection({
    candidates,
    provider: "custom-gateway",
    modelId: "legacy-completions",
  });
  assert.equal(defaulted.ok, true);
  if (defaulted.ok) {
    assert.equal(defaulted.thinkingLevel, "off");
    assert.deepEqual(defaulted.supportedThinkingLevels, ["off"]);
  }

  const rejected = resolveBridgeModelSelection({
    candidates,
    provider: "custom-gateway",
    modelId: "legacy-completions",
    requestedThinkingLevel: "medium",
  });
  assert.equal(rejected.ok, false);
});

// ── decideRuntimeIdle ────────────────────────────────────────────────────────

void test("decideRuntimeIdle waits for a start before ever reporting idle", () => {
  const awaiting = decideRuntimeIdle({
    baselineMessageCount: 2,
    currentMessageCount: 2,
    isBusy: false,
    sawStart: false,
    observedElapsedMs: 0,
    startupGraceMs: 30_000,
  });
  assert.deepEqual(awaiting, { started: false, idle: false, done: false, reason: "awaiting-start" });

  const startTimeout = decideRuntimeIdle({
    baselineMessageCount: 2,
    currentMessageCount: 2,
    isBusy: false,
    sawStart: false,
    observedElapsedMs: 30_000,
    startupGraceMs: 30_000,
  });
  assert.deepEqual(startTimeout, { started: false, idle: false, done: true, reason: "start-timeout" });
});

void test("decideRuntimeIdle reports running while busy and idle once busy clears after start", () => {
  const running = decideRuntimeIdle({
    baselineMessageCount: 2,
    currentMessageCount: 2,
    isBusy: true,
    sawStart: false,
    observedElapsedMs: 100,
    startupGraceMs: 30_000,
  });
  assert.equal(running.reason, "running");
  assert.equal(running.started, true);
  assert.equal(running.done, false);

  const idleByCount = decideRuntimeIdle({
    baselineMessageCount: 2,
    currentMessageCount: 4,
    isBusy: false,
    sawStart: false,
    observedElapsedMs: 100,
    startupGraceMs: 30_000,
  });
  assert.deepEqual(idleByCount, { started: true, idle: true, done: true, reason: "idle" });

  const idleBySticky = decideRuntimeIdle({
    baselineMessageCount: 2,
    currentMessageCount: 2,
    isBusy: false,
    sawStart: true,
    observedElapsedMs: 5_000,
    startupGraceMs: 30_000,
  });
  assert.deepEqual(idleBySticky, { started: true, idle: true, done: true, reason: "idle" });
});

// ── transcript export ────────────────────────────────────────────────────────

function sampleTranscript(): AgentMessage[] {
  return [
    userMessage("Write SMOKE into A1, then tell me what changed", 1),
    assistantMessage({
      text: "Writing now.",
      toolCalls: [
        { id: "call-1", name: "write_range", arguments: { address: "Sheet1!A1", value: "SMOKE" } },
        { id: "call-2", name: "read_range", arguments: { address: "Sheet1!A1" } },
      ],
      usage: { input: 100, output: 40, totalTokens: 140, reasoning: 12 },
      stopReason: "toolUse",
      timestamp: 2,
    }),
    toolResult({ toolCallId: "call-1", toolName: "write_range", text: "ok", isError: false, timestamp: 3 }),
    toolResult({ toolCallId: "call-2", toolName: "read_range", text: "boom", isError: true, timestamp: 4 }),
    assistantMessage({
      text: "I wrote SMOKE into Sheet1!A1.",
      usage: { input: 50, output: 20, totalTokens: 70 },
      stopReason: "endTurn",
      timestamp: 5,
    }),
  ];
}

void test("buildTranscriptExport aggregates counts, tools, usage, and reply", () => {
  const report = buildTranscriptExport(sampleTranscript());

  assert.equal(report.messageCount, 5);
  assert.equal(report.userCount, 1);
  assert.equal(report.assistantCount, 2);
  assert.equal(report.toolResultCount, 2);
  assert.equal(report.toolCallCount, 2);
  assert.equal(report.toolErrorCount, 1);

  assert.deepEqual(report.usage, {
    input: 150,
    output: 60,
    cacheRead: 0,
    cacheWrite: 0,
    reasoning: 12,
    totalTokens: 210,
  });

  const readTool = report.tools.find((tool) => tool.name === "read_range");
  const writeTool = report.tools.find((tool) => tool.name === "write_range");
  assert.deepEqual(readTool, { name: "read_range", calls: 1, errors: 1 });
  assert.deepEqual(writeTool, { name: "write_range", calls: 1, errors: 0 });

  assert.equal(report.reply?.text, "I wrote SMOKE into Sheet1!A1.");
  assert.equal(report.reply?.truncated, false);
  assert.equal(report.lastAssistant?.stopReason, "endTurn");
  assert.deepEqual(report.stopReasons, { toolUse: 1, endTurn: 1 });
  assert.equal(report.lastToolCall?.name, "read_range");
  assert.equal(report.lastToolCall?.status, "error");
});

void test("buildTranscriptExport bounds reply, message text, and message count", () => {
  const messages: AgentMessage[] = [
    userMessage("x".repeat(50), 1),
    assistantMessage({
      text: "y".repeat(50),
      toolCalls: [{ id: "call-1", name: "write_range", arguments: { note: "z".repeat(50) } }],
      timestamp: 2,
    }),
    assistantMessage({ text: "R".repeat(50), timestamp: 3 }),
  ];

  const report = buildTranscriptExport(messages, {
    maxReplyChars: 10,
    maxMessages: 1,
    maxMessageTextChars: 5,
  });

  assert.equal(report.reply?.truncated, true);
  assert.equal(report.reply?.fullLength, 50);
  assert.equal(report.reply?.text.length, 11); // 10 chars + ellipsis
  assert.equal(report.messages.length, 1);
  assert.equal(report.messagesTruncated, true);
  const last = report.messages[0];
  assert.ok(last);
  assert.equal(last.textTruncated, true);
  assert.ok((last.text ?? "").length <= 6);
});

void test("buildTranscriptExport handles an empty transcript", () => {
  const report = buildTranscriptExport([]);
  assert.equal(report.messageCount, 0);
  assert.equal(report.reply, null);
  assert.equal(report.lastAssistant, null);
  assert.equal(report.lastToolCall, null);
  assert.deepEqual(report.tools, []);
  assert.deepEqual(report.stopReasons, {});
});

void test("transcript export and status pieces never leak tool args, user text, or tool-result text", () => {
  const SECRET = "S3CRET-do-not-leak-abc123";
  const messages: AgentMessage[] = [
    userMessage(`please use ${SECRET} to authenticate`, 1),
    assistantMessage({
      text: "Working on it.",
      toolCalls: [{ id: "call-1", name: "http_fetch", arguments: { token: SECRET, url: "https://x" } }],
      stopReason: "toolUse",
      timestamp: 2,
    }),
    toolResult({
      toolCallId: "call-1",
      toolName: "http_fetch",
      text: `response body containing ${SECRET}`,
      isError: false,
      timestamp: 3,
    }),
    assistantMessage({ text: "Done.", stopReason: "endTurn", timestamp: 4 }),
  ];

  // Generous text caps to prove exclusion is structural, not just truncation.
  const report = buildTranscriptExport(messages, { maxReplyChars: 5_000, maxMessageTextChars: 5_000 });
  const serialized = JSON.stringify(report);
  assert.equal(serialized.includes(SECRET), false, "secret from args/user/tool-result must not appear in export");

  // Status path pieces (lastToolCall + assistant snippet) must also be clean.
  assert.equal(JSON.stringify(summarizeLastToolCall(messages)).includes(SECRET), false);
  const lastAssistant = messages[messages.length - 1];
  assert.ok(lastAssistant);
  assert.equal(assistantTextSnippet(lastAssistant).includes(SECRET), false);

  // Structural signal survives without the raw content.
  assert.equal(report.toolCallCount, 1);
  assert.ok(report.tools.find((tool) => tool.name === "http_fetch"));
  const userMsg = report.messages.find((message) => message.role === "user");
  assert.ok(userMsg && (userMsg.textLength ?? 0) > 0 && userMsg.text === undefined);
  const toolMsg = report.messages.find((message) => message.role === "toolResult");
  assert.ok(toolMsg && (toolMsg.textLength ?? 0) > 0 && toolMsg.text === undefined);
});

void test("summarizeLastToolCall reports pending, ok, error, and null", () => {
  assert.equal(summarizeLastToolCall([]), null);

  const pending = summarizeLastToolCall([
    assistantMessage({
      toolCalls: [{ id: "call-1", name: "write_range", arguments: { address: "A1" } }],
      timestamp: 1,
    }),
  ]);
  assert.equal(pending?.status, "pending");
  assert.equal(pending?.name, "write_range");

  const ok = summarizeLastToolCall([
    assistantMessage({
      toolCalls: [{ id: "call-1", name: "write_range", arguments: { address: "A1" } }],
      timestamp: 1,
    }),
    toolResult({ toolCallId: "call-1", toolName: "write_range", text: "ok", isError: false, timestamp: 2 }),
  ]);
  assert.equal(ok?.status, "ok");

  const errored = summarizeLastToolCall([
    assistantMessage({
      toolCalls: [{ id: "call-1", name: "read_range", arguments: { address: "A1" } }],
      timestamp: 1,
    }),
    toolResult({ toolCallId: "call-1", toolName: "read_range", text: "boom", isError: true, timestamp: 2 }),
  ]);
  assert.equal(errored?.status, "error");
});
