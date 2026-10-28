// Transcript-as-turns contracts (src/ui/messages/turn-model.ts) and the
// rows a turn group draws (src/ui/tool-rows.ts). Native rendering is covered
// by tests/browser/ui-transcript.browser-test.ts.
import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { AssistantMessage, ToolResultMessage } from "@earendil-works/pi-ai";

import { buildTranscriptBlocks } from "../src/ui/messages/turn-model.ts";
import { toolActionVerb } from "../src/ui/tool-rows.ts";

const usage = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } };

function assistant(timestamp: number, content: AssistantMessage["content"], stopReason: AssistantMessage["stopReason"] = "toolUse"): AssistantMessage {
  return { role: "assistant", content, api: "openai-responses", provider: "openai", model: "fixture", usage, stopReason, timestamp };
}

function result(toolCallId: string, text: string, timestamp: number, extra: Partial<ToolResultMessage> = {}): ToolResultMessage {
  return { role: "toolResult", toolCallId, toolName: "fill_formula", content: [{ type: "text", text }], isError: false, timestamp, ...extra };
}

const call = (id: string, range: string) => ({ type: "toolCall" as const, id, name: "fill_formula", arguments: { range, formula: "=1" } });

function build(messages: AgentMessage[], options: { streaming?: AssistantMessage | null; running?: boolean; pending?: string[] } = {}) {
  const results = new Map<string, ToolResultMessage>();
  for (const message of messages) if (message.role === "toolResult") results.set(message.toolCallId, message);
  return buildTranscriptBlocks({
    messages,
    streamingMessage: options.streaming ?? null,
    tools: [],
    results,
    pendingToolCalls: new Set(options.pending ?? []),
    running: options.running ?? false,
  });
}

function shape(blocks: ReturnType<typeof build>): string[] {
  return blocks.map((block) => {
    if (block.type === "turnGroup") {
      const steps = block.group.steps.map((step) => step.type === "tool" ? `${step.toolCall.id}:${step.state}` : `note:${step.content.type === "text" ? step.content.text : step.content.thinking}`);
      return `group[${steps.join(",")}]${block.group.endedAt === null ? " live" : ""}${block.group.answerStarted ? " answered" : ""}${block.group.changes ? " receipt" : ""}`;
    }
    const range = block.contentFrom === undefined ? "" : `[${block.contentFrom},${block.contentTo ?? "end"}]`;
    return `${block.message.role}${range}${block.streaming ? " streaming" : ""}`;
  });
}

void test("a turn of several responses makes one group: lead, every call and the prose between them, then the answer", () => {
  const messages: AgentMessage[] = [
    { role: "user", content: "Fill the rows", timestamp: 1 },
    assistant(10, [{ type: "thinking", thinking: "Plan." }, { type: "text", text: "I'll read first." }, call("r1", "A1"), call("r2", "A2")]),
    result("r1", "Filled A1", 11),
    result("r2", "Error: #REF!", 12, { isError: true }),
    assistant(20, [{ type: "text", text: "Now the totals." }, call("r3", "A3")]),
    result("r3", "Filled A3", 21),
    assistant(30, [{ type: "thinking", thinking: "Done." }, { type: "text", text: "All filled." }], "stop"),
    { role: "turnChanges", turnId: "t", timestamp: 31, outcome: "finished", changes: [{ address: "A1", changedCount: 1, backupIds: [] }] },
    { role: "user", content: "Thanks", timestamp: 40 },
    assistant(50, [{ type: "text", text: "You're welcome." }], "stop"),
  ];
  assert.deepEqual(shape(build(messages)), [
    "user",
    "assistant[0,2]",
    "group[r1:done,r2:failed,note:Now the totals.,r3:done] answered receipt",
    "assistant[0,2]",
    "user",
    "assistant[0,1]",
  ]);
});

void test("while the turn runs, calls without a result are live and the streamed response joins the group", () => {
  const first = assistant(10, [call("r1", "A1"), call("r2", "A2")]);
  const messages: AgentMessage[] = [{ role: "user", content: "Go", timestamp: 1 }, first, result("r1", "Filled A1", 11)];
  assert.deepEqual(shape(build(messages, { running: true, pending: ["r2"] })), ["user", "group[r1:done,r2:live] live"]);

  // The next response streams a tool call: its row appears in the same group, live.
  const streaming = assistant(20, [{ type: "text", text: "Next." }, call("r3", "A3")]);
  const done: AgentMessage[] = [...messages, result("r2", "Filled A2", 12)];
  assert.deepEqual(shape(build(done, { running: true, streaming })), ["user", "group[r1:done,r2:done,note:Next.,r3:live] live"]);

  // The answer streams: the group is finished and collapses, the answer draws below it.
  const answer = assistant(30, [{ type: "text", text: "All" }]);
  const withThird: AgentMessage[] = [...done, assistant(20, [call("r3", "A3")]), result("r3", "Filled A3", 21)];
  assert.deepEqual(shape(build(withThird, { running: true, streaming: answer })), ["user", "group[r1:done,r2:done,r3:done] answered", "assistant[0,1] streaming"]);
});

void test("a stopped turn marks calls without a result as stopped and keeps its stop marker after the group", () => {
  const messages: AgentMessage[] = [
    { role: "user", content: "Go", timestamp: 1 },
    assistant(10, [call("r1", "A1"), call("r2", "A2")], "aborted"),
    result("r1", "Filled A1", 11),
  ];
  const blocks = build(messages);
  assert.deepEqual(shape(blocks), ["user", "group[r1:done,r2:stopped]", "assistant[2,2]"]);
  const group = blocks[1];
  assert.ok(group?.type === "turnGroup");
  assert.equal(group.group.endedAt, 11);
});

void test("context updates do not split a turn, but a receipt ends it", () => {
  const blocks = build([
    assistant(10, [call("r1", "A1")]), result("r1", "ok", 11),
    { role: "autoContext", kind: "changes", text: "User edited A2", timestamp: 12 },
    { role: "system", content: "Tool loadout changed", timestamp: 13 },
    assistant(20, [call("r2", "A2")]), result("r2", "ok", 21),
    { role: "turnChanges", turnId: "first", timestamp: 22, outcome: "finished", changes: [] },
    assistant(30, [call("r3", "A3")]), result("r3", "ok", 31),
  ]);
  assert.deepEqual(shape(blocks), ["group[r1:done,r2:done] receipt", "group[r3:done]"]);
});

void test("thinking-only and empty streams keep the group live until answer text starts", () => {
  const messages = [assistant(1000, [call("r1", "A1")]), result("r1", "ok", 4000)];
  for (const content of [[], [{ type: "thinking" as const, thinking: "Check totals" }], [{ type: "text" as const, text: "  " }]]) {
    const blocks = build(messages, { running: true, streaming: assistant(5000, content) });
    assert.equal(shape(blocks)[0], "group[r1:done] live");
  }
  const blocks = build(messages, { running: true, streaming: assistant(5000, [{ type: "text", text: "Done" }]) });
  assert.deepEqual(shape(blocks), ["group[r1:done] answered", "assistant[0,1] streaming"]);
  const group = blocks[0];
  assert.ok(group?.type === "turnGroup");
  assert.equal(group.group.startedAt, 1000);
  assert.equal(group.group.endedAt, 5000);
});

void test("elapsed time covers lead responses and later tool results, not just the last response", () => {
  const blocks = build([
    assistant(1000, [{ type: "text", text: "Planning" }]),
    assistant(2000, [call("r1", "A1")]), result("r1", "ok", 5000),
    assistant(4000, [{ type: "text", text: "Done" }], "stop"),
  ]);
  const group = blocks.find((block) => block.type === "turnGroup");
  assert.ok(group?.type === "turnGroup");
  assert.equal(group.group.startedAt, 1000);
  assert.equal(group.group.endedAt, 5000);
});

void test("a provider failure during a tool call is failed, not stopped or live", () => {
  assert.deepEqual(shape(build([assistant(10, [call("r1", "A1")], "error")], { running: true })),
    ["group[r1:failed]", "assistant[1,1]"]);
});

void test("prose after the last call stays outside the group when another response follows", () => {
  const messages = [
    assistant(10, [call("r1", "A1"), { type: "text" as const, text: "Filled A1." }]),
    result("r1", "ok", 11),
    assistant(20, [{ type: "thinking", thinking: "Check" }], "stop"),
  ];
  assert.deepEqual(shape(build(messages)), ["group[r1:done] answered", "assistant[1,2]", "assistant[0,1]"]);
});

void test("trailing prose is not an answer while its tool calls still run", () => {
  const response = assistant(10, [call("r1", "A1"), { type: "text", text: "Filled A1." }]);
  assert.deepEqual(shape(build([response], { running: true })), ["group[r1:live] live", "assistant[1,2]"]);
  assert.deepEqual(shape(build([response, result("r1", "ok", 20)], { running: true })),
    ["group[r1:done] answered", "assistant[1,2]"]);
});

void test("approval buttons name the action, falling back for extension tools", () => {
  assert.equal(toolActionVerb("modify_structure", { action: "delete_rows", sheet: "Summary", position: 40, count: 13 }), "Delete");
  assert.equal(toolActionVerb("write_cells", { start_cell: "A1", values: [[1]] }), "Write");
  assert.equal(toolActionVerb("sum_rows", {}), null);
});
