// Transcript-as-turns contracts (src/ui/messages/turn-model.ts) and the
// rows a turn group draws (src/ui/tool-rows.ts). Native rendering is covered
// by tests/browser/ui-transcript.browser-test.ts.
import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { AssistantMessage, ToolResultMessage } from "@earendil-works/pi-ai";

import { buildTranscriptBlocks } from "../src/ui/messages/turn-model.ts";
import { describeToolRow, toolActionVerb } from "../src/ui/tool-rows.ts";

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
      const steps = block.group.steps.map((step) => step.type === "tool" ? `${step.toolCall.id}:${step.state}` : `note@${step.contentIndex}`);
      return `group[${steps.join(",")}]${block.group.live ? " live" : ""}${block.group.answerStarted ? " answered" : ""}${block.group.changes ? " receipt" : ""}`;
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
    "group[r1:done,r2:failed,note@0,r3:done] answered receipt",
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
  assert.deepEqual(shape(build(done, { running: true, streaming })), ["user", "group[r1:done,r2:done,note@0,r3:live] live"]);

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

void test("rows read in the tense of their state and carry one fact", () => {
  const filled = result("f", "Filled **Sheet1!B2:B9** — 8 changed", 1, {
    details: { kind: "fill_formula", blocked: false, address: "Sheet1!B2:B9", changes: { changedCount: 8, truncated: false, sample: [] } },
  });
  const params = { range: "Sheet1!B2:B9", formula: "=A2*2" };
  assert.deepEqual(pick(describeToolRow("fill_formula", params, filled, "done")), { text: "Filled Sheet1!B2:B9", meta: "8 cells", kind: "sheet", reason: null });
  assert.deepEqual(pick(describeToolRow("fill_formula", params, undefined, "live")), { text: "Filling Sheet1!B2:B9", meta: null, kind: "sheet", reason: null });
  assert.deepEqual(pick(describeToolRow("fill_formula", params, result("f", "Error: Sheet1!A2 is protected.", 1, { isError: true }), "failed")),
    { text: "Could not fill Sheet1!B2:B9", meta: null, kind: "sheet", reason: "Sheet1!A2 is protected." });
  assert.deepEqual(pick(describeToolRow("fill_formula", params, undefined, "stopped")), { text: "Stopped filling Sheet1!B2:B9", meta: null, kind: "sheet", reason: null });

  const read = result("r", "**Sheet1!A1:C12** (12×3)\n\n| | A |", 1, { toolName: "read_range" });
  assert.deepEqual(pick(describeToolRow("read_range", { range: "Sheet1!A1:C12" }, read, "done")), { text: "Read Sheet1!A1:C12", meta: "36 cells", kind: "sheet", reason: null });

  const search = result("w", "1. ONS", 1, { toolName: "web_search", details: { kind: "web_search", ok: true, provider: "brave", query: "UK CPI", sentQuery: "UK CPI", maxResults: 10, resultCount: 6 } });
  assert.deepEqual(pick(describeToolRow("web_search", { query: "UK CPI" }, search, "done")), { text: "Searched the web for \"UK CPI\"", meta: "6 results", kind: "web", reason: null });
  assert.equal(describeToolRow("python_transform_range", { range: "A1" }, undefined, "live").kind, "python");

  // An extension tool has no verb table: its label carries the row.
  assert.equal(describeToolRow("sum_rows", {}, undefined, "live", "Sum rows").text, "Running Sum rows");
  assert.equal(describeToolRow("sum_rows", {}, result("x", "oops", 1, { isError: true, toolName: "sum_rows" }), "failed").text, "Could not run sum_rows");

  // The approval button names the change.
  assert.equal(toolActionVerb("modify_structure", { action: "delete_rows", sheet: "Summary", position: 40, count: 13 }), "Delete");
  assert.equal(toolActionVerb("write_cells", { start_cell: "A1", values: [[1]] }), "Write");
  assert.equal(toolActionVerb("sum_rows", {}), null);
});

function pick(row: ReturnType<typeof describeToolRow>) {
  return { text: row.text, meta: row.meta, kind: row.kind, reason: row.reason };
}
