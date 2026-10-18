import assert from "node:assert/strict";
import { test } from "node:test";

import { type AgentMessage, type AgentTool, type AgentToolResult } from "@earendil-works/pi-agent-core";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { MemoryStorage } from "@earendil-works/pi-durable";
import type { Context } from "@earendil-works/pi-ai";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import { Type } from "typebox";

import { openDurableRuntime } from "../src/durable/runtime.ts";
import { DurableSessions } from "../src/durable/sessions.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import { applyToolOutputTruncation } from "../src/tools/output-truncation.ts";
import { getToolOutputTruncationDetails } from "../src/tools/tool-details.ts";

const parameters = Type.Object({});
type Result = AgentToolResult<unknown>;
type ToolResultMessage = Extract<AgentMessage, { role: "toolResult" }>;

function tool(name: string, result: Result, update?: Result): AgentTool {
  return {
    name,
    label: name,
    description: name,
    parameters,
    execute: (_id, _params, _signal, onUpdate) => {
      if (update && onUpdate) onUpdate(update);
      return Promise.resolve(result);
    },
  };
}

function text(result: Result): string {
  const block = result.content.find((candidate) => candidate.type === "text");
  if (!block || block.type !== "text") throw new Error("Expected text output");
  return block.text;
}

async function modelVisibleResult(name: string, result: Result): Promise<ToolResultMessage> {
  const faux = fauxProvider({ models: [{ id: "truncation", contextWindow: 128_000, maxTokens: 4_096 }] });
  const models = createModels();
  faux.setResponses([fauxAssistantMessage("done")]);
  let request: Context | undefined;
  const message: ToolResultMessage = {
    role: "toolResult",
    toolCallId: "call-large",
    toolName: name,
    content: result.content,
    details: result.details,
    isError: false,
    timestamp: 1,
  };
  models.setProvider({ ...faux.provider, streamSimple: (model, transcript, options) => {
    request = structuredClone(transcript);
    return faux.provider.streamSimple(model, transcript, options);
  },
  });
  const runtime = await openDurableRuntime({ storage: new MemoryStorage(), models,
    workbookId: "truncation", coordinator: createWorkbookCoordinator(), autoCompaction: false,
    mutationHost: { prepare: () => { throw new Error("Unexpected mutation"); }, observe: () => { throw new Error("Unexpected mutation"); } },
  });
  try {
    const conversation = await new DurableSessions(runtime.harness, models, faux.getModel()).create(faux.getModel(), "off");
    await runtime.install(conversation, { tools: [], systemPrompt: "" }, (history) => Promise.resolve(history));
    await runtime.harness.commit(async (tx) => {
      await tx.appendEntry(conversation.id, { kind: "test.history", model: [
        { ...fauxAssistantMessage(""), stopReason: "toolUse", content: [{ type: "toolCall", id: message.toolCallId, name, arguments: {} }] }, message,
      ] });
    }, context);
    const submission = await conversation.submit({ type: "input", content: "continue" }, context);
    assert.equal((await submission.wait(context)).status, "done");
    const visible = request?.messages.find(
      (candidate) => candidate.role === "toolResult" && candidate.toolCallId === "call-large",
    );
    if (!visible || visible.role !== "toolResult") throw new Error("Expected model-visible tool result");
    return visible;
  } finally { await runtime.harness.close(context); }
}

function visibleText(message: ToolResultMessage): string {
  const block = message.content.find((candidate) => candidate.type === "text");
  if (!block || block.type !== "text") throw new Error("Expected model-visible text");
  return block.text;
}

void test("the model receives the head of a large workbook result with truncation metadata", async () => {
  const [wrapped] = applyToolOutputTruncation([
    tool("read_range", { content: [{ type: "text", text: "first\nsecond\nthird" }], details: {} }),
  ], { limits: { maxLines: 2, maxBytes: 10_000 } });
  if (!wrapped) throw new Error("Expected wrapped tool");

  const visible = await modelVisibleResult("read_range", await wrapped.execute("call", {}));
  assert.equal(visibleText(visible), "first\nsecond\n\n[Output truncated: showing first 2 of 3 lines; limits: 2 lines / 9.8KB]");
  assert.deepEqual(getToolOutputTruncationDetails(visible.details), {
    version: 1,
    truncated: true,
    strategy: "head",
    truncatedBy: "lines",
    totalLines: 3,
    totalBytes: 18,
    outputLines: 2,
    outputBytes: 12,
    maxLines: 2,
    maxBytes: 10_000,
  });
});

void test("the model receives the tail of large log-style output", async () => {
  const [wrapped] = applyToolOutputTruncation([
    tool("python_run", { content: [{ type: "text", text: "first\nsecond\nthird" }], details: {} }),
  ], { limits: { maxLines: 2, maxBytes: 10_000 } });
  if (!wrapped) throw new Error("Expected wrapped tool");

  const visible = await modelVisibleResult("python_run", await wrapped.execute("call", {}));
  assert.equal(visibleText(visible), "second\nthird\n\n[Output truncated: showing last 2 of 3 lines; limits: 2 lines / 9.8KB]");
});

void test("the model receives the saved full-output path when truncation persistence succeeds", async () => {
  const [wrapped] = applyToolOutputTruncation([
    tool("search_workbook", { content: [{ type: "text", text: "abcdef" }], details: {} }),
  ], {
    limits: { maxLines: 10, maxBytes: 3 },
    saveTruncatedOutput: () => Promise.resolve(".tool-output/complete.txt"),
  });
  if (!wrapped) throw new Error("Expected wrapped tool");

  const visible = await modelVisibleResult("search_workbook", await wrapped.execute("call", {}));
  assert.match(visibleText(visible), /full output saved to Files workspace: \.tool-output\/complete\.txt/u);
  assert.equal(getToolOutputTruncationDetails(visible.details)?.fullOutputWorkspacePath, ".tool-output/complete.txt");
});

void test("the model still receives image content beside truncated text", async () => {
  const [wrapped] = applyToolOutputTruncation([
    tool("charts", {
      content: [
        { type: "text", text: "first\nsecond\nthird" },
        { type: "image", data: "AA==", mimeType: "image/png" },
      ],
      details: {},
    }),
  ], { limits: { maxLines: 2, maxBytes: 10_000 } });
  if (!wrapped) throw new Error("Expected wrapped tool");

  const visible = await modelVisibleResult("charts", await wrapped.execute("call", {}));
  assert.deepEqual(visible.content[1], { type: "image", data: "AA==", mimeType: "image/png" });
});

void test("a model-facing tool execution truncates its streaming update", async () => {
  const [wrapped] = applyToolOutputTruncation([
    tool(
      "python_run",
      { content: [{ type: "text", text: "done" }], details: {} },
      { content: [{ type: "text", text: "first\nsecond\nthird" }], details: {} },
    ),
  ], { limits: { maxLines: 2, maxBytes: 10_000 } });
  if (!wrapped) throw new Error("Expected wrapped tool");
  let update: Result | undefined;

  await wrapped.execute("call", {}, undefined, (partial) => { update = partial; });
  if (!update) throw new Error("Expected streaming update");
  assert.equal(text(update), "second\nthird\n\n[Output truncated: showing last 2 of 3 lines; limits: 2 lines / 9.8KB]");
});
