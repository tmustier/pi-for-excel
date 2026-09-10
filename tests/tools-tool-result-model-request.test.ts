import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage } from "@earendil-works/pi-agent-core";
import {
  createModels,
  fauxAssistantMessage,
  fauxProvider,
  type Context,
} from "@earendil-works/pi-ai";

import { createConvertToLlm } from "../src/messages/convert-to-llm.ts";

type ToolResultMessage = Extract<AgentMessage, { role: "toolResult" }>;

function toolResult(
  id: string,
  text: string,
  options: { image?: boolean; legacy?: boolean } = {},
): ToolResultMessage {
  const message: ToolResultMessage = {
    role: "toolResult",
    toolCallId: id,
    toolName: "read_range",
    content: [
      { type: "text", text },
      ...(options.image ? [{ type: "image" as const, data: "AA==", mimeType: "image/png" }] : []),
    ],
    isError: false,
    timestamp: Number(id),
  };
  if (options.legacy) Reflect.set(message, "content", text);
  return message;
}

async function sendToModel(messages: AgentMessage[]): Promise<{ request: Context; history: AgentMessage[] }> {
  const faux = fauxProvider({ models: [{ id: "tool-result-request", contextWindow: 32_000, maxTokens: 4_096 }] });
  const models = createModels();
  models.setProvider(faux.provider);
  faux.setResponses([fauxAssistantMessage("done")]);
  const requests: Context[] = [];
  const agent = new Agent({
    initialState: { model: faux.getModel(), messages, tools: [] },
    convertToLlm: createConvertToLlm({ getContextWindow: () => 32_000 }),
    streamFn: (model, context, options) => {
      requests.push(structuredClone(context));
      return models.streamSimple(model, context, options);
    },
  });

  await agent.prompt("continue");
  const request = requests[0];
  if (!request) throw new Error("Expected an outbound model request.");
  return { request, history: agent.state.messages };
}

function requestToolResult(request: Context, id: string): ToolResultMessage {
  const message = request.messages.find(
    (candidate) => candidate.role === "toolResult" && candidate.toolCallId === id,
  );
  if (!message || message.role !== "toolResult") {
    throw new Error(`Expected tool result ${id} in outbound request.`);
  }
  return message;
}

function textOf(message: ToolResultMessage): string {
  const block = message.content.find((candidate) => candidate.type === "text");
  if (!block || block.type !== "text") throw new Error("Expected text tool result.");
  return block.text;
}

void test("model request compacts an older large tool result but preserves recent output and chat history", async () => {
  const large = "x".repeat(1_500);
  const older = toolResult("1", large);
  const { request, history } = await sendToModel([
    older,
    toolResult("2", large),
    toolResult("3", large),
  ]);

  assert.match(textOf(requestToolResult(request, "1")), /^\[Compacted tool result\] read_range/u);
  assert.equal(textOf(requestToolResult(request, "3")), large);
  const historyOlder = history.find(
    (message) => message.role === "toolResult" && message.toolCallId === "1",
  );
  if (!historyOlder || historyOlder.role !== "toolResult") {
    throw new Error("Expected original tool result in chat history.");
  }
  assert.equal(textOf(historyOlder), large);
});

void test("model request keeps an older small tool result intact", async () => {
  const { request } = await sendToModel([
    toolResult("1", "short"),
    toolResult("2", "recent"),
    toolResult("3", "recent"),
  ]);

  assert.equal(textOf(requestToolResult(request, "1")), "short");
});

void test("model request compacts an older legacy string tool result", async () => {
  const { request } = await sendToModel([
    toolResult("1", "legacy".repeat(350), { legacy: true }),
    toolResult("2", "recent"),
    toolResult("3", "recent"),
  ]);

  assert.match(textOf(requestToolResult(request, "1")), /^\[Compacted tool result\]/u);
});

void test("model request replaces an older image tool result with a text-only summary", async () => {
  const { request } = await sendToModel([
    toolResult("1", "ok", { image: true }),
    toolResult("2", "recent"),
    toolResult("3", "recent"),
  ]);

  const older = requestToolResult(request, "1");
  assert.equal(older.content.length, 1);
  assert.match(textOf(older), /1 image block/u);
});
