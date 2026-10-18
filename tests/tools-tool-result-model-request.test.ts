import assert from "node:assert/strict";
import { test } from "node:test";

import { type AgentMessage } from "@earendil-works/pi-agent-core";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { MemoryStorage } from "@earendil-works/pi-durable";
import {
  type Context,
} from "@earendil-works/pi-ai";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";

import { openDurableRuntime } from "../src/durable/runtime.ts";
import { DurableSessions } from "../src/durable/sessions.ts";
import { ConversationPresenter } from "../src/durable/conversation-view.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";

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
  const faux = fauxProvider({ tokensPerSecond: Infinity, models: [{ id: "tool-result-request", contextWindow: 32_000, maxTokens: 4_096 }] });
  const models = createModels();
  faux.setResponses([fauxAssistantMessage("done")]);
  const requests: Context[] = [];
  models.setProvider({ ...faux.provider,
    streamSimple: (model, transcript, options) => {
      requests.push(structuredClone(transcript));
      return faux.provider.streamSimple(model, transcript, options);
    },
  });
  const runtime = await openDurableRuntime({ storage: new MemoryStorage(), models,
    workbookId: "tool-result", coordinator: createWorkbookCoordinator(), autoCompaction: false,
    mutationHost: { prepare: () => { throw new Error("Unexpected mutation"); }, observe: () => { throw new Error("Unexpected mutation"); } },
  });
  try {
    const sessions = new DurableSessions(runtime.harness, models, faux.getModel());
    const conversation = await sessions.create(faux.getModel(), "off");
    await runtime.install(conversation, { tools: [], systemPrompt: "" }, (history) => Promise.resolve(history));
    await runtime.harness.commit(async (tx) => {
      for (const message of messages) {
        if (message.role === "toolResult") await tx.appendEntry(conversation.id, {
          kind: "test.calls", model: [{ ...fauxAssistantMessage(""),
            content: [{ type: "toolCall", id: message.toolCallId, name: message.toolName, arguments: {} }], stopReason: "toolUse" }],
        });
        await tx.appendEntry(conversation.id, { kind: "test.history", model: [message] });
      }
    }, context);
    const submission = await conversation.submit({ type: "input", content: "continue" }, context);
    assert.equal((await submission.wait(context)).status, "done");
    const request = requests[0];
    if (!request) throw new Error("Expected an outbound model request.");
    const presenter = await ConversationPresenter.open(runtime.harness, conversation, models, faux.getModel());
    try { return { request, history: presenter.current.messages }; }
    finally { await presenter.dispose(); }
  } finally { await runtime.harness.close(context); }
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
