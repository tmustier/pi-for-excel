import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentMessage } from "@earendil-works/pi-agent-core";

import { shapeToolResultsForLlm } from "../src/messages/tool-result-shaping.ts";

type ToolResultMessage = Extract<AgentMessage, { role: "toolResult" }>;
type UserMessage = Extract<AgentMessage, { role: "user" }>;
type TextBlock = Extract<ToolResultMessage["content"][number], { type: "text" }>;

function isToolResultMessage(message: AgentMessage | undefined): message is ToolResultMessage {
  return message?.role === "toolResult";
}

function getToolResultMessage(messages: readonly AgentMessage[], index: number): ToolResultMessage {
  const message = messages[index];
  if (!isToolResultMessage(message)) {
    throw new Error(`Expected tool result message at index ${index}.`);
  }

  return message;
}

function isTextBlock(block: ToolResultMessage["content"][number]): block is TextBlock {
  return block.type === "text";
}

function getTextBlock(message: ToolResultMessage): TextBlock {
  const block = message.content.find(isTextBlock);
  if (!block) {
    throw new Error("Expected text block in tool result message.");
  }

  return block;
}

function createUser(text: string, timestamp: number): UserMessage {
  return {
    role: "user",
    content: text,
    timestamp,
  };
}

function createToolResult(args: {
  toolCallId: string;
  toolName: string;
  text: string;
  timestamp: number;
  includeImage?: boolean;
  isError?: boolean;
}): ToolResultMessage {
  const content: ToolResultMessage["content"] = [
    {
      type: "text",
      text: args.text,
    },
  ];

  if (args.includeImage) {
    content.push({
      type: "image",
      data: "AA==",
      mimeType: "image/png",
    });
  }

  return {
    role: "toolResult",
    toolCallId: args.toolCallId,
    toolName: args.toolName,
    content,
    isError: args.isError ?? false,
    timestamp: args.timestamp,
  };
}

function forceLegacyStringContent(message: ToolResultMessage, text: string): void {
  Reflect.set(message, "content", text);
}

void test("compacts older large tool results but preserves recent ones", () => {
  const long = "x".repeat(1500);
  const older = createToolResult({
    toolCallId: "1",
    toolName: "read_range",
    text: long,
    timestamp: 1,
  });

  const recent = createToolResult({
    toolCallId: "2",
    toolName: "read_range",
    text: long,
    timestamp: 2,
  });

  const messages: AgentMessage[] = [createUser("analyze", 0), older, recent];
  const shaped = shapeToolResultsForLlm(messages, {
    recentToolResultsToKeep: 1,
    maxCharsBeforeCompaction: 1200,
    previewChars: 40,
  });

  const shapedOlder = getToolResultMessage(shaped, 1);
  const olderText = getTextBlock(shapedOlder);
  assert.ok(olderText.text.startsWith("[Compacted tool result] read_range"));

  const shapedRecent = getToolResultMessage(shaped, 2);
  const recentText = getTextBlock(shapedRecent);
  assert.equal(recentText.text, long);

  const originalOlderText = getTextBlock(older);
  assert.equal(originalOlderText.text, long);
});

void test("keeps older small tool results intact", () => {
  const older = createToolResult({
    toolCallId: "1",
    toolName: "read_range",
    text: "short",
    timestamp: 1,
  });

  const recent = createToolResult({
    toolCallId: "2",
    toolName: "read_range",
    text: "short",
    timestamp: 2,
  });

  const shaped = shapeToolResultsForLlm([older, recent], {
    recentToolResultsToKeep: 1,
    maxCharsBeforeCompaction: 1200,
    previewChars: 40,
  });

  const shapedOlder = getToolResultMessage(shaped, 0);
  const olderText = getTextBlock(shapedOlder);
  assert.equal(olderText.text, "short");
});

void test("compacts older legacy string tool-result payloads", () => {
  const long = "legacy".repeat(350);

  const older = createToolResult({
    toolCallId: "1",
    toolName: "read_range",
    text: long,
    timestamp: 1,
  });
  forceLegacyStringContent(older, long);

  const recent = createToolResult({
    toolCallId: "2",
    toolName: "read_range",
    text: "ok",
    timestamp: 2,
  });

  const shaped = shapeToolResultsForLlm([older, recent], {
    recentToolResultsToKeep: 1,
    maxCharsBeforeCompaction: 1200,
    previewChars: 40,
  });

  const shapedOlder = getToolResultMessage(shaped, 0);
  const olderText = getTextBlock(shapedOlder);
  assert.match(olderText.text, /^\[Compacted tool result\]/);
});

void test("compacts older image tool results even with small text", () => {
  const older = createToolResult({
    toolCallId: "1",
    toolName: "read_range",
    text: "ok",
    timestamp: 1,
    includeImage: true,
  });

  const recent = createToolResult({
    toolCallId: "2",
    toolName: "read_range",
    text: "ok",
    timestamp: 2,
  });

  const shaped = shapeToolResultsForLlm([older, recent], {
    recentToolResultsToKeep: 1,
    maxCharsBeforeCompaction: 1200,
    previewChars: 40,
  });

  const shapedOlder = getToolResultMessage(shaped, 0);
  assert.equal(shapedOlder.content.length, 1);
  const onlyBlock = getTextBlock(shapedOlder);
  assert.match(onlyBlock.text, /image block/);
});
