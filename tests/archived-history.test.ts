import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { ToolResultMessage, UserMessage } from "@earendil-works/pi-ai";

import {
  createArchivedMessagesMessage,
  splitArchivedMessages,
} from "../src/messages/archived-history.ts";
import { createCompactionSummaryMessage } from "../src/messages/compaction.ts";

function createUser(text: string, timestamp: number): UserMessage {
  return {
    role: "user",
    content: text,
    timestamp,
  };
}

function createToolResult(text: string, timestamp: number): ToolResultMessage {
  return {
    role: "toolResult",
    toolCallId: `tc-${timestamp}`,
    toolName: "read_range",
    content: [{ type: "text", text }],
    isError: false,
    timestamp,
  };
}

void test("splitArchivedMessages removes archive wrapper and preserves active messages", () => {
  const archive = createArchivedMessagesMessage({
    existingArchivedMessages: [],
    newlyArchivedMessages: [
      createUser("old user", 1),
      createToolResult("old tool", 2),
    ],
    timestamp: 10,
  });

  const compactionSummary = createCompactionSummaryMessage({
    summary: "summary",
    messageCountBefore: 2,
    timestamp: 11,
  });

  const activeUser = createUser("new user", 12);

  const split = splitArchivedMessages([archive, compactionSummary, activeUser]);

  assert.equal(split.archivedMessages.length, 2);
  assert.equal(split.messagesWithoutArchived.length, 2);
  assert.equal(split.messagesWithoutArchived[0]?.role, "compactionSummary");
  assert.equal(split.messagesWithoutArchived[1]?.role, "user");
});

void test("createArchivedMessagesMessage flattens nested archived payloads", () => {
  const nestedArchive = createArchivedMessagesMessage({
    existingArchivedMessages: [],
    newlyArchivedMessages: [createUser("nested", 1)],
    timestamp: 2,
  });

  const existingArchivedMessages: AgentMessage[] = [nestedArchive];

  const merged = createArchivedMessagesMessage({
    existingArchivedMessages,
    newlyArchivedMessages: [createUser("fresh", 3), createToolResult("tool", 4)],
    timestamp: 5,
  });

  assert.equal(merged.archivedMessages.length, 3);
  assert.equal(merged.archivedMessages.some((m) => m.role === "archivedMessages"), false);
  assert.equal(merged.archivedChatMessageCount, 2);
});
