import type { AgentMessage } from "@earendil-works/pi-agent-core";

export interface ArchivedMessagesMessage {
  role: "archivedMessages";
  archivedMessages: AgentMessage[];
  archivedChatMessageCount: number;
  timestamp: number;
}

declare module "@earendil-works/pi-agent-core" {
  interface CustomAgentMessages {
    archivedMessages: ArchivedMessagesMessage;
  }
}

function isArchivedMessagesMessage(message: AgentMessage): message is ArchivedMessagesMessage {
  return message.role === "archivedMessages";
}

function countChatMessages(messages: readonly AgentMessage[]): number {
  let count = 0;

  for (const message of messages) {
    if (
      message.role === "user" ||
      message.role === "assistant" ||
      message.role === "user-with-attachments"
    ) {
      count += 1;
    }
  }

  return count;
}

function flattenArchivedPayload(messages: readonly AgentMessage[]): AgentMessage[] {
  const flattened: AgentMessage[] = [];

  for (const message of messages) {
    if (message.role === "artifact") {
      // Artifacts are UI-only and already excluded from MessageList/LLM context.
      continue;
    }

    if (isArchivedMessagesMessage(message)) {
      flattened.push(...flattenArchivedPayload(message.archivedMessages));
      continue;
    }

    flattened.push(message);
  }

  return flattened;
}

export function splitArchivedMessages(messages: readonly AgentMessage[]): {
  archivedMessages: AgentMessage[];
  messagesWithoutArchived: AgentMessage[];
} {
  const archivedMessages: AgentMessage[] = [];
  const messagesWithoutArchived: AgentMessage[] = [];

  for (const message of messages) {
    if (isArchivedMessagesMessage(message)) {
      archivedMessages.push(...flattenArchivedPayload(message.archivedMessages));
      continue;
    }

    messagesWithoutArchived.push(message);
  }

  return { archivedMessages, messagesWithoutArchived };
}

export function createArchivedMessagesMessage(args: {
  existingArchivedMessages: AgentMessage[];
  newlyArchivedMessages: AgentMessage[];
  timestamp: number;
}): ArchivedMessagesMessage {
  const archivedMessages = flattenArchivedPayload([
    ...args.existingArchivedMessages,
    ...args.newlyArchivedMessages,
  ]);

  return {
    role: "archivedMessages",
    archivedMessages,
    archivedChatMessageCount: countChatMessages(archivedMessages),
    timestamp: args.timestamp,
  };
}
