import type { UserMessage } from "@earendil-works/pi-ai";

export const AUTO_CONTEXT_PREFIX = "[Auto-context]";

export type AutoContextKind = "workbook" | "files" | "changes";

export interface AutoContextMessage {
  role: "autoContext";
  kind: AutoContextKind;
  key?: string;
  workbookId?: string | null;
  text: string;
  timestamp: number;
}

declare module "@earendil-works/pi-agent-core" {
  interface CustomAgentMessages {
    autoContext: AutoContextMessage;
  }
}

export function autoContextToUserMessage(message: AutoContextMessage): UserMessage {
  return {
    role: "user",
    content: [{ type: "text", text: `${AUTO_CONTEXT_PREFIX} ${message.text}` }],
    timestamp: message.timestamp,
  };
}
