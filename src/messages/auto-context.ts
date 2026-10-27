/**
 * Auto-context notes stored in the transcript.
 *
 * Pi tells the agent about the workbook without being asked. What is a
 * snapshot that stays true until something changes (workbook structure, the
 * Files workspace summary) or an event that happened (cells the user edited)
 * is stored here, in the transcript, at the point Pi observed it. Stored notes
 * become part of the cached prompt prefix, stay visible for the rest of the
 * turn and in later turns, and are replaced only when their subject changes.
 *
 * The current selection is different: it is state, re-read on every request
 * and sent as the last message, never stored (see
 * `src/taskpane/context-injection.ts`).
 *
 * The role is UI-invisible (no renderer) and is turned into a user message for
 * the model by `convertToLlm`. Agent-facing text: never route it through `t()`.
 */

import type { UserMessage } from "@earendil-works/pi-ai";

export const AUTO_CONTEXT_PREFIX = "[Auto-context]";

export type AutoContextKind = "workbook" | "files" | "changes";

export interface AutoContextMessage {
  role: "autoContext";
  kind: AutoContextKind;
  /**
   * What a snapshot note describes, so a later request can tell whether it is
   * still current: the workbook id plus a hash of the blueprint, or the Files
   * workspace signature. Edit notes are events and carry no key.
   */
  key?: string;
  /** Workbook the note was taken from (workbook notes). */
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
