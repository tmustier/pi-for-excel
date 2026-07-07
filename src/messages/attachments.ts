/**
 * First-party attachment + custom message role types.
 *
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono) as part of the UI ownership migration
 * (docs/ui-ownership.md).
 *
 * Pi for Excel does not currently create attachment or artifact messages
 * (document parsing was never bundled), but the roles remain part of the
 * message union so that:
 * - restored/shared sessions containing them keep loading, and
 * - convertToLlm() keeps handling them exactly as before.
 */

import type { ImageContent, TextContent } from "@earendil-works/pi-ai";
import type { AgentMessage } from "@earendil-works/pi-agent-core";

export interface Attachment {
  id: string;
  type: "image" | "document";
  fileName: string;
  mimeType: string;
  size: number;
  /** Base64 file content. */
  content: string;
  extractedText?: string;
  /** Base64 preview image (PNG unless `type` is "image"). */
  preview?: string;
}

export type UserMessageWithAttachments = {
  role: "user-with-attachments";
  content: string | (TextContent | ImageContent)[];
  timestamp: number;
  attachments?: Attachment[];
};

export interface ArtifactMessage {
  role: "artifact";
  action: "create" | "update" | "delete";
  filename: string;
  content?: string;
  title?: string;
  timestamp: string;
}

declare module "@earendil-works/pi-agent-core" {
  interface CustomAgentMessages {
    "user-with-attachments": UserMessageWithAttachments;
    artifact: ArtifactMessage;
  }
}

export function isUserMessageWithAttachments(
  message: AgentMessage,
): message is UserMessageWithAttachments {
  return message.role === "user-with-attachments";
}

export function isArtifactMessage(message: AgentMessage): message is ArtifactMessage {
  return message.role === "artifact";
}

/**
 * Convert attachments to content blocks for the LLM.
 * - Images become ImageContent blocks.
 * - Documents with extractedText become TextContent blocks with a filename header.
 */
export function convertAttachments(attachments: Attachment[]): (TextContent | ImageContent)[] {
  const content: (TextContent | ImageContent)[] = [];

  for (const attachment of attachments) {
    if (attachment.type === "image") {
      content.push({
        type: "image",
        data: attachment.content,
        mimeType: attachment.mimeType,
      });
    } else if (attachment.type === "document" && attachment.extractedText) {
      content.push({
        type: "text",
        text: `\n\n[Document: ${attachment.fileName}]\n${attachment.extractedText}`,
      });
    }
  }

  return content;
}
