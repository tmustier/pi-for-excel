import {
  getFilesWorkspace,
  type FilesWorkspaceAuditContext,
} from "../files/workspace.js";
import type { ToolOutputTruncationStoreArgs } from "./output-truncation.js";

export const DEFAULT_TOOL_OUTPUT_SAVE_MAX_BYTES = 512 * 1024;

const TOOL_OUTPUT_AUDIT_CONTEXT: FilesWorkspaceAuditContext = {
  actor: "system",
  source: "tool-output-truncation",
};

function sanitizePathToken(token: string): string {
  const normalized = token.trim().toLowerCase().replaceAll(/[^a-z0-9_-]+/g, "-");
  return normalized.length > 0 ? normalized : "tool";
}

function buildOutputWorkspacePath(toolName: string, toolCallId: string): string {
  const safeToolName = sanitizePathToken(toolName);
  const safeToolCallId = sanitizePathToken(toolCallId).slice(0, 32);
  const stamp = new Date().toISOString().replaceAll(/[:.]/g, "-");
  return `.tool-output/${stamp}-${safeToolName}-${safeToolCallId}.txt`;
}

export async function saveTruncatedToolOutputToWorkspace(
  args: ToolOutputTruncationStoreArgs,
): Promise<string | undefined> {
  if (args.fullText.length === 0) {
    return undefined;
  }

  if (args.truncation.totalBytes > DEFAULT_TOOL_OUTPUT_SAVE_MAX_BYTES) {
    return undefined;
  }

  const workspacePath = buildOutputWorkspacePath(args.toolName, args.toolCallId);
  const workspace = getFilesWorkspace();

  const headerLines = [
    `Tool: ${args.toolName}`,
    `Tool call ID: ${args.toolCallId}`,
    `Saved at: ${new Date().toISOString()}`,
    `Strategy: ${args.truncation.strategy}`,
    `Total payload: ${args.truncation.totalLines.toLocaleString()} lines, ${args.truncation.totalBytes.toLocaleString()} bytes`,
    "",
  ];

  await workspace.writeTextFile(
    workspacePath,
    `${headerLines.join("\n")}${args.fullText}`,
    "text/plain",
    { audit: TOOL_OUTPUT_AUDIT_CONTEXT },
  );

  return workspacePath;
}
