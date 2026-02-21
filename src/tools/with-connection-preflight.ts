import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";

import {
  ConnectionManager,
  looksLikeConnectionAuthFailure,
} from "../connections/manager.js";
import type {
  ConnectionSnapshot,
  ConnectionToolErrorCode,
  ConnectionToolErrorDetails,
} from "../connections/types.js";
import { getToolRequiredConnectionIds } from "./connection-requirements.js";

function normalizeErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }

  return String(error);
}

function mapStatusToErrorCode(status: ConnectionSnapshot["status"]): ConnectionToolErrorCode {
  if (status === "missing") return "missing_connection";
  if (status === "invalid") return "invalid_connection";
  if (status === "error") return "connection_auth_failed";
  return "invalid_connection";
}

function buildErrorMessage(details: ConnectionToolErrorDetails): string {
  if (details.errorCode === "missing_connection") {
    return `Connection \"${details.connectionTitle}\" is not configured. ${details.setupHint}.`;
  }

  if (details.errorCode === "invalid_connection") {
    const reasonSuffix = details.reason ? ` (${details.reason})` : "";
    return `Connection \"${details.connectionTitle}\" is invalid${reasonSuffix}. ${details.setupHint}.`;
  }

  const reasonSuffix = details.reason ? ` (${details.reason})` : "";
  return `Connection \"${details.connectionTitle}\" failed authentication${reasonSuffix}. ${details.setupHint}.`;
}

function buildConnectionErrorResult(args: {
  snapshot: ConnectionSnapshot;
  errorCode: ConnectionToolErrorCode;
  reason?: string;
}): AgentToolResult<ConnectionToolErrorDetails> {
  const details: ConnectionToolErrorDetails = {
    kind: "connection_error",
    ok: false,
    errorCode: args.errorCode,
    connectionId: args.snapshot.connectionId,
    connectionTitle: args.snapshot.title,
    status: args.snapshot.status,
    setupHint: args.snapshot.setupHint,
    reason: args.reason,
  };

  const message = buildErrorMessage(details);
  return {
    content: [{ type: "text", text: message }],
    details,
  };
}

function buildUnregisteredConnectionResult(connectionId: string): AgentToolResult<ConnectionToolErrorDetails> {
  const details: ConnectionToolErrorDetails = {
    kind: "connection_error",
    ok: false,
    errorCode: "invalid_connection",
    connectionId,
    connectionTitle: connectionId,
    status: "invalid",
    setupHint: "Reload the extension, then open /tools → Connections.",
    reason: "Connection requirement is not registered in this session.",
  };

  return {
    content: [{ type: "text", text: buildErrorMessage(details) }],
    details,
  };
}

function wrapTool(tool: AgentTool, connectionManager: ConnectionManager): AgentTool {
  const requiredConnectionIds = getToolRequiredConnectionIds(tool);
  if (requiredConnectionIds.length === 0) {
    return tool;
  }

  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      const snapshots: ConnectionSnapshot[] = [];

      for (const connectionId of requiredConnectionIds) {
        const snapshot = await connectionManager.getSnapshot(connectionId);
        if (!snapshot) {
          return buildUnregisteredConnectionResult(connectionId);
        }

        snapshots.push(snapshot);

        if (snapshot.status !== "connected") {
          const errorCode = mapStatusToErrorCode(snapshot.status);
          return buildConnectionErrorResult({
            snapshot,
            errorCode,
            reason: snapshot.lastError,
          });
        }
      }

      try {
        return await tool.execute(toolCallId, params, signal, onUpdate);
      } catch (error: unknown) {
        const errorMessage = normalizeErrorMessage(error);
        const primarySnapshot = snapshots[0];

        if (primarySnapshot && looksLikeConnectionAuthFailure(errorMessage)) {
          try {
            await connectionManager.markRuntimeAuthFailure(primarySnapshot.connectionId, {
              message: errorMessage,
            });
          } catch {
            // best-effort status update only
          }

          const refreshedSnapshot = await connectionManager.getSnapshot(primarySnapshot.connectionId);
          const snapshotForResponse: ConnectionSnapshot = refreshedSnapshot ?? {
            ...primarySnapshot,
            status: "error",
          };

          return buildConnectionErrorResult({
            snapshot: snapshotForResponse,
            errorCode: "connection_auth_failed",
            reason: snapshotForResponse.lastError ?? errorMessage,
          });
        }

        throw error;
      }
    },
  };
}

export function withConnectionPreflight(
  tools: AgentTool[],
  args: {
    connectionManager: ConnectionManager;
  },
): AgentTool[] {
  return tools.map((tool) => wrapTool(tool, args.connectionManager));
}
