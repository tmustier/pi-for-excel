/**
 * get_recent_changes — Report user edits since the last message.
 */

import { Type } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import type { ChangeTracker } from "../context/change-tracker.js";

const schema = Type.Object({});

export function createGetRecentChangesTool(changeTracker: ChangeTracker): AgentTool<typeof schema> {
  return {
    name: "get_recent_changes",
    label: "Recent Changes",
    description:
      "Report user edits since the last message (based on change tracking). " +
      "Use this when the user asks what changed.",
    parameters: schema,
    execute: async (): Promise<AgentToolResult<undefined>> => {
      const changes = changeTracker.flush();
      if (!changes) {
        return {
          content: [{ type: "text", text: "No user changes since the last message." }],
          details: undefined,
        };
      }
      return {
        content: [{ type: "text", text: changes }],
        details: undefined,
      };
    },
  };
}
