/**
 * read_selection — Read the active selection with surrounding context.
 *
 * Mirrors the auto-context selection injection so the agent can call it explicitly.
 */

import { Type } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { readSelectionContext } from "../context/selection.js";

const schema = Type.Object({});

export function createReadSelectionTool(): AgentTool<typeof schema> {
  return {
    name: "read_selection",
    label: "Read Selection",
    description:
      "Read the user's current selection with surrounding context (±5 rows, up to 20 columns). " +
      "Use this when the user references the selected cell or nearby data.",
    parameters: schema,
    execute: async (): Promise<AgentToolResult<undefined>> => {
      try {
        const selection = await readSelectionContext();
        if (!selection) {
          return {
            content: [{ type: "text", text: "Error: could not read selection." }],
            details: undefined,
          };
        }
        return {
          content: [{ type: "text", text: selection.text }],
          details: undefined,
        };
      } catch (e: any) {
        return {
          content: [{ type: "text", text: `Error reading selection: ${e.message}` }],
          details: undefined,
        };
      }
    },
  };
}
