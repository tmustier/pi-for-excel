/**
 * Builtin command for Tools & MCP management UI.
 */

import {
  INTEGRATIONS_MANAGER_LABEL_LOWER,
  TOOLS_COMMAND_NAME,
} from "../../integrations/naming.js";
import type { ExtensionsHubTab } from "./extensions-hub-overlay.js";
import type { SlashCommand } from "../types.js";

export interface ToolsCommandActions {
  openExtensionsHub: (tab?: ExtensionsHubTab) => void | Promise<void>;
}

export function createToolsCommands(actions: ToolsCommandActions): SlashCommand[] {
  return [
    {
      name: TOOLS_COMMAND_NAME,
      description: `Manage ${INTEGRATIONS_MANAGER_LABEL_LOWER} (alias for /extensions connections)`,
      source: "builtin",
      execute: () => {
        void actions.openExtensionsHub("connections");
      },
    },
  ];
}
