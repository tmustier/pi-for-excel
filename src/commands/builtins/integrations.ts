/**
 * Builtin command for integration management UI.
 */

import {
  INTEGRATIONS_COMMAND_NAME,
  INTEGRATIONS_MANAGER_LABEL_LOWER,
  TOOLS_COMMAND_NAME,
} from "../../integrations/naming.js";
import type { SlashCommand } from "../types.js";

export interface IntegrationsCommandActions {
  openIntegrationsManager: () => void | Promise<void>;
}

export function createIntegrationsCommands(actions: IntegrationsCommandActions): SlashCommand[] {
  const openManager = () => {
    void actions.openIntegrationsManager();
  };

  return [
    {
      name: TOOLS_COMMAND_NAME,
      description: `Manage ${INTEGRATIONS_MANAGER_LABEL_LOWER} (web search, page fetch, MCP)`,
      source: "builtin",
      execute: openManager,
    },
    {
      name: INTEGRATIONS_COMMAND_NAME,
      description: `Alias for /${TOOLS_COMMAND_NAME}`,
      source: "builtin",
      execute: openManager,
    },
  ];
}
