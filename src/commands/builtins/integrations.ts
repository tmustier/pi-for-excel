/**
 * Builtin command for integration management UI.
 */

import { INTEGRATIONS_COMMAND_NAME, INTEGRATIONS_MANAGER_LABEL_LOWER } from "../../integrations/naming.js";
import type { SlashCommand } from "../types.js";

export interface IntegrationsCommandActions {
  openIntegrationsManager: () => void | Promise<void>;
}

export function createIntegrationsCommands(actions: IntegrationsCommandActions): SlashCommand[] {
  return [
    {
      name: INTEGRATIONS_COMMAND_NAME,
      description: `Manage ${INTEGRATIONS_MANAGER_LABEL_LOWER} (web search, page fetch, MCP)`,
      source: "builtin",
      execute: () => {
        void actions.openIntegrationsManager();
      },
    },
  ];
}
