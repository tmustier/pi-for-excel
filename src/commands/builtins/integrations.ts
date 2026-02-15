/**
 * Builtin command for integration management UI.
 */

import {
  INTEGRATIONS_COMMAND_NAME,
  INTEGRATIONS_MANAGER_LABEL_LOWER,
  TOOLS_COMMAND_NAME,
} from "../../integrations/naming.js";
import type { AddonsSection } from "./addons-overlay.js";
import type { SlashCommand } from "../types.js";

export interface IntegrationsCommandActions {
  openAddonsManager: (section?: AddonsSection) => void | Promise<void>;
}

export function createIntegrationsCommands(actions: IntegrationsCommandActions): SlashCommand[] {
  const openConnections = () => {
    void actions.openAddonsManager("connections");
  };

  return [
    {
      name: TOOLS_COMMAND_NAME,
      description: `Manage ${INTEGRATIONS_MANAGER_LABEL_LOWER} (alias for /addons connections)`,
      source: "builtin",
      execute: openConnections,
    },
    {
      name: INTEGRATIONS_COMMAND_NAME,
      description: `Alias for /${TOOLS_COMMAND_NAME}`,
      source: "builtin",
      execute: openConnections,
    },
  ];
}
