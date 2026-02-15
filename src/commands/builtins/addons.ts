/**
 * Builtin command for Add-ons entrypoint UI.
 */

import type { SlashCommand } from "../types.js";

export interface AddonsCommandActions {
  openAddonsManager: () => void | Promise<void>;
}

export function createAddonsCommands(actions: AddonsCommandActions): SlashCommand[] {
  return [
    {
      name: "addons",
      description: "Open Add-ons (Tools & MCP, Skills, Extensions)",
      source: "builtin",
      execute: () => {
        void actions.openAddonsManager();
      },
    },
  ];
}
