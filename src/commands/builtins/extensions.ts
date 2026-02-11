/**
 * Builtin command for extension management UI.
 */

import type { SlashCommand } from "../types.js";

export interface ExtensionsCommandActions {
  openExtensionsManager: () => void | Promise<void>;
}

export function createExtensionsCommands(actions: ExtensionsCommandActions): SlashCommand[] {
  return [
    {
      name: "extensions",
      description: "Manage installed extensions",
      source: "builtin",
      execute: () => {
        void actions.openExtensionsManager();
      },
    },
  ];
}
