/**
 * Builtin command for plugin management UI.
 */

import type { ExtensionsHubTab } from "./extensions-hub-overlay.js";
import type { SlashCommand } from "../types.js";

export interface ExtensionsCommandActions {
  openExtensionsHub: (tab?: ExtensionsHubTab) => void | Promise<void>;
}

export function createExtensionsCommands(actions: ExtensionsCommandActions): SlashCommand[] {
  return [
    {
      name: "plugins",
      description: "Manage installed plugins (alias for /extensions plugins)",
      source: "builtin",
      execute: () => {
        void actions.openExtensionsHub("plugins");
      },
    },
  ];
}
