/**
 * Builtin command for plugin management UI.
 */

import type { AddonsSection } from "./addons-overlay.js";
import type { SlashCommand } from "../types.js";

export interface ExtensionsCommandActions {
  openAddonsManager: (section?: AddonsSection) => void | Promise<void>;
}

export function createExtensionsCommands(actions: ExtensionsCommandActions): SlashCommand[] {
  return [
    {
      name: "plugins",
      description: "Manage installed plugins (alias for /extensions plugins)",
      source: "builtin",
      execute: () => {
        void actions.openAddonsManager("plugins");
      },
    },
  ];
}
