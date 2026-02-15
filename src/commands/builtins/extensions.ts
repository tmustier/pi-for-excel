/**
 * Builtin command for extension management UI.
 */

import type { AddonsSection } from "./addons-overlay.js";
import type { SlashCommand } from "../types.js";

export interface ExtensionsCommandActions {
  openAddonsManager: (section?: AddonsSection) => void | Promise<void>;
}

export function createExtensionsCommands(actions: ExtensionsCommandActions): SlashCommand[] {
  return [
    {
      name: "extensions",
      description: "Manage installed extensions (alias for /addons extensions)",
      source: "builtin",
      execute: () => {
        void actions.openAddonsManager("extensions");
      },
    },
  ];
}
