/**
 * Builtin commands for unified extensions management UI.
 */

import type { ExtensionsHubTab } from "./extensions-hub-overlay.js";
import type { SlashCommand } from "../types.js";

export interface AddonsCommandActions {
  openExtensionsHub: (tab?: ExtensionsHubTab) => void | Promise<void>;
}

export function createAddonsCommands(actions: AddonsCommandActions): SlashCommand[] {
  return [
    {
      name: "extensions",
      description: "Open Extensions (connections, plugins, skills)",
      source: "builtin",
      execute: () => {
        void actions.openExtensionsHub();
      },
    },
  ];
}
