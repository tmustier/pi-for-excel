/**
 * Builtin commands for unified extensions management UI.
 */

import type { AddonsSection } from "./addons-overlay.js";
import type { SlashCommand } from "../types.js";

export interface AddonsCommandActions {
  openAddonsManager: (section?: AddonsSection) => void | Promise<void>;
}

export function createAddonsCommands(actions: AddonsCommandActions): SlashCommand[] {
  return [
    {
      name: "extensions",
      description: "Open Extensions (connections, plugins, skills)",
      source: "builtin",
      execute: () => {
        void actions.openAddonsManager();
      },
    },
  ];
}
