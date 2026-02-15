/**
 * Builtin command for unified add-ons management UI.
 */

import type { AddonsSection } from "./addons-overlay.js";
import type { SlashCommand } from "../types.js";

export interface AddonsCommandActions {
  openAddonsManager: (section?: AddonsSection) => void | Promise<void>;
}

export function createAddonsCommands(actions: AddonsCommandActions): SlashCommand[] {
  return [
    {
      name: "addons",
      description: "Open Add-ons (connections, extensions, skills)",
      source: "builtin",
      execute: () => {
        void actions.openAddonsManager();
      },
    },
  ];
}
