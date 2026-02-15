/**
 * Builtin command for skills catalog UI.
 */

import type { ExtensionsHubTab } from "./extensions-hub-overlay.js";
import type { SlashCommand } from "../types.js";

export interface SkillsCommandActions {
  openExtensionsHub: (tab?: ExtensionsHubTab) => void | Promise<void>;
}

export function createSkillsCommands(actions: SkillsCommandActions): SlashCommand[] {
  return [
    {
      name: "skills",
      description: "Browse available skills (alias for /extensions skills)",
      source: "builtin",
      execute: () => {
        void actions.openExtensionsHub("skills");
      },
    },
  ];
}
