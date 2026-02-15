/**
 * Builtin command for skills catalog UI.
 */

import type { AddonsSection } from "./addons-overlay.js";
import type { SlashCommand } from "../types.js";

export interface SkillsCommandActions {
  openAddonsManager: (section?: AddonsSection) => void | Promise<void>;
}

export function createSkillsCommands(actions: SkillsCommandActions): SlashCommand[] {
  return [
    {
      name: "skills",
      description: "Browse available skills (alias for /extensions skills)",
      source: "builtin",
      execute: () => {
        void actions.openAddonsManager("skills");
      },
    },
  ];
}
