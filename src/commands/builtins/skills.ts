/**
 * Builtin command for skills catalog UI.
 */

import type { SlashCommand } from "../types.js";

export interface SkillsCommandActions {
  openSkillsManager: () => void | Promise<void>;
}

export function createSkillsCommands(actions: SkillsCommandActions): SlashCommand[] {
  return [
    {
      name: "skills",
      description: "Browse available skills (bundled + external)",
      source: "builtin",
      execute: () => {
        void actions.openSkillsManager();
      },
    },
  ];
}
