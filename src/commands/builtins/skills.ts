/**
 * Builtin command for skill management UI.
 */

import type { SlashCommand } from "../types.js";

export interface SkillsCommandActions {
  openSkillsManager: () => void | Promise<void>;
}

export function createSkillsCommands(actions: SkillsCommandActions): SlashCommand[] {
  return [
    {
      name: "skills",
      description: "Manage skills (web search, MCP)",
      source: "builtin",
      execute: () => {
        void actions.openSkillsManager();
      },
    },
  ];
}
