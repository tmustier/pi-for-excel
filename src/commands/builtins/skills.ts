/**
 * Builtin command for skills catalog UI.
 */

import type { ExtensionsHubTab } from "./settings-pages/index.js";
import type { SlashCommand } from "../types.js";
import { t } from "../../language/index.js";

export interface SkillsCommandActions {
  openExtensionsHub: (tab?: ExtensionsHubTab) => void | Promise<void>;
}

export function createSkillsCommands(actions: SkillsCommandActions): SlashCommand[] {
  return [
    {
      name: "skills",
      description: t("command.skills.desc"),
      source: "builtin",
      execute: () => {
        void actions.openExtensionsHub("skills");
      },
    },
  ];
}
