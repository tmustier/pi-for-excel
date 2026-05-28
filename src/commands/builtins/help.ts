/**
 * Builtin help / UX commands.
 */

import type { SlashCommand } from "../types.js";
import { showShortcutsDialog } from "./overlays.js";
import { t } from "../../language/index.js";

export function createHelpCommands(): SlashCommand[] {
  return [
    {
      name: "shortcuts",
      description: t("command.help.desc"),
      source: "builtin",
      execute: () => {
        showShortcutsDialog();
      },
    },
  ];
}
