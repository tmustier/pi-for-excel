/**
 * Builtin command for the Files workspace overlay.
 */

import type { SlashCommand } from "../types.js";
import { t } from "../../language/index.js";

export interface FilesCommandActions {
  openFilesWorkspace: () => void | Promise<void>;
}

export function createFilesCommands(actions: FilesCommandActions): SlashCommand[] {
  return [
    {
      name: "files",
      description: t("command.files.desc"),
      source: "builtin",
      execute: () => {
        void actions.openFilesWorkspace();
      },
    },
  ];
}
