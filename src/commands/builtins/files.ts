/**
 * Builtin command for the Files workspace overlay.
 */

import type { SlashCommand } from "../types.js";

export interface FilesCommandActions {
  openFilesWorkspace: () => void | Promise<void>;
}

export function createFilesCommands(actions: FilesCommandActions): SlashCommand[] {
  return [
    {
      name: "files",
      description: "Browse workspace files",
      source: "builtin",
      execute: () => {
        void actions.openFilesWorkspace();
      },
    },
  ];
}
