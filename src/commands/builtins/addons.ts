/**
 * Builtin commands for unified extensions management UI.
 */

import type { ExtensionsHubTab } from "./extensions-hub-overlay.js";
import type { SlashCommand } from "../types.js";
import { t } from "../../language/index.js";

export interface AddonsCommandActions {
  openExtensionsHub: (tab?: ExtensionsHubTab) => void | Promise<void>;
}

export function createAddonsCommands(actions: AddonsCommandActions): SlashCommand[] {
  return [
    {
      name: "extensions",
      description: t("command.addons.desc"),
      source: "builtin",
      execute: () => {
        void actions.openExtensionsHub();
      },
    },
  ];
}
