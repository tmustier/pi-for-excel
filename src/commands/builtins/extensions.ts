/**
 * Builtin command for plugin management UI.
 */

import type { ExtensionsHubTab } from "./settings-pages/index.js";
import type { SlashCommand } from "../types.js";
import { t } from "../../language/index.js";

export interface ExtensionsCommandActions {
  openExtensionsHub: (tab?: ExtensionsHubTab) => void | Promise<void>;
}

export function createExtensionsCommands(actions: ExtensionsCommandActions): SlashCommand[] {
  return [
    {
      name: "plugins",
      description: t("command.plugins.desc"),
      source: "builtin",
      execute: () => {
        void actions.openExtensionsHub("plugins");
      },
    },
  ];
}
