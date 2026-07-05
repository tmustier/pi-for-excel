/**
 * Debug helpers.
 *
 * Intentionally lightweight: this is for temporary instrumentation toggles.
 */

import type { SlashCommand } from "../types.js";
import { showToast } from "../../ui/toast.js";
import { t } from "../../language/index.js";
import { isDebugEnabled, setDebugEnabled, toggleDebugEnabled } from "../../debug/debug.js";

function normalizeArg(args: string): string {
  return args.trim().toLowerCase();
}

export function createDebugCommands(): SlashCommand[] {
  return [
    {
      name: "debug",
      description: t("command.debug.toggle"),
      source: "builtin",
      execute: (args: string) => {
        const a = normalizeArg(args);

        if (a === "" || a === "toggle") {
          const enabled = toggleDebugEnabled();
          showToast(t(enabled ? "command.debug.enabled" : "command.debug.disabled"));
          return;
        }

        if (a === "on" || a === "true" || a === "1") {
          setDebugEnabled(true);
          showToast(t("command.debug.enabled"));
          return;
        }

        if (a === "off" || a === "false" || a === "0") {
          setDebugEnabled(false);
          showToast(t("command.debug.disabled"));
          return;
        }

        if (a === "status") {
          showToast(t(isDebugEnabled() ? "command.debug.is_enabled" : "command.debug.is_disabled"));
          return;
        }

        showToast(t("command.debug.usage"));
      },
    },
  ];
}
