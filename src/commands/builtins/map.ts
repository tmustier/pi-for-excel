/**
 * Workbook map command — /map opens the workbook X-ray overlay.
 */

import { getCurrentSpreadsheetHost } from "../../host/current.js";
import { t } from "../../language/index.js";
import { showToast } from "../../ui/toast.js";
import { showWorkbookMapDialog } from "../../ui/workbook-map.js";
import type { SlashCommand } from "../types.js";

export function createMapCommands(): SlashCommand[] {
  return [
    {
      name: "map",
      description: t("command.map.desc"),
      source: "builtin",
      execute: () => {
        if (getCurrentSpreadsheetHost().kind !== "office") {
          showToast(t("map.toast.unsupportedHost"));
          return;
        }
        showWorkbookMapDialog();
      },
    },
  ];
}
