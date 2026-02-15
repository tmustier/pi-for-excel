/**
 * Builtin settings/auth commands.
 */

import { ApiKeysTab, ProxyTab, SettingsDialog } from "@mariozechner/pi-web-ui/dist/dialogs/SettingsDialog.js";

import type { ExecutionMode } from "../../execution/mode.js";
import { formatExecutionModeLabel, toggleExecutionMode } from "../../execution/mode.js";
import { showToast } from "../../ui/toast.js";
import type { SlashCommand } from "../types.js";
import { showProviderPicker } from "./overlays.js";

export interface SettingsCommandActions {
  openInstructionsEditor: () => Promise<void>;
  getExecutionMode: () => Promise<ExecutionMode>;
  setExecutionMode: (mode: ExecutionMode) => Promise<void>;
}

function modeUsageText(): string {
  return "Usage: /yolo [on|off|toggle|status]";
}

function modeDescription(mode: ExecutionMode): string {
  if (mode === "yolo") {
    return `${formatExecutionModeLabel(mode)} mode — Pi applies workbook changes immediately.`;
  }

  return `${formatExecutionModeLabel(mode)} mode — Pi asks before each workbook change.`;
}

function parseExecutionModeArg(input: string): "status" | "toggle" | ExecutionMode | null {
  const normalized = input.trim().toLowerCase();
  if (normalized.length === 0) return "status";

  if (normalized === "status" || normalized === "show" || normalized === "get") {
    return "status";
  }

  if (normalized === "toggle") {
    return "toggle";
  }

  if (normalized === "on" || normalized === "yolo" || normalized === "fast") {
    return "yolo";
  }

  if (normalized === "off" || normalized === "safe" || normalized === "cautious") {
    return "safe";
  }

  return null;
}

export function createSettingsCommands(actions: SettingsCommandActions): SlashCommand[] {
  return [
    {
      name: "settings",
      description: "Settings (API keys + CORS proxy)",
      source: "builtin",
      execute: () => {
        void SettingsDialog.open([new ApiKeysTab(), new ProxyTab()]);
      },
    },
    {
      name: "login",
      description: "Add or change provider API keys",
      source: "builtin",
      execute: async () => {
        await showProviderPicker();
      },
    },
    {
      name: "yolo",
      description: "Toggle execution mode (Auto vs Confirm)",
      source: "builtin",
      execute: async (args: string) => {
        const command = parseExecutionModeArg(args);
        if (!command) {
          showToast(modeUsageText());
          return;
        }

        const currentMode = await actions.getExecutionMode();
        if (command === "status") {
          showToast(modeDescription(currentMode));
          return;
        }

        const nextMode = command === "toggle"
          ? toggleExecutionMode(currentMode)
          : command;

        if (nextMode === currentMode) {
          showToast(modeDescription(currentMode));
          return;
        }

        await actions.setExecutionMode(nextMode);
        showToast(modeDescription(nextMode));
      },
    },
    {
      name: "rules",
      description: "Edit rules for Pi (all files + this file)",
      source: "builtin",
      execute: async () => {
        await actions.openInstructionsEditor();
      },
    },
  ];
}
