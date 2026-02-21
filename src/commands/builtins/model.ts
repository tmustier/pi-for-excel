/**
 * Builtin model-related commands.
 */

import type { Agent } from "@mariozechner/pi-agent-core";

import type { SlashCommand } from "../types.js";
import { showToast } from "../../ui/toast.js";

export type ActiveAgentProvider = () => Agent | null;

export interface ModelCommandActions {
  getActiveAgent: ActiveAgentProvider;
  openModelSelector: () => void;
}

export function createModelCommands(actions: ModelCommandActions): SlashCommand[] {
  const runModelSelector = (): void => {
    const agent = actions.getActiveAgent();
    if (!agent) {
      showToast("No active session");
      return;
    }

    actions.openModelSelector();
  };

  return [
    {
      name: "model",
      description: "Change the AI model",
      source: "builtin",
      execute: runModelSelector,
    },
    {
      name: "default-models",
      description: "Cycle models with Ctrl+P",
      source: "builtin",
      execute: () => {
        // TODO: implement scoped models dialog
        // For now, open model selector as a placeholder
        runModelSelector();
      },
    },
  ];
}
