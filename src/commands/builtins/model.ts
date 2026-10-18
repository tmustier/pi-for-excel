/**
 * Builtin model-related commands.
 */

import type { ConversationPresenter } from "../../durable/conversation-view.js";

import type { SlashCommand } from "../types.js";
import { showToast } from "../../ui/toast.js";
import { t } from "../../language/index.js";

export type ActiveConversationProvider = () => ConversationPresenter | null;

export interface ModelCommandActions {
  getActiveConversation: ActiveConversationProvider;
  openModelSelector: () => void;
}

export function createModelCommands(actions: ModelCommandActions): SlashCommand[] {
  const runModelSelector = (): void => {
    const agent = actions.getActiveConversation();
    if (!agent) {
      showToast(t("command.model.no_session"));
      return;
    }

    actions.openModelSelector();
  };

  return [
    {
      name: "model",
      description: t("command.model.desc"),
      source: "builtin",
      execute: runModelSelector,
    },
    {
      name: "default-models",
      description: t("command.model.cycle"),
      source: "builtin",
      execute: () => {
        // TODO: implement scoped models dialog
        // For now, open model selector as a placeholder
        runModelSelector();
      },
    },
  ];
}
