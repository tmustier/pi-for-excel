/**
 * Builtin session management commands.
 */

import type { SlashCommand } from "../types.js";
import type { ResumeDialogTarget } from "./resume-target.js";
import { showToast } from "../../ui/toast.js";

export interface SessionCommandActions {
  renameActiveSession: (title: string) => Promise<void>;
  createRuntime: () => Promise<void>;
  openResumeDialog: (defaultTarget?: ResumeDialogTarget) => Promise<void>;
  reopenLastClosed: () => Promise<void>;
}

export function createSessionIdentityCommands(actions: SessionCommandActions): SlashCommand[] {
  return [
    {
      name: "name",
      description: "Name the current chat session",
      source: "builtin",
      execute: async (args: string) => {
        const title = args.trim();
        if (!title) {
          showToast("Usage: /name My Session Name");
          return;
        }

        await actions.renameActiveSession(title);
        showToast(`Session named: ${title}`);
      },
    },
    {
      name: "share-session",
      description: "Share session as a link",
      source: "builtin",
      execute: () => {
        showToast("Session sharing coming soon");
      },
    },
  ];
}

export function createSessionLifecycleCommands(actions: SessionCommandActions): SlashCommand[] {
  return [
    {
      name: "new",
      description: "Start a new chat session tab",
      source: "builtin",
      execute: async () => {
        await actions.createRuntime();
      },
    },
    {
      name: "resume",
      description: "Resume a previous session (opens in new tab)",
      source: "builtin",
      execute: async () => {
        await actions.openResumeDialog("new_tab");
      },
    },
    {
      name: "resume-here",
      description: "Resume a previous session into the current tab",
      source: "builtin",
      execute: async () => {
        await actions.openResumeDialog("replace_current");
      },
    },
    {
      name: "reopen",
      description: "Reopen the most recently closed session tab",
      source: "builtin",
      execute: async () => {
        await actions.reopenLastClosed();
      },
    },
  ];
}
