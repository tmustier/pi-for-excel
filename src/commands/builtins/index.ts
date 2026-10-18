/**
 * Register all builtin slash commands.
 */

import { commandRegistry, type SlashCommand } from "../types.js";

import { createModelCommands, type ActiveConversationProvider } from "./model.js";
import type { DurableSessions } from "../../durable/sessions.js";
import { createSettingsCommands, type SettingsCommandActions } from "./settings.js";
import { createExperimentalCommands } from "./experimental.js";
import { createDebugCommands } from "./debug.js";
import { createClipboardCommands } from "./clipboard.js";
import { createExportCommands, createCompactCommands } from "./export.js";
import { createSessionIdentityCommands, createSessionLifecycleCommands, type SessionCommandActions } from "./session.js";
import { createHelpCommands } from "./help.js";
import { createExtensionsCommands, type ExtensionsCommandActions } from "./extensions.js";
import { createAddonsCommands, type AddonsCommandActions } from "./addons.js";
import { createToolsCommands, type ToolsCommandActions } from "./tools.js";
import { createSkillsCommands, type SkillsCommandActions } from "./skills.js";
import { createFilesCommands, type FilesCommandActions } from "./files.js";

export interface BuiltinsContext
  extends SessionCommandActions,
    SettingsCommandActions,
    AddonsCommandActions,
    ExtensionsCommandActions,
    ToolsCommandActions,
    SkillsCommandActions,
    FilesCommandActions {
  getActiveConversation: ActiveConversationProvider;
  getSessions: () => DurableSessions;
  openModelSelector: () => void;
}

/** Register all built-in commands. Call once after runtime manager is ready. */
export function registerBuiltins(context: BuiltinsContext): void {
  // Keep registration order stable: this is the order shown in the command menu.
  const builtins: SlashCommand[] = [
    ...createModelCommands({
      getActiveConversation: context.getActiveConversation,
      openModelSelector: context.openModelSelector,
    }),
    ...createSettingsCommands(context),
    ...createAddonsCommands(context),
    ...createToolsCommands(context),
    ...createSkillsCommands(context),
    ...createFilesCommands(context),
    ...createExperimentalCommands(),
    ...createDebugCommands(),
    ...createClipboardCommands(context.getActiveConversation),
    ...createExportCommands(context.getActiveConversation, context.getSessions),
    ...createSessionIdentityCommands(context),
    ...createHelpCommands(),
    ...createExtensionsCommands(context),
    ...createSessionLifecycleCommands(context),
    ...createCompactCommands(context.getActiveConversation),
  ];

  for (const cmd of builtins) {
    commandRegistry.register(cmd);
  }
}
