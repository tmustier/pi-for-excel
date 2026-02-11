/**
 * Register all builtin slash commands.
 */

import { commandRegistry, type SlashCommand } from "../types.js";

import { createModelCommands, type ActiveAgentProvider } from "./model.js";
import { createSettingsCommands, type SettingsCommandActions } from "./settings.js";
import { createExperimentalCommands } from "./experimental.js";
import { createDebugCommands } from "./debug.js";
import { createClipboardCommands } from "./clipboard.js";
import { createExportCommands, createCompactCommands } from "./export.js";
import { createSessionIdentityCommands, createSessionLifecycleCommands, type SessionCommandActions } from "./session.js";
import { createHelpCommands } from "./help.js";
import { createExtensionsCommands, type ExtensionsCommandActions } from "./extensions.js";

export interface BuiltinsContext
  extends SessionCommandActions, SettingsCommandActions, ExtensionsCommandActions {
  getActiveAgent: ActiveAgentProvider;
}

/** Register all built-in commands. Call once after runtime manager is ready. */
export function registerBuiltins(context: BuiltinsContext): void {
  // Keep registration order stable: this is the order shown in the command menu.
  const builtins: SlashCommand[] = [
    ...createModelCommands(context.getActiveAgent),
    ...createSettingsCommands(context),
    ...createExperimentalCommands(),
    ...createDebugCommands(),
    ...createClipboardCommands(context.getActiveAgent),
    ...createExportCommands(context.getActiveAgent),
    ...createSessionIdentityCommands(context),
    ...createHelpCommands(),
    ...createExtensionsCommands(context),
    ...createSessionLifecycleCommands(context),
    ...createCompactCommands(context.getActiveAgent),
  ];

  for (const cmd of builtins) {
    commandRegistry.register(cmd);
  }
}
