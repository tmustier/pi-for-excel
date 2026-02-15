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
import { createAddonsCommands, type AddonsCommandActions } from "./addons.js";
import { createIntegrationsCommands, type IntegrationsCommandActions } from "./integrations.js";
import { createSkillsCommands, type SkillsCommandActions } from "./skills.js";

export interface BuiltinsContext
  extends SessionCommandActions,
    SettingsCommandActions,
    AddonsCommandActions,
    ExtensionsCommandActions,
    IntegrationsCommandActions,
    SkillsCommandActions {
  getActiveAgent: ActiveAgentProvider;
}

/** Register all built-in commands. Call once after runtime manager is ready. */
export function registerBuiltins(context: BuiltinsContext): void {
  // Keep registration order stable: this is the order shown in the command menu.
  const builtins: SlashCommand[] = [
    ...createModelCommands(context.getActiveAgent),
    ...createSettingsCommands(context),
    ...createAddonsCommands(context),
    ...createIntegrationsCommands(context),
    ...createSkillsCommands(context),
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
