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
  getActiveAgent: ActiveAgentProvider;
  openModelSelector: () => void;
}

/** Register all built-in commands. Call once after runtime manager is ready. */
export function registerBuiltins(context: BuiltinsContext): void {
  // Most-used conversation commands lead each group in the discovery menu.
  const builtins: SlashCommand[] = [
    ...createSessionLifecycleCommands(context),
    ...createCompactCommands(context.getActiveAgent),
    ...createModelCommands({
      getActiveAgent: context.getActiveAgent,
      openModelSelector: context.openModelSelector,
    }),
    ...createSettingsCommands(context),
    ...createAddonsCommands(context),
    ...createToolsCommands(context),
    ...createSkillsCommands(context),
    ...createFilesCommands(context),
    ...createExperimentalCommands(),
    ...createDebugCommands(),
    ...createClipboardCommands(context.getActiveAgent),
    ...createExportCommands(context.getActiveAgent),
    ...createSessionIdentityCommands(context),
    ...createHelpCommands(),
    ...createExtensionsCommands(context),
  ];

  for (const cmd of builtins) {
    const menuGroup = ["history", "backup", "revert", "rules", "yolo"].includes(cmd.name) ? "workbook"
      : ["new", "resume", "reopen", "compact", "name", "copy", "export"].includes(cmd.name) ? "chat" : "setup";
    const menuVisibility = ["tools", "plugins", "skills"].includes(cmd.name) ? "hidden"
      : ["debug", "experimental"].includes(cmd.name) ? "exact" : undefined;
    commandRegistry.register({ ...cmd, menuGroup, ...(menuVisibility ? { menuVisibility } : {}) });
  }
}
