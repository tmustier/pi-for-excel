/**
 * Runtime dependencies for the settings pages, wired once from
 * `taskpane/init.ts` via `configureSettingsPages`.
 */

import type { ConnectionManager } from "../../../connections/manager.js";
import type { ExecutionMode } from "../../../execution/mode.js";
import type { ExtensionRuntimeManager } from "../../../extensions/runtime-manager.js";
import type { ModelSwitchBehavior } from "../../../models/switch-behavior.js";
import type { BackupsPageCallbacks } from "./backups-page.js";

export interface WorkbookContextSnapshot {
  workbookId: string | null;
  workbookLabel: string;
}

/** Dependencies shared by the extensions pages (connections/plugins/skills). */
export interface ExtensionsHubDependencies {
  getActiveSessionId: () => string | null;
  resolveWorkbookContext: () => Promise<WorkbookContextSnapshot>;
  extensionManager: ExtensionRuntimeManager;
  connectionManager: ConnectionManager;
  onChanged?: () => Promise<void> | void;
}



export interface SettingsPagesDependencies {
  getExecutionMode?: () => ExecutionMode;
  setExecutionMode?: (mode: ExecutionMode) => Promise<void>;
  getModelSwitchBehavior?: () => ModelSwitchBehavior;
  setModelSwitchBehavior?: (behavior: ModelSwitchBehavior) => Promise<void>;
  extensionsHub?: ExtensionsHubDependencies;
  onRulesSaved?: () => Promise<void> | void;
  backups?: BackupsPageCallbacks;
}

let dependencies: SettingsPagesDependencies = {};

export function configureSettingsPages(next: SettingsPagesDependencies): void {
  dependencies = { ...next };
}

export function getSettingsPagesDependencies(): SettingsPagesDependencies {
  return dependencies;
}
