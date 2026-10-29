import type { Api, Model, Models } from "@earendil-works/pi-ai";
import type { ConnectionManager } from "../../../connections/manager.js";
import type { ExecutionMode } from "../../../execution/mode.js";
import type { ExtensionRuntimeManager } from "../../../extensions/runtime-manager.js";
import type { BackupsPageCallbacks } from "./backups-page.js";

export interface WorkbookContextSnapshot {
  workbookId: string | null;
  workbookLabel: string;
}

export interface ExtensionsHubDependencies {
  getActiveModel?: () => Model<Api> | null;
  getActiveSessionId: () => string | null;
  resolveWorkbookContext: () => Promise<WorkbookContextSnapshot>;
  extensionManager: ExtensionRuntimeManager;
  connectionManager: ConnectionManager;
  onChanged?: () => Promise<void> | void;
}



export interface SettingsPagesDependencies {
  getExecutionMode?: () => ExecutionMode;
  setExecutionMode?: (mode: ExecutionMode) => Promise<void>;
  models?: Pick<Models, "getModels">;
  extensionsHub?: ExtensionsHubDependencies;
  onRulesSaved?: () => Promise<void> | void;
  backups?: BackupsPageCallbacks;
  isAnyTabBusy?: () => boolean;
}

let dependencies: SettingsPagesDependencies = {};

export function configureSettingsPages(next: SettingsPagesDependencies): void {
  dependencies = { ...next };
}

export function getSettingsPagesDependencies(): SettingsPagesDependencies {
  return dependencies;
}
