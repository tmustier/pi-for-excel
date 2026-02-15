import type { ExtensionRuntimeStatus } from "../../extensions/runtime-manager.js";
import type { WorkbookContextSnapshot } from "./integrations-overlay-types.js";
import type {
  IntegrationSettingsStore,
} from "../../integrations/store.js";
import type { WebSearchConfigStore } from "../../tools/web-search-config.js";
import type { McpConfigStore } from "../../tools/mcp-config.js";
import type { AgentSkillDefinition } from "../../skills/types.js";

export type AddonsSection = "connections" | "extensions" | "skills";

export interface ShowAddonsDialogOptions {
  section?: AddonsSection;
}

export interface AddonsSettingsStore extends IntegrationSettingsStore, WebSearchConfigStore, McpConfigStore {
  delete?: (key: string) => Promise<void>;
}

export interface AddonsDialogActions {
  getActiveSessionId: () => string | null;
  resolveWorkbookContext: () => Promise<WorkbookContextSnapshot>;
  onChanged?: () => Promise<void> | void;
  openIntegrationsManager: () => void;
  openSkillsManager: () => void;
  openExtensionsManager: () => void;
  listExtensions: () => ExtensionRuntimeStatus[];
  setExtensionEnabled: (entryId: string, enabled: boolean) => Promise<void>;
}

export interface ConnectionsSnapshot {
  sessionId: string | null;
  workbookId: string | null;
  workbookLabel: string;
  externalToolsEnabled: boolean;
  sessionIntegrationIds: string[];
  workbookIntegrationIds: string[];
  activeIntegrationIds: string[];
  webSearchStatusText: string;
  mcpStatusText: string;
  pythonBridgeUrl: string;
}

export interface SkillsSnapshot {
  skills: AgentSkillDefinition[];
  activeNames: Set<string>;
  disabledNames: Set<string>;
  externalDiscoveryEnabled: boolean;
  externalLoadError: string | null;
  activationLoadError: string | null;
}
