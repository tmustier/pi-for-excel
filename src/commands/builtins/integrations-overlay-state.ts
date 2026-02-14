import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import { INTEGRATION_IDS } from "../../integrations/catalog.js";
import {
  getExternalToolsEnabled,
  getSessionIntegrationIds,
  getWorkbookIntegrationIds,
  resolveConfiguredIntegrationIds,
  type IntegrationSettingsStore,
} from "../../integrations/store.js";
import {
  loadWebSearchProviderConfig,
  type WebSearchConfigStore,
  type WebSearchProvider,
} from "../../tools/web-search-config.js";
import {
  loadMcpServers,
  type McpConfigStore,
} from "../../tools/mcp-config.js";
import type {
  IntegrationsDialogDependencies,
  IntegrationsSnapshot,
} from "./integrations-overlay-types.js";

export type IntegrationsSettingsStore = IntegrationSettingsStore & WebSearchConfigStore & McpConfigStore;

export function normalizeWebSearchProvider(value: string): WebSearchProvider {
  if (value === "serper" || value === "tavily" || value === "brave") {
    return value;
  }

  return "serper";
}

export function getSettingsStore(): Promise<IntegrationsSettingsStore> {
  return Promise.resolve(getAppStorage().settings);
}

export async function buildSnapshot(
  dependencies: IntegrationsDialogDependencies,
): Promise<IntegrationsSnapshot> {
  const settings = await getSettingsStore();
  const sessionId = dependencies.getActiveSessionId();
  if (!sessionId) {
    throw new Error("No active session.");
  }

  const workbookContext = await dependencies.resolveWorkbookContext();

  const [
    externalToolsEnabled,
    sessionIntegrationIds,
    workbookIntegrationIds,
    activeIntegrationIds,
    webSearchConfig,
    mcpServers,
  ] = await Promise.all([
    getExternalToolsEnabled(settings),
    getSessionIntegrationIds(settings, sessionId, INTEGRATION_IDS),
    workbookContext.workbookId
      ? getWorkbookIntegrationIds(settings, workbookContext.workbookId, INTEGRATION_IDS)
      : Promise.resolve([]),
    resolveConfiguredIntegrationIds({
      settings,
      sessionId,
      workbookId: workbookContext.workbookId,
      knownIntegrationIds: INTEGRATION_IDS,
    }),
    loadWebSearchProviderConfig(settings),
    loadMcpServers(settings),
  ]);

  return {
    sessionId,
    workbookId: workbookContext.workbookId,
    workbookLabel: workbookContext.workbookLabel,
    externalToolsEnabled,
    sessionIntegrationIds,
    workbookIntegrationIds,
    activeIntegrationIds,
    webSearchConfig,
    mcpServers,
  };
}
