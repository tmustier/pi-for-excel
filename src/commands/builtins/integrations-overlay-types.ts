import type { WebSearchProviderConfig } from "../../tools/web-search-config.js";
import type { McpServerConfig } from "../../tools/mcp-config.js";

export interface WorkbookContextSnapshot {
  workbookId: string | null;
  workbookLabel: string;
}

export interface IntegrationsDialogDependencies {
  getActiveSessionId: () => string | null;
  resolveWorkbookContext: () => Promise<WorkbookContextSnapshot>;
  onChanged?: () => Promise<void> | void;
}

export interface IntegrationsSnapshot {
  sessionId: string;
  workbookId: string | null;
  workbookLabel: string;
  externalToolsEnabled: boolean;
  sessionIntegrationIds: string[];
  workbookIntegrationIds: string[];
  activeIntegrationIds: string[];
  webSearchConfig: WebSearchProviderConfig;
  mcpServers: McpServerConfig[];
}

export type IntegrationMutationReason = "toggle" | "scope" | "external-toggle" | "config";

export function getIntegrationsErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }

  return String(error);
}
