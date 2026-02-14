/**
 * Built-in integration catalog.
 *
 * Integrations bundle:
 * - additional system guidance
 * - one or more tools
 */

import type { AgentTool } from "@mariozechner/pi-agent-core";

import { createMcpTool } from "../tools/mcp.js";
import { createFetchPageTool } from "../tools/fetch-page.js";
import { createWebSearchTool } from "../tools/web-search.js";

export const INTEGRATION_IDS = ["web_search", "mcp_tools"] as const;
export type IntegrationId = (typeof INTEGRATION_IDS)[number];

export interface IntegrationPromptEntry {
  id: IntegrationId;
  title: string;
  instructions: string;
  agentSkillName?: string;
  warning?: string;
}

export interface IntegrationDefinition {
  id: IntegrationId;
  title: string;
  description: string;
  /** Standards-based Agent Skill identity (agentskills.io). */
  agentSkillName?: string;
  warning?: string;
  toolNames: readonly string[];
  instructions: string;
  createTools: () => AgentTool[];
}

const INTEGRATION_DEFINITIONS: Record<IntegrationId, IntegrationDefinition> = {
  web_search: {
    id: "web_search",
    title: "Web Search",
    description: "Search external web content (Serper/Tavily/Brave) and fetch readable page content.",
    agentSkillName: "web-search",
    warning: "External network access: queries and fetched URLs are sent to the configured provider/target host.",
    toolNames: ["web_search", "fetch_page"],
    instructions:
      "Use web_search when workbook context is insufficient and fresh external facts are needed. "
      + "After finding promising URLs, use fetch_page to read page content before synthesizing. "
      + "Cite sources from tool results as [1], [2], etc. Avoid web search when the answer is already in the workbook.",
    createTools: () => [createWebSearchTool(), createFetchPageTool()],
  },
  mcp_tools: {
    id: "mcp_tools",
    title: "MCP Gateway",
    description: "Call tools from user-configured MCP servers.",
    agentSkillName: "mcp-gateway",
    warning: "External tools: MCP servers may execute arbitrary remote actions.",
    toolNames: ["mcp"],
    instructions:
      "Use the mcp tool only when a configured external capability is needed. "
      + "Prefer listing/describing tools before invoking them, and clearly state which server/tool was used.",
    createTools: () => [createMcpTool()],
  },
};

export function listIntegrationDefinitions(): IntegrationDefinition[] {
  return INTEGRATION_IDS.map((integrationId) => INTEGRATION_DEFINITIONS[integrationId]);
}

export function getIntegrationDefinition(integrationId: string): IntegrationDefinition | null {
  if (!Object.hasOwn(INTEGRATION_DEFINITIONS, integrationId)) return null;

  if (integrationId === "web_search") return INTEGRATION_DEFINITIONS.web_search;
  if (integrationId === "mcp_tools") return INTEGRATION_DEFINITIONS.mcp_tools;
  return null;
}

export function createToolsForIntegrations(integrationIds: readonly string[]): AgentTool[] {
  const tools: AgentTool[] = [];

  for (const integrationId of integrationIds) {
    const definition = getIntegrationDefinition(integrationId);
    if (!definition) continue;
    tools.push(...definition.createTools());
  }

  return tools;
}

export function buildIntegrationPromptEntries(integrationIds: readonly string[]): IntegrationPromptEntry[] {
  const entries: IntegrationPromptEntry[] = [];

  for (const integrationId of integrationIds) {
    const definition = getIntegrationDefinition(integrationId);
    if (!definition) continue;

    entries.push({
      id: definition.id,
      title: definition.title,
      instructions: definition.instructions,
      agentSkillName: definition.agentSkillName,
      warning: definition.warning,
    });
  }

  return entries;
}

export function getIntegrationToolNames(): string[] {
  const names = new Set<string>();

  for (const definition of listIntegrationDefinitions()) {
    for (const toolName of definition.toolNames) {
      names.add(toolName);
    }
  }

  return Array.from(names.values());
}
