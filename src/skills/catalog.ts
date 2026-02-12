/**
 * Built-in skill catalog.
 *
 * Skills bundle:
 * - additional system guidance
 * - one or more tools
 */

import type { AgentTool } from "@mariozechner/pi-agent-core";

import { createMcpTool } from "../tools/mcp.js";
import { createWebSearchTool } from "../tools/web-search.js";

export const SKILL_IDS = ["web_search", "mcp_tools"] as const;
export type SkillId = (typeof SKILL_IDS)[number];

export interface SkillPromptEntry {
  id: SkillId;
  title: string;
  instructions: string;
  warning?: string;
}

export interface SkillDefinition {
  id: SkillId;
  title: string;
  description: string;
  warning?: string;
  toolNames: readonly string[];
  instructions: string;
  createTools: () => AgentTool[];
}

const SKILL_DEFINITIONS: Record<SkillId, SkillDefinition> = {
  web_search: {
    id: "web_search",
    title: "Web Search",
    description: "Search external web content (Brave Search) for up-to-date facts.",
    warning: "External network access: query text is sent to the search provider.",
    toolNames: ["web_search"],
    instructions:
      "Use web_search when workbook context is insufficient and fresh external facts are needed. "
      + "Cite sources from tool results as [1], [2], etc. Avoid web search when the answer is already in the workbook.",
    createTools: () => [createWebSearchTool()],
  },
  mcp_tools: {
    id: "mcp_tools",
    title: "MCP Gateway",
    description: "Call tools from user-configured MCP servers.",
    warning: "External tools: MCP servers may execute arbitrary remote actions.",
    toolNames: ["mcp"],
    instructions:
      "Use the mcp tool only when a configured external capability is needed. "
      + "Prefer listing/describing tools before invoking them, and clearly state which server/tool was used.",
    createTools: () => [createMcpTool()],
  },
};

export function listSkillDefinitions(): SkillDefinition[] {
  return SKILL_IDS.map((skillId) => SKILL_DEFINITIONS[skillId]);
}

export function getSkillDefinition(skillId: string): SkillDefinition | null {
  if (!Object.hasOwn(SKILL_DEFINITIONS, skillId)) return null;

  if (skillId === "web_search") return SKILL_DEFINITIONS.web_search;
  if (skillId === "mcp_tools") return SKILL_DEFINITIONS.mcp_tools;
  return null;
}

export function createToolsForSkills(skillIds: readonly string[]): AgentTool[] {
  const tools: AgentTool[] = [];

  for (const skillId of skillIds) {
    const definition = getSkillDefinition(skillId);
    if (!definition) continue;
    tools.push(...definition.createTools());
  }

  return tools;
}

export function buildSkillPromptEntries(skillIds: readonly string[]): SkillPromptEntry[] {
  const entries: SkillPromptEntry[] = [];

  for (const skillId of skillIds) {
    const definition = getSkillDefinition(skillId);
    if (!definition) continue;

    entries.push({
      id: definition.id,
      title: definition.title,
      instructions: definition.instructions,
      warning: definition.warning,
    });
  }

  return entries;
}

export function getSkillToolNames(): string[] {
  const names = new Set<string>();

  for (const definition of listSkillDefinitions()) {
    for (const toolName of definition.toolNames) {
      names.add(toolName);
    }
  }

  return Array.from(names.values());
}
