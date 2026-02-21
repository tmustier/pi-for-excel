/**
 * Bundled Agent Skills catalog.
 *
 * Source of truth: top-level `skills/<skill-name>/SKILL.md` files.
 */

import mcpGatewaySkillMarkdown from "../../skills/mcp-gateway/SKILL.md?raw";
import pythonBridgeSkillMarkdown from "../../skills/python-bridge/SKILL.md?raw";
import tmuxBridgeSkillMarkdown from "../../skills/tmux-bridge/SKILL.md?raw";
import webSearchSkillMarkdown from "../../skills/web-search/SKILL.md?raw";

import { parseSkillDocument, type ParsedSkillFrontmatter } from "./frontmatter.js";
import type {
  AgentSkillDefinition,
  AgentSkillPromptEntry,
  AgentSkillSourceKind,
} from "./types.js";

interface BundledSkillSource {
  location: string;
  markdown: string;
}

const BUNDLED_SKILL_SOURCES: readonly BundledSkillSource[] = [
  {
    location: "skills/web-search/SKILL.md",
    markdown: webSearchSkillMarkdown,
  },
  {
    location: "skills/mcp-gateway/SKILL.md",
    markdown: mcpGatewaySkillMarkdown,
  },
  {
    location: "skills/tmux-bridge/SKILL.md",
    markdown: tmuxBridgeSkillMarkdown,
  },
  {
    location: "skills/python-bridge/SKILL.md",
    markdown: pythonBridgeSkillMarkdown,
  },
] as const;

function buildDefinition(args: {
  location: string;
  markdown: string;
  frontmatter: ParsedSkillFrontmatter;
  sourceKind: AgentSkillSourceKind;
  body: string;
}): AgentSkillDefinition {
  return {
    name: args.frontmatter.name,
    description: args.frontmatter.description,
    compatibility: args.frontmatter.compatibility,
    location: args.location,
    sourceKind: args.sourceKind,
    markdown: args.markdown,
    body: args.body,
  };
}

function buildCatalog(): AgentSkillDefinition[] {
  const definitions: AgentSkillDefinition[] = [];

  for (const source of BUNDLED_SKILL_SOURCES) {
    const parsed = parseSkillDocument(source.markdown);
    if (!parsed) {
      console.warn(`[skills] Invalid SKILL.md frontmatter: ${source.location}`);
      continue;
    }

    definitions.push(buildDefinition({
      location: source.location,
      markdown: source.markdown,
      frontmatter: parsed.frontmatter,
      sourceKind: "bundled",
      body: parsed.body,
    }));
  }

  definitions.sort((left, right) => left.name.localeCompare(right.name));
  return definitions;
}

const CATALOG = buildCatalog();

export function mergeAgentSkillDefinitions(
  preferred: readonly AgentSkillDefinition[],
  fallback: readonly AgentSkillDefinition[],
): AgentSkillDefinition[] {
  const byName = new Map<string, AgentSkillDefinition>();

  for (const skill of preferred) {
    byName.set(skill.name.toLowerCase(), skill);
  }

  for (const skill of fallback) {
    const key = skill.name.toLowerCase();
    if (!byName.has(key)) {
      byName.set(key, skill);
    }
  }

  return Array.from(byName.values()).sort((left, right) => left.name.localeCompare(right.name));
}

export function listAgentSkills(): AgentSkillDefinition[] {
  return [...CATALOG];
}

export function getAgentSkillByName(name: string): AgentSkillDefinition | null {
  const needle = name.trim().toLowerCase();
  if (needle.length === 0) return null;

  const found = CATALOG.find((entry) => entry.name.toLowerCase() === needle);
  return found ?? null;
}

export function buildAgentSkillPromptEntries(skills: readonly AgentSkillDefinition[]): AgentSkillPromptEntry[] {
  return skills.map((entry) => ({
    name: entry.name,
    description: entry.description,
    location: entry.location,
  }));
}

export function getAgentSkillPromptEntries(): AgentSkillPromptEntry[] {
  return buildAgentSkillPromptEntries(CATALOG);
}
