/**
 * Bundled Agent Skills catalog.
 *
 * Source of truth: top-level `skills/<skill-name>/SKILL.md` files.
 */

import { loadRawMarkdownFromTestGlob } from "../utils/test-raw-markdown-glob.js";
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

function toRepoRelativePath(globPath: string): string {
  const normalized = globPath.replaceAll("\\", "/");
  if (normalized.startsWith("./")) {
    return normalized.slice(2);
  }

  let cursor = normalized;
  while (cursor.startsWith("../")) {
    cursor = cursor.slice(3);
  }

  return cursor;
}

function isBundledSkillLocation(value: string): boolean {
  if (!value.startsWith("skills/")) {
    return false;
  }

  const segments = value.split("/");
  if (segments.length !== 3) {
    return false;
  }

  return segments[2] === "SKILL.md";
}

function readBundledSkillMarkdownByPath(): Record<string, string> {
  try {
    return import.meta.glob<string>("../../skills/*/SKILL.md", {
      eager: true,
      query: "?raw",
      import: "default",
    });
  } catch {
    return loadRawMarkdownFromTestGlob("../../skills/*/SKILL.md", import.meta.url);
  }
}

function buildBundledSkillSources(): BundledSkillSource[] {
  return Object.entries(readBundledSkillMarkdownByPath())
    .map(([globPath, markdown]) => ({
      location: toRepoRelativePath(globPath),
      markdown,
    }))
    .filter((source) => isBundledSkillLocation(source.location))
    .sort((left, right) => left.location.localeCompare(right.location));
}

const BUNDLED_SKILL_SOURCES: readonly BundledSkillSource[] = buildBundledSkillSources();

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
