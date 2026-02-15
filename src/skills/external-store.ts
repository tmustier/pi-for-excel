/**
 * External Agent Skills discovery store (feature-flagged).
 *
 * This is intentionally opt-in and local-only.
 */

import type {
  AgentSkillDefinition,
  AgentSkillSourceKind,
} from "./types.js";
import { parseSkillDocument } from "./frontmatter.js";
import { isRecord } from "../utils/type-guards.js";

export const EXTERNAL_AGENT_SKILLS_STORAGE_KEY = "skills.external.v1.catalog";

export interface ExternalSkillSettingsStore {
  get: (key: string) => Promise<unknown>;
}

export interface ExternalSkillMutableSettingsStore extends ExternalSkillSettingsStore {
  set: (key: string, value: unknown) => Promise<void>;
}

interface StoredExternalSkillItem {
  location: string;
  markdown: string;
}

function isStoredExternalSkillItem(value: unknown): value is StoredExternalSkillItem {
  if (!isRecord(value)) return false;

  return (
    typeof value.location === "string"
    && typeof value.markdown === "string"
  );
}

function parseStoredExternalSkillItems(raw: unknown): StoredExternalSkillItem[] {
  if (!isRecord(raw)) return [];

  const version = raw.version;
  if (version !== 1) return [];

  const items = raw.items;
  if (!Array.isArray(items)) return [];

  return items.filter((item) => isStoredExternalSkillItem(item));
}

function normalizeSkillName(name: string): string {
  const trimmed = name.trim();
  if (trimmed.length === 0) {
    throw new Error("Skill name cannot be empty.");
  }

  return trimmed;
}

function buildExternalSkillDefinition(args: {
  item: StoredExternalSkillItem;
  sourceKind: AgentSkillSourceKind;
}): AgentSkillDefinition | null {
  const parsed = parseSkillDocument(args.item.markdown);
  if (!parsed) {
    return null;
  }

  return {
    name: parsed.frontmatter.name,
    description: parsed.frontmatter.description,
    compatibility: parsed.frontmatter.compatibility,
    location: args.item.location,
    sourceKind: args.sourceKind,
    markdown: args.item.markdown,
    body: parsed.body,
  };
}

/**
 * Loads externally configured skills from settings.
 *
 * Expected shape:
 * {
 *   version: 1,
 *   items: [{ location: string, markdown: string }, ...]
 * }
 */
export async function loadExternalAgentSkillsFromSettings(
  settings: ExternalSkillSettingsStore,
): Promise<AgentSkillDefinition[]> {
  const raw = await settings.get(EXTERNAL_AGENT_SKILLS_STORAGE_KEY);
  const storedItems = parseStoredExternalSkillItems(raw);

  const loaded: AgentSkillDefinition[] = [];

  for (const item of storedItems) {
    const skill = buildExternalSkillDefinition({
      item,
      sourceKind: "external",
    });

    if (!skill) {
      console.warn(`[skills] Invalid external SKILL.md frontmatter: ${item.location}`);
      continue;
    }

    loaded.push(skill);
  }

  loaded.sort((left, right) => left.name.localeCompare(right.name));
  return loaded;
}

export interface UpsertExternalAgentSkillResult {
  name: string;
  location: string;
}

export async function upsertExternalAgentSkillInSettings(args: {
  settings: ExternalSkillMutableSettingsStore;
  markdown: string;
  expectedName?: string;
}): Promise<UpsertExternalAgentSkillResult> {
  const parsed = parseSkillDocument(args.markdown);
  if (!parsed) {
    throw new Error("Invalid SKILL.md document: expected frontmatter with name and description.");
  }

  if (args.expectedName !== undefined) {
    const normalizedExpected = normalizeSkillName(args.expectedName);
    if (parsed.frontmatter.name.toLowerCase() !== normalizedExpected.toLowerCase()) {
      throw new Error(
        `Skill name mismatch: expected "${normalizedExpected}" but markdown declares "${parsed.frontmatter.name}".`,
      );
    }
  }

  const existing = parseStoredExternalSkillItems(await args.settings.get(EXTERNAL_AGENT_SKILLS_STORAGE_KEY));
  const targetName = parsed.frontmatter.name.toLowerCase();

  const nextItems = existing.filter((item) => {
    const itemParsed = parseSkillDocument(item.markdown);
    if (!itemParsed) {
      return true;
    }

    return itemParsed.frontmatter.name.toLowerCase() !== targetName;
  });

  const location = `skills/external/${parsed.frontmatter.name}/SKILL.md`;
  nextItems.push({
    location,
    markdown: args.markdown,
  });

  await args.settings.set(EXTERNAL_AGENT_SKILLS_STORAGE_KEY, {
    version: 1,
    items: nextItems,
  });

  return {
    name: parsed.frontmatter.name,
    location,
  };
}

export async function removeExternalAgentSkillFromSettings(args: {
  settings: ExternalSkillMutableSettingsStore;
  name: string;
}): Promise<boolean> {
  const normalizedName = normalizeSkillName(args.name).toLowerCase();
  const existing = parseStoredExternalSkillItems(await args.settings.get(EXTERNAL_AGENT_SKILLS_STORAGE_KEY));

  let removed = false;
  const nextItems = existing.filter((item) => {
    const itemParsed = parseSkillDocument(item.markdown);
    if (!itemParsed) {
      return true;
    }

    const keep = itemParsed.frontmatter.name.toLowerCase() !== normalizedName;
    if (!keep) {
      removed = true;
    }

    return keep;
  });

  await args.settings.set(EXTERNAL_AGENT_SKILLS_STORAGE_KEY, {
    version: 1,
    items: nextItems,
  });

  return removed;
}
