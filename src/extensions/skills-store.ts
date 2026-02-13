import { mergeAgentSkillDefinitions, listAgentSkills } from "../skills/catalog.js";
import { EXTERNAL_AGENT_SKILLS_STORAGE_KEY, loadExternalAgentSkillsFromSettings } from "../skills/external-store.js";
import { parseSkillDocument } from "../skills/frontmatter.js";
import { isRecord } from "../utils/type-guards.js";

interface ExternalSkillItem {
  location: string;
  markdown: string;
}

interface ExternalSkillDocument {
  version: 1;
  items: ExternalSkillItem[];
}

export interface SkillsStoreSettings {
  get(key: string): Promise<unknown>;
  set(key: string, value: unknown): Promise<void>;
}

export interface SkillSummaryItem {
  name: string;
  description: string;
  sourceKind: string;
}

function parseExternalSkillDocument(raw: unknown): ExternalSkillDocument {
  if (!isRecord(raw) || raw.version !== 1 || !Array.isArray(raw.items)) {
    return { version: 1, items: [] };
  }

  const items: ExternalSkillItem[] = [];

  for (const item of raw.items) {
    if (!isRecord(item)) {
      continue;
    }

    if (typeof item.location !== "string" || typeof item.markdown !== "string") {
      continue;
    }

    items.push({
      location: item.location,
      markdown: item.markdown,
    });
  }

  return {
    version: 1,
    items,
  };
}

async function loadMergedSkills(settings: SkillsStoreSettings) {
  const bundled = listAgentSkills();
  const external = await loadExternalAgentSkillsFromSettings(settings);
  return mergeAgentSkillDefinitions(external, bundled);
}

function normalizeSkillName(name: string): string {
  const trimmed = name.trim();
  if (trimmed.length === 0) {
    throw new Error("Skill name cannot be empty.");
  }

  return trimmed;
}

export async function listExtensionSkillSummaries(settings: SkillsStoreSettings): Promise<SkillSummaryItem[]> {
  const merged = await loadMergedSkills(settings);

  return merged.map((skill) => ({
    name: skill.name,
    description: skill.description,
    sourceKind: skill.sourceKind,
  }));
}

export async function readExtensionSkill(settings: SkillsStoreSettings, name: string): Promise<string> {
  const normalizedName = normalizeSkillName(name).toLowerCase();
  const merged = await loadMergedSkills(settings);

  const match = merged.find((skill) => skill.name.toLowerCase() === normalizedName);
  if (!match) {
    throw new Error(`Skill not found: ${name}`);
  }

  return match.markdown;
}

export async function installExternalExtensionSkill(
  settings: SkillsStoreSettings,
  requestedName: string,
  markdown: string,
): Promise<void> {
  const parsed = parseSkillDocument(markdown);
  if (!parsed) {
    throw new Error("Invalid SKILL.md document: expected frontmatter with name and description.");
  }

  const normalizedRequested = normalizeSkillName(requestedName);
  if (parsed.frontmatter.name.toLowerCase() !== normalizedRequested.toLowerCase()) {
    throw new Error(
      `Skill name mismatch: requested "${normalizedRequested}" but markdown declares "${parsed.frontmatter.name}".`,
    );
  }

  const existing = parseExternalSkillDocument(await settings.get(EXTERNAL_AGENT_SKILLS_STORAGE_KEY));
  const targetName = parsed.frontmatter.name.toLowerCase();

  const nextItems = existing.items.filter((item) => {
    const itemParsed = parseSkillDocument(item.markdown);
    if (!itemParsed) {
      return true;
    }

    return itemParsed.frontmatter.name.toLowerCase() !== targetName;
  });

  nextItems.push({
    location: `skills/external/${parsed.frontmatter.name}/SKILL.md`,
    markdown,
  });

  await settings.set(EXTERNAL_AGENT_SKILLS_STORAGE_KEY, {
    version: 1,
    items: nextItems,
  });
}

export async function uninstallExternalExtensionSkill(settings: SkillsStoreSettings, name: string): Promise<void> {
  const normalizedName = normalizeSkillName(name).toLowerCase();
  const existing = parseExternalSkillDocument(await settings.get(EXTERNAL_AGENT_SKILLS_STORAGE_KEY));

  const nextItems = existing.items.filter((item) => {
    const itemParsed = parseSkillDocument(item.markdown);
    if (!itemParsed) {
      return true;
    }

    return itemParsed.frontmatter.name.toLowerCase() !== normalizedName;
  });

  await settings.set(EXTERNAL_AGENT_SKILLS_STORAGE_KEY, {
    version: 1,
    items: nextItems,
  });
}
