import { mergeAgentSkillDefinitions, listAgentSkills } from "../skills/catalog.js";
import {
  loadExternalAgentSkillsFromSettings,
  removeExternalAgentSkillFromSettings,
  upsertExternalAgentSkillInSettings,
} from "../skills/external-store.js";

export interface SkillsStoreSettings {
  get(key: string): Promise<unknown>;
  set(key: string, value: unknown): Promise<void>;
}

export interface SkillSummaryItem {
  name: string;
  description: string;
  sourceKind: string;
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
  await upsertExternalAgentSkillInSettings({
    settings,
    markdown,
    expectedName: requestedName,
  });
}

export async function uninstallExternalExtensionSkill(settings: SkillsStoreSettings, name: string): Promise<void> {
  await removeExternalAgentSkillFromSettings({
    settings,
    name,
  });
}
