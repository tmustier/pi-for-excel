import { mergeAgentSkillDefinitions, listAgentSkills } from "../skills/catalog.js";
import {
  loadExternalAgentSkills,
  removeExternalAgentSkill,
  upsertExternalAgentSkill,
} from "../skills/external-store.js";

export interface SkillSummaryItem {
  name: string;
  description: string;
  sourceKind: string;
}

async function loadMergedSkills() {
  const bundled = listAgentSkills();
  const external = await loadExternalAgentSkills();
  return mergeAgentSkillDefinitions(bundled, external);
}

function normalizeSkillName(name: string): string {
  const trimmed = name.trim();
  if (trimmed.length === 0) {
    throw new Error("Skill name cannot be empty.");
  }

  return trimmed;
}

export async function listExtensionSkillSummaries(): Promise<SkillSummaryItem[]> {
  const merged = await loadMergedSkills();

  return merged.map((skill) => ({
    name: skill.name,
    description: skill.description,
    sourceKind: skill.sourceKind,
  }));
}

export async function readExtensionSkill(name: string): Promise<string> {
  const normalizedName = normalizeSkillName(name).toLowerCase();
  const merged = await loadMergedSkills();

  const match = merged.find((skill) => skill.name.toLowerCase() === normalizedName);
  if (!match) {
    throw new Error(`Skill not found: ${name}`);
  }

  return match.markdown;
}

export async function installExternalExtensionSkill(
  requestedName: string,
  markdown: string,
): Promise<void> {
  await upsertExternalAgentSkill({
    markdown,
    expectedName: requestedName,
  });
}

export async function uninstallExternalExtensionSkill(name: string): Promise<void> {
  await removeExternalAgentSkill({
    name,
  });
}
