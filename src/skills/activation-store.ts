function isSkillsActivationStorePayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Agent Skill activation persistence.
 *
 * Stores disabled skill names so users can opt out of prompt injection per skill.
 */

import type { AgentSkillDefinition } from "./types.js";

export const SKILL_ACTIVATION_STORAGE_KEY = "skills.activation.v1";

export interface SkillActivationSettingsStore {
  get: (key: string) => Promise<DynamicValue>;
}

export interface SkillActivationMutableSettingsStore extends SkillActivationSettingsStore {
  set: (key: string, value: DynamicValue) => Promise<void>;
}

interface StoredSkillActivationDocument {
  version: 1;
  disabledNames: string[];
}

function normalizeSkillName(name: string): string {
  const normalized = name.trim().toLowerCase();
  if (normalized.length === 0) {
    throw new Error("Skill name cannot be empty.");
  }

  return normalized;
}

function normalizeSkillNamesList(raw: DynamicValue): string[] {
  if (!Array.isArray(raw)) {
    return [];
  }

  const normalized = new Set<string>();

  for (const value of raw) {
    if (typeof value !== "string") {
      continue;
    }

    const trimmed = value.trim().toLowerCase();
    if (trimmed.length === 0) {
      continue;
    }

    normalized.add(trimmed);
  }

  return Array.from(normalized).sort((left, right) => left.localeCompare(right));
}

function parseStoredDisabledSkillNames(raw: DynamicValue): string[] {
  if (Array.isArray(raw)) {
    return normalizeSkillNamesList(raw);
  }

  if (!isSkillsActivationStorePayloadShape(raw) || raw.version !== 1) {
    return [];
  }

  return normalizeSkillNamesList(raw.disabledNames);
}

function buildStoredDocument(disabledNames: string[]): StoredSkillActivationDocument {
  return {
    version: 1,
    disabledNames,
  };
}

function arrayShallowEqual(left: readonly string[], right: readonly string[]): boolean {
  if (left.length !== right.length) {
    return false;
  }

  for (let i = 0; i < left.length; i += 1) {
    if (left[i] !== right[i]) {
      return false;
    }
  }

  return true;
}

export async function loadDisabledSkillNamesFromSettings(
  settings: SkillActivationSettingsStore,
): Promise<Set<string>> {
  const raw = await settings.get(SKILL_ACTIVATION_STORAGE_KEY);
  const names = parseStoredDisabledSkillNames(raw);
  return new Set(names);
}

export function filterAgentSkillsByEnabledState(args: {
  skills: readonly AgentSkillDefinition[];
  disabledSkillNames: ReadonlySet<string>;
}): AgentSkillDefinition[] {
  return args.skills.filter((skill) => {
    const normalized = skill.name.trim().toLowerCase();
    return !args.disabledSkillNames.has(normalized);
  });
}

export interface SetSkillEnabledResult {
  name: string;
  enabled: boolean;
  changed: boolean;
}

export async function setSkillEnabledInSettings(args: {
  settings: SkillActivationMutableSettingsStore;
  name: string;
  enabled: boolean;
}): Promise<SetSkillEnabledResult> {
  const normalizedName = normalizeSkillName(args.name);
  const existingNames = parseStoredDisabledSkillNames(await args.settings.get(SKILL_ACTIVATION_STORAGE_KEY));
  const nextDisabled = new Set(existingNames);

  if (args.enabled) {
    nextDisabled.delete(normalizedName);
  } else {
    nextDisabled.add(normalizedName);
  }

  const nextNames = Array.from(nextDisabled).sort((left, right) => left.localeCompare(right));
  const changed = !arrayShallowEqual(existingNames, nextNames);

  if (changed) {
    await args.settings.set(SKILL_ACTIVATION_STORAGE_KEY, buildStoredDocument(nextNames));
  }

  return {
    name: normalizedName,
    enabled: args.enabled,
    changed,
  };
}
