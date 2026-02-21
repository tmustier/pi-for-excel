import type { AgentTool } from "@mariozechner/pi-agent-core";

function normalizeConnectionId(rawValue: string): string {
  const normalized = rawValue.trim().toLowerCase();
  if (normalized.length === 0) {
    throw new Error("Tool connection requirement cannot be empty.");
  }

  return normalized;
}

function normalizeRawRequirementList(rawValue: unknown): string[] {
  if (rawValue === undefined || rawValue === null) {
    return [];
  }

  if (typeof rawValue === "string") {
    return [normalizeConnectionId(rawValue)];
  }

  if (!Array.isArray(rawValue)) {
    throw new Error("requiresConnection must be a string or array of strings.");
  }

  const normalized: string[] = [];
  for (const value of rawValue) {
    if (typeof value !== "string") {
      throw new Error("requiresConnection entries must be strings.");
    }

    normalized.push(normalizeConnectionId(value));
  }

  return normalized;
}

export function getToolRequiredConnectionIds(tool: AgentTool): string[] {
  const rawRequirement: unknown = Reflect.get(tool, "requiresConnection");
  const normalized = normalizeRawRequirementList(rawRequirement);
  return Array.from(new Set(normalized));
}

export function collectRequiredConnectionIds(tools: readonly AgentTool[]): string[] {
  const ids = new Set<string>();

  for (const tool of tools) {
    const requirements = getToolRequiredConnectionIds(tool);
    for (const requirement of requirements) {
      ids.add(requirement);
    }
  }

  return Array.from(ids.values()).sort((left, right) => left.localeCompare(right));
}
