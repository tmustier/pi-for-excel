import type { AgentTool } from "@earendil-works/pi-agent-core";

export interface ToolConnectionMetadata {
  readonly requiresConnection?: string | readonly string[];
}

export type ConnectionAwareAgentTool = AgentTool & ToolConnectionMetadata;

interface ToolConnectionMetadataBoundary {
  requiresConnection?: DynamicValue;
}

function normalizeConnectionId(rawValue: string): string {
  const normalized = rawValue.trim().toLowerCase();
  if (normalized.length === 0) {
    throw new Error("Tool connection requirement cannot be empty.");
  }

  return normalized;
}

function parseToolConnectionMetadata(tool: AgentTool): ToolConnectionMetadata {
  // AgentTool is extensible at the extension boundary; this parser validates its optional metadata.
  const boundary = tool as AgentTool & ToolConnectionMetadataBoundary;
  const rawValue = boundary.requiresConnection;

  if (rawValue === undefined || rawValue === null) {
    return {};
  }

  if (typeof rawValue === "string") {
    return { requiresConnection: normalizeConnectionId(rawValue) };
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

  return { requiresConnection: normalized };
}

export function getToolRequiredConnectionIds(tool: AgentTool): string[] {
  const metadata = parseToolConnectionMetadata(tool);
  const requirement = metadata.requiresConnection;
  const normalized = typeof requirement === "string" ? [requirement] : requirement ?? [];
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
