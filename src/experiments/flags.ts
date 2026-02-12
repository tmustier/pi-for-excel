/**
 * Experimental feature flags.
 *
 * These flags are opt-in, local-only toggles intended for power users and
 * in-progress capabilities. Features can read these synchronously at runtime.
 */

import { ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY } from "../commands/extension-source-policy.js";
import { dispatchExperimentalFeatureChanged } from "./events.js";

export type ExperimentalFeatureId =
  | "tmux_bridge"
  | "python_bridge"
  | "mcp_tools"
  | "files_workspace"
  | "remote_extension_urls"
  | "extension_permission_gates";

export type ExperimentalFeatureWiring = "wired" | "flag-only";

export interface ExperimentalFeatureDefinition {
  id: ExperimentalFeatureId;
  /** Slash-command token, e.g. `/experimental on tmux-bridge` */
  slug: string;
  /** Alternate tokens accepted by the resolver. */
  aliases: readonly string[];
  title: string;
  description: string;
  warning?: string;
  wiring: ExperimentalFeatureWiring;
  storageKey: string;
}

const EXPERIMENTAL_FEATURES = [
  {
    id: "tmux_bridge",
    slug: "tmux-bridge",
    aliases: ["tmux", "tmux-local-bridge"],
    title: "Tmux local bridge",
    description: "Allow local tmux bridge integration for interactive shell sessions.",
    wiring: "wired",
    storageKey: "pi.experimental.tmuxBridge",
  },
  {
    id: "python_bridge",
    slug: "python-bridge",
    aliases: ["python", "libreoffice-bridge"],
    title: "Python / LibreOffice bridge",
    description: "Allow local Python and LibreOffice bridge integrations.",
    wiring: "wired",
    storageKey: "pi.experimental.pythonBridge",
  },
  {
    id: "mcp_tools",
    slug: "mcp-tools",
    aliases: ["mcp", "external-tools"],
    title: "MCP / external tools",
    description: "Allow experimental external tool integration (MCP, web tools).",
    wiring: "flag-only",
    storageKey: "pi.experimental.mcpTools",
  },
  {
    id: "files_workspace",
    slug: "files-workspace",
    aliases: ["files", "workspace-files", "artifacts"],
    title: "Files workspace",
    description: "Allow experimental file workspace and artifact tooling.",
    wiring: "flag-only",
    storageKey: "pi.experimental.filesWorkspace",
  },
  {
    id: "remote_extension_urls",
    slug: "remote-extension-urls",
    aliases: ["remote-extensions", "extensions-urls"],
    title: "Remote extension URLs",
    description: "Allow loading extensions from remote http(s) URLs.",
    warning: "Unsafe: remote extension code can read workbook data and credentials.",
    wiring: "wired",
    storageKey: ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY,
  },
  {
    id: "extension_permission_gates",
    slug: "extension-permissions",
    aliases: ["extensions-permissions", "extension-capability-gates"],
    title: "Extension permission gates",
    description: "Enforce per-extension capability permissions when extensions activate.",
    wiring: "wired",
    storageKey: "pi.experimental.extensionPermissionGates",
  },
] as const satisfies readonly ExperimentalFeatureDefinition[];

export interface ExperimentalFeatureSnapshot extends ExperimentalFeatureDefinition {
  enabled: boolean;
}

function normalizeFeatureToken(input: string): string {
  return input.trim().toLowerCase().replace(/[\s_]+/g, "-");
}

function parseStoredBoolean(raw: string | null): boolean {
  return raw === "1" || raw === "true";
}

function formatStoredBoolean(value: boolean): string {
  return value ? "1" : "0";
}

function safeGetItem(key: string): string | null {
  try {
    if (typeof localStorage === "undefined") return null;
    return localStorage.getItem(key);
  } catch {
    return null;
  }
}

function safeSetItem(key: string, value: string): void {
  try {
    if (typeof localStorage === "undefined") return;
    localStorage.setItem(key, value);
  } catch {
    // ignore (private mode / disabled storage)
  }
}

function getFeatureDefinition(featureId: ExperimentalFeatureId): ExperimentalFeatureDefinition {
  for (const feature of EXPERIMENTAL_FEATURES) {
    if (feature.id === featureId) return feature;
  }

  throw new Error(`Unknown experimental feature: ${featureId}`);
}

export function listExperimentalFeatures(): readonly ExperimentalFeatureDefinition[] {
  return EXPERIMENTAL_FEATURES;
}

export function getExperimentalFeatureSlugs(): string[] {
  return EXPERIMENTAL_FEATURES.map((feature) => feature.slug);
}

export function resolveExperimentalFeature(input: string): ExperimentalFeatureDefinition | null {
  const token = normalizeFeatureToken(input);
  if (!token) return null;

  for (const feature of EXPERIMENTAL_FEATURES) {
    if (token === normalizeFeatureToken(feature.slug)) {
      return feature;
    }

    for (const alias of feature.aliases) {
      if (token === normalizeFeatureToken(alias)) {
        return feature;
      }
    }
  }

  return null;
}

export function isExperimentalFeatureEnabled(featureId: ExperimentalFeatureId): boolean {
  const feature = getFeatureDefinition(featureId);
  return parseStoredBoolean(safeGetItem(feature.storageKey));
}

export function setExperimentalFeatureEnabled(
  featureId: ExperimentalFeatureId,
  enabled: boolean,
): void {
  const feature = getFeatureDefinition(featureId);
  const previous = isExperimentalFeatureEnabled(featureId);

  safeSetItem(feature.storageKey, formatStoredBoolean(enabled));

  if (previous === enabled) return;

  dispatchExperimentalFeatureChanged({
    featureId,
    enabled,
  });
}

export function toggleExperimentalFeature(featureId: ExperimentalFeatureId): boolean {
  const next = !isExperimentalFeatureEnabled(featureId);
  setExperimentalFeatureEnabled(featureId, next);
  return next;
}

export function getExperimentalFeatureSnapshots(): ExperimentalFeatureSnapshot[] {
  return EXPERIMENTAL_FEATURES.map((feature) => ({
    ...feature,
    enabled: isExperimentalFeatureEnabled(feature.id),
  }));
}
