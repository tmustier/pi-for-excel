/**
 * Experimental feature flags.
 *
 * These flags are local-only toggles for in-progress capabilities and rollout
 * controls. Most are opt-in; some can be default-on with a persisted override.
 */

import { ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY } from "../commands/extension-source-policy.js";
import { dispatchExperimentalFeatureChanged } from "./events.js";

export type ExperimentalFeatureId =
  | "tmux_bridge"
  | "python_bridge"
  | "mcp_tools"
  | "files_workspace"
  | "external_skills_discovery"
  | "remote_extension_urls"
  | "extension_permission_gates"
  | "extension_sandbox_runtime"
  | "extension_widget_v2"
  | "office_js_execute";

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
  /** Used when no explicit value has been persisted yet. */
  defaultEnabled?: boolean;
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
    description: "Allow assistant write/delete in files workspace (list/read + built-in docs stay available).",
    wiring: "wired",
    storageKey: "pi.experimental.filesWorkspace",
  },
  {
    id: "external_skills_discovery",
    slug: "external-skills-discovery",
    aliases: ["external-skills", "skills-discovery"],
    title: "External skills discovery",
    description: "Allow loading Agent Skills from local externally configured SKILL.md sources.",
    warning: "Security: external skills can contain untrusted instructions and executable references.",
    wiring: "flag-only",
    storageKey: "pi.experimental.externalSkillsDiscovery",
  },
  {
    id: "office_js_execute",
    slug: "office-js-execute",
    aliases: ["execute-office-js", "office-js", "excel-js"],
    title: "Direct Office.js execution",
    description: "Allow direct execute_office_js calls with Excel.RequestContext.",
    warning: "High impact: arbitrary Office.js can mutate workbook structure, formulas, and formatting.",
    wiring: "wired",
    storageKey: "pi.experimental.officeJsExecute",
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
  {
    id: "extension_sandbox_runtime",
    slug: "extension-sandbox-rollback",
    aliases: ["extension-sandbox", "extensions-sandbox", "sandboxed-extensions", "extension-host-fallback"],
    title: "Extension sandbox rollback",
    description: "Temporarily route untrusted extensions back to host runtime (kill switch).",
    warning: "Use only as a rollback path. Default behavior runs untrusted extensions in sandbox iframes.",
    wiring: "wired",
    storageKey: "pi.experimental.extensionSandboxHostFallback",
  },
  {
    id: "extension_widget_v2",
    slug: "extension-widget-v2",
    aliases: ["extensions-widget-v2", "widget-v2", "extension-widgets"],
    title: "Extension widget API v2",
    description: "Enable additive multi-widget lifecycle APIs (upsert/remove/clear) with deterministic placement.",
    wiring: "wired",
    storageKey: "pi.experimental.extensionWidgetV2",
  },
] as const satisfies readonly ExperimentalFeatureDefinition[];

export interface ExperimentalFeatureSnapshot extends ExperimentalFeatureDefinition {
  enabled: boolean;
}

function normalizeFeatureToken(input: string): string {
  return input.trim().toLowerCase().replace(/[\s_]+/g, "-");
}

function parseStoredBoolean(raw: string | null): boolean | null {
  if (raw === null) {
    return null;
  }

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
  const stored = parseStoredBoolean(safeGetItem(feature.storageKey));

  if (stored !== null) {
    return stored;
  }

  return feature.defaultEnabled ?? false;
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
