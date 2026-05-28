/**
 * Experimental feature flags.
 *
 * These flags are local-only toggles for in-progress capabilities and rollout
 * controls. Most are opt-in; some can be default-on with a persisted override.
 */

import { t } from "../language/index.js";
import { ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY } from "../commands/extension-source-policy.js";
import { dispatchExperimentalFeatureChanged } from "./events.js";

export type ExperimentalFeatureId =
  | "ui_dark_mode"
  | "remote_extension_urls"
  | "extension_permission_gates"
  | "extension_sandbox_runtime"
  | "extension_widget_v2";

export type ExperimentalFeatureWiring = "wired" | "flag-only";

export interface ExperimentalFeatureDefinition {
  id: ExperimentalFeatureId;
  /** Slash-command token, e.g. `/experimental on extension-permissions` */
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

function getFeatures(): readonly ExperimentalFeatureDefinition[] {
  return [
    {
      id: "ui_dark_mode",
      slug: "dark-mode",
      aliases: ["theme-dark", "ui-dark-mode"],
      title: t("experimental.dark_mode.title"),
      description: t("experimental.dark_mode.desc"),
      wiring: "wired",
      storageKey: "pi.experimental.uiDarkMode",
    },
    {
      id: "remote_extension_urls",
      slug: "remote-extension-urls",
      aliases: ["remote-extensions", "extensions-urls"],
      title: t("experimental.remote_urls.title"),
      description: t("experimental.remote_urls.desc"),
      warning: t("experimental.remote_urls.warning"),
      wiring: "wired",
      storageKey: ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY,
    },
    {
      id: "extension_permission_gates",
      slug: "extension-permissions",
      aliases: ["extensions-permissions", "extension-capability-gates"],
      title: t("experimental.permission_gates.title"),
      description: t("experimental.permission_gates.desc"),
      wiring: "wired",
      storageKey: "pi.experimental.extensionPermissionGates",
    },
    {
      id: "extension_sandbox_runtime",
      slug: "extension-sandbox-rollback",
      aliases: ["extension-sandbox", "extensions-sandbox", "sandboxed-extensions", "extension-host-fallback"],
      title: t("experimental.sandbox_rollback.title"),
      description: t("experimental.sandbox_rollback.desc"),
      warning: t("experimental.sandbox_rollback.warning"),
      wiring: "wired",
      storageKey: "pi.experimental.extensionSandboxHostFallback",
    },
    {
      id: "extension_widget_v2",
      slug: "extension-widget-v2",
      aliases: ["extensions-widget-v2", "widget-v2", "extension-widgets"],
      title: t("experimental.widget_v2.title"),
      description: t("experimental.widget_v2.desc"),
      wiring: "wired",
      storageKey: "pi.experimental.extensionWidgetV2",
    },
  ];
}

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
  for (const feature of getFeatures()) {
    if (feature.id === featureId) return feature;
  }

  throw new Error(`Unknown experimental feature: ${featureId}`);
}

export function listExperimentalFeatures(): readonly ExperimentalFeatureDefinition[] {
  return getFeatures();
}

export function getExperimentalFeatureSlugs(): string[] {
  return getFeatures().map((feature) => feature.slug);
}

export function resolveExperimentalFeature(input: string): ExperimentalFeatureDefinition | null {
  const token = normalizeFeatureToken(input);
  if (!token) return null;

  for (const feature of getFeatures()) {
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
  return getFeatures().map((feature) => ({
    ...feature,
    enabled: isExperimentalFeatureEnabled(feature.id),
  }));
}
