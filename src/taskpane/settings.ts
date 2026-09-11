import {
  DEFAULT_PROXY_URL,
  resolveRuntimeDefaultProxyUrl,
} from "../auth/proxy-validation.js";
import type { SpreadsheetHostKind } from "../host/index.js";
import type { SupportedLanguage } from "../language/index.js";
import type { SettingsReader, SettingsWriter } from "../storage/local/settings-store.js";

export const TASKPANE_LANGUAGE_SETTING_KEY = "language";
export const TASKPANE_PROXY_ENABLED_SETTING_KEY = "proxy.enabled";
export const TASKPANE_PROXY_URL_SETTING_KEY = "proxy.url";

export interface TaskpaneProxySettings {
  enabled: boolean;
  url: string | null;
}

function parseProxyEnabled(value: unknown): boolean {
  // Preserve the legacy persisted contract: any truthy value enabled proxying.
  return Boolean(value);
}

/**
 * Reads the taskpane language preference.
 *
 * English is the documented default when the setting is absent, malformed, or
 * cannot be read. Persisted values retain their existing string format.
 */
export async function readTaskpaneLanguage(
  settings: SettingsReader,
): Promise<SupportedLanguage> {
  try {
    const value = await settings.get(TASKPANE_LANGUAGE_SETTING_KEY);
    return value === "zh-CN" ? "zh-CN" : "en";
  } catch {
    return "en";
  }
}

/**
 * Reads the taskpane proxy preferences.
 *
 * Missing or malformed values use documented defaults. Storage read failures
 * propagate so callers cannot mistake an unknown URL for an absent URL before
 * writing a default. The enabled flag preserves legacy JavaScript truthiness
 * for persisted non-boolean values.
 */
export async function readTaskpaneProxySettings(
  settings: SettingsReader,
): Promise<TaskpaneProxySettings> {
  const [enabled, url] = await Promise.all([
    settings.get(TASKPANE_PROXY_ENABLED_SETTING_KEY),
    settings.get(TASKPANE_PROXY_URL_SETTING_KEY),
  ]);

  return {
    enabled: parseProxyEnabled(enabled),
    url: typeof url === "string" ? url.trim() : null,
  };
}

export async function ensureDefaultProxyUrl(
  settings: SettingsWriter,
  hostKind: SpreadsheetHostKind,
): Promise<void> {
  try {
    const runtimeDefaultProxyUrl = resolveRuntimeDefaultProxyUrl({ hostKind });
    const proxySettings = await readTaskpaneProxySettings(settings);
    const storedProxyUrl = proxySettings.url ?? "";
    if (storedProxyUrl.length > 0 && !(storedProxyUrl === DEFAULT_PROXY_URL && runtimeDefaultProxyUrl !== DEFAULT_PROXY_URL)) {
      return;
    }

    await settings.set(TASKPANE_PROXY_URL_SETTING_KEY, runtimeDefaultProxyUrl);
  } catch {
    // A failed read is unknown state, not permission to replace the stored URL.
  }
}
