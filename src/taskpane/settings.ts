import type { SupportedLanguage } from "../language/index.js";

export const TASKPANE_LANGUAGE_SETTING_KEY = "language";
export const TASKPANE_PROXY_ENABLED_SETTING_KEY = "proxy.enabled";
export const TASKPANE_PROXY_URL_SETTING_KEY = "proxy.url";

export interface TaskpaneSettingsStore {
  get(key: string): Promise<DynamicValue>;
}

export interface TaskpaneProxySettings {
  enabled: boolean;
  url: string | null;
}

function parseProxyEnabled(value: DynamicValue): boolean {
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
  settings: TaskpaneSettingsStore,
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
 * Proxying defaults to disabled with no configured URL when either setting
 * cannot be read. The enabled flag preserves legacy JavaScript truthiness for
 * persisted non-boolean values.
 */
export async function readTaskpaneProxySettings(
  settings: TaskpaneSettingsStore,
): Promise<TaskpaneProxySettings> {
  try {
    const [enabled, url] = await Promise.all([
      settings.get(TASKPANE_PROXY_ENABLED_SETTING_KEY),
      settings.get(TASKPANE_PROXY_URL_SETTING_KEY),
    ]);

    return {
      enabled: parseProxyEnabled(enabled),
      url: typeof url === "string" ? url.trim() : null,
    };
  } catch {
    return { enabled: false, url: null };
  }
}
