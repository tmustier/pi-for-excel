import {
  DEFAULT_PROXY_URL,
  isLoopbackProxyUrl,
  resolveRuntimeDefaultProxyUrl,
  validateOfficeProxyUrl,
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
  return Boolean(value);
}

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
  }
}

export async function saveTaskpaneProxyUrl(settings: SettingsWriter, rawUrl: string): Promise<string> {
  const trimmed = rawUrl.trim();
  const normalized = validateOfficeProxyUrl(trimmed.length > 0 ? trimmed : DEFAULT_PROXY_URL);
  await settings.set(TASKPANE_PROXY_URL_SETTING_KEY, normalized);
  return normalized;
}

export function needsRemoteProxyWarning(proxy: TaskpaneProxySettings): boolean {
  return proxy.enabled && proxy.url !== null && proxy.url.length > 0 && !isLoopbackProxyUrl(proxy.url);
}
