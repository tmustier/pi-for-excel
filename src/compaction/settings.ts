import type { SettingsReader } from "../storage/local/settings-store.js";

export const COMPACTION_ENABLED_SETTING_KEY = "compaction.enabled";

/**
 * Reads the persisted auto-compaction preference.
 *
 * Auto-compaction defaults to enabled when the setting is absent, malformed, or
 * cannot be read. Only persisted boolean values override that safe default.
 */
export async function readAutoCompactionEnabled(
  settings: SettingsReader,
): Promise<boolean> {
  try {
    const value = await settings.get(COMPACTION_ENABLED_SETTING_KEY);
    return typeof value === "boolean" ? value : true;
  } catch {
    return true;
  }
}
