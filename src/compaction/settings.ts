import type { SettingsReader } from "../storage/local/settings-store.js";

export const COMPACTION_ENABLED_SETTING_KEY = "compaction.enabled";

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
