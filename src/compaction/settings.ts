export const COMPACTION_ENABLED_SETTING_KEY = "compaction.enabled";

export interface CompactionSettingsStore {
  get(key: string): Promise<DynamicValue>;
}

/**
 * Reads the persisted auto-compaction preference.
 *
 * Auto-compaction defaults to enabled when the setting is absent, malformed, or
 * cannot be read. Only persisted boolean values override that safe default.
 */
export async function readAutoCompactionEnabled(
  settings: CompactionSettingsStore,
): Promise<boolean> {
  try {
    const value = await settings.get(COMPACTION_ENABLED_SETTING_KEY);
    return typeof value === "boolean" ? value : true;
  } catch {
    return true;
  }
}
