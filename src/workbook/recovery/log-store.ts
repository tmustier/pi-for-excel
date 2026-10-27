/**
 * Persistence helpers for workbook recovery snapshots.
 *
 * Every taskpane on this computer shares one stored list, so writes go
 * through `SettingsUpdater.update` (an atomic read-modify-write of what is
 * stored now) rather than writing back a list loaded earlier.
 */

import {
  clampRetentionLimit,
  MAX_RECOVERY_ENTRIES,
  RETENTION_LIMIT_SETTING_KEY,
} from "./constants.js";
import type { SettingsReader, SettingsUpdater } from "../../storage/local/settings-store.js";

export const RECOVERY_SETTING_KEY = "workbook.recovery-snapshots.v1";

export async function defaultGetSettingsStore(): Promise<SettingsUpdater | null> {
  try {
    const storageModule = await import("../../storage/local/app-storage.js");
    return storageModule.getAppStorage().settings;
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// Retention limit
// ---------------------------------------------------------------------------

/** "Keep at most N backups" per workbook; the default is the maximum. */
export async function readRetentionLimit(
  settings: SettingsReader | null = null,
): Promise<number> {
  const store = settings ?? await defaultGetSettingsStore();
  if (!store) return MAX_RECOVERY_ENTRIES;

  try {
    const raw = await store.get(RETENTION_LIMIT_SETTING_KEY);
    return clampRetentionLimit(raw);
  } catch {
    return MAX_RECOVERY_ENTRIES;
  }
}

export async function writeRetentionLimit(limit: number): Promise<void> {
  const settings = await defaultGetSettingsStore();
  if (!settings) {
    throw new Error("Settings storage is unavailable.");
  }

  await settings.set(RETENTION_LIMIT_SETTING_KEY, clampRetentionLimit(limit));
}
