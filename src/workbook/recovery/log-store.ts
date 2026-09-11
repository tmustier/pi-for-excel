function isWorkbookRecoveryLogStorePayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Persistence helpers for workbook recovery snapshots.
 */

import {
  clampRetentionLimit,
  MAX_RECOVERY_ENTRIES,
  RETENTION_LIMIT_SETTING_KEY,
} from "./constants.js";
import type { SettingsWriter } from "../../storage/local/settings-store.js";

export const RECOVERY_SETTING_KEY = "workbook.recovery-snapshots.v1";

function isSettingsStoreLike(value: unknown): value is SettingsWriter {
  if (!isWorkbookRecoveryLogStorePayloadShape(value)) return false;

  return (
    typeof value.get === "function" &&
    typeof value.set === "function"
  );
}

export async function defaultGetSettingsStore(): Promise<SettingsWriter | null> {
  try {
    const storageModule = await import("../../storage/local/app-storage.js");
    const appStorage = storageModule.getAppStorage();
    const settings = isWorkbookRecoveryLogStorePayloadShape(appStorage) ? appStorage.settings : null;
    return isSettingsStoreLike(settings) ? settings : null;
  } catch {
    return null;
  }
}

export async function readPersistedWorkbookRecoveryPayload(
  settings: SettingsWriter | null,
): Promise<unknown> {
  if (!settings) {
    return null;
  }

  return settings.get(RECOVERY_SETTING_KEY);
}

export async function writePersistedWorkbookRecoveryPayload(
  settings: SettingsWriter | null,
  payload: unknown,
): Promise<void> {
  if (!settings) {
    return;
  }

  try {
    await settings.set(RECOVERY_SETTING_KEY, payload);
  } catch {
    // ignore persistence failures
  }
}

// ---------------------------------------------------------------------------
// Retention limit
// ---------------------------------------------------------------------------

export async function readRetentionLimit(): Promise<number> {
  const settings = await defaultGetSettingsStore();
  if (!settings) return MAX_RECOVERY_ENTRIES;

  try {
    const raw = await settings.get(RETENTION_LIMIT_SETTING_KEY);
    return clampRetentionLimit(raw);
  } catch {
    return MAX_RECOVERY_ENTRIES;
  }
}

export async function writeRetentionLimit(limit: number): Promise<void> {
  const settings = await defaultGetSettingsStore();
  if (!settings) return;

  try {
    await settings.set(RETENTION_LIMIT_SETTING_KEY, clampRetentionLimit(limit));
  } catch {
    // ignore persistence failures
  }
}
