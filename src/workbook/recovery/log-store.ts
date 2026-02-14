/**
 * Persistence helpers for workbook recovery snapshots.
 */

import { isRecord } from "../../utils/type-guards.js";
import {
  clampRetentionLimit,
  MAX_RECOVERY_ENTRIES,
  RETENTION_LIMIT_SETTING_KEY,
} from "./constants.js";

export const RECOVERY_SETTING_KEY = "workbook.recovery-snapshots.v1";

export interface SettingsStoreLike {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: unknown): Promise<void>;
}

function isSettingsStoreLike(value: unknown): value is SettingsStoreLike {
  if (!isRecord(value)) return false;

  return (
    typeof value.get === "function" &&
    typeof value.set === "function"
  );
}

export async function defaultGetSettingsStore(): Promise<SettingsStoreLike | null> {
  try {
    const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
    const appStorage = storageModule.getAppStorage();
    const settings = isRecord(appStorage) ? appStorage.settings : null;
    return isSettingsStoreLike(settings) ? settings : null;
  } catch {
    return null;
  }
}

export async function readPersistedWorkbookRecoveryPayload(
  settings: SettingsStoreLike | null,
): Promise<unknown> {
  if (!settings) {
    return null;
  }

  try {
    return await settings.get<unknown>(RECOVERY_SETTING_KEY);
  } catch {
    return null;
  }
}

export async function writePersistedWorkbookRecoveryPayload(
  settings: SettingsStoreLike | null,
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
    const raw = await settings.get<unknown>(RETENTION_LIMIT_SETTING_KEY);
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
