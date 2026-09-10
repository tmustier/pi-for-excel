import type { WorkbookRecoverySnapshot } from "../../src/workbook/recovery-log.ts";

export const RECOVERY_SETTING_KEY = "workbook.recovery-snapshots.v1";

export interface InMemorySettingsStore {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: DynamicValue): Promise<void>;
}

export interface RecoveringInMemorySettingsStore extends InMemorySettingsStore {
  seed(key: string, value: DynamicValue): void;
}

export function createRecoveringInMemorySettingsStore(): RecoveringInMemorySettingsStore {
  const values = new Map<string, DynamicValue>();
  let failRecoveryRead = true;

  return {
    get: <T>(key: string): Promise<T | null> => {
      if (key === RECOVERY_SETTING_KEY && failRecoveryRead) {
        failRecoveryRead = false;
        return Promise.reject(new Error("seeded recovery read failure"));
      }
      const value = values.get(key);
      return Promise.resolve(value === undefined ? null : value as T);
    },
    set: (key: string, value: DynamicValue): Promise<void> => {
      values.set(key, value);
      return Promise.resolve();
    },
    seed: (key: string, value: DynamicValue): void => {
      values.set(key, value);
    },
  };
}

export function createInMemorySettingsStore(): InMemorySettingsStore {
  const values = new Map<string, DynamicValue>();

  return {
    get: <T>(key: string): Promise<T | null> => {
      const value = values.get(key);
      return Promise.resolve(value === undefined ? null : value as T);
    },
    set: (key: string, value: DynamicValue): Promise<void> => {
      values.set(key, value);
      return Promise.resolve();
    },
  };
}

export function findSnapshotById(snapshots: WorkbookRecoverySnapshot[], id: string): WorkbookRecoverySnapshot | null {
  for (const snapshot of snapshots) {
    if (snapshot.id === id) {
      return snapshot;
    }
  }

  return null;
}

export function withoutUndefined(value: DynamicValue): DynamicValue {
  return JSON.parse(JSON.stringify(value));
}
