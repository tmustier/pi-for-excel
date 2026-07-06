import type { WorkbookRecoverySnapshot } from "../src/workbook/recovery-log.ts";

export const RECOVERY_SETTING_KEY = "workbook.recovery-snapshots.v1";

export interface InMemorySettingsStore {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: DynamicValue): Promise<void>;
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
