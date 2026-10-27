import type { WorkbookRecoverySnapshot } from "../../src/workbook/recovery-log.ts";

export const RECOVERY_SETTING_KEY = "workbook.recovery-snapshots.v1";

export interface InMemorySettingsStore {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: unknown): Promise<void>;
  update<T, U>(key: string, parse: (raw: unknown) => T, change: (current: T) => U | undefined): Promise<void>;
}

function updateIn(values: Map<string, unknown>) {
  return <T, U>(key: string, parse: (raw: unknown) => T, change: (current: T) => U | undefined): Promise<void> => {
    const next = change(parse(values.get(key) ?? null));
    if (next === undefined) values.delete(key);
    else values.set(key, next);
    return Promise.resolve();
  };
}

export interface RecoveringInMemorySettingsStore extends InMemorySettingsStore {
  seed(key: string, value: unknown): void;
}

export function createRecoveringInMemorySettingsStore(): RecoveringInMemorySettingsStore {
  const values = new Map<string, unknown>();
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
    set: (key: string, value: unknown): Promise<void> => {
      values.set(key, value);
      return Promise.resolve();
    },
    // A read-modify-write reads first, so it fails the same seeded way.
    update: <T, U>(key: string, parse: (raw: unknown) => T, change: (current: T) => U | undefined): Promise<void> => {
      if (key === RECOVERY_SETTING_KEY && failRecoveryRead) {
        failRecoveryRead = false;
        return Promise.reject(new Error("seeded recovery read failure"));
      }
      return updateIn(values)(key, parse, change);
    },
    seed: (key: string, value: unknown): void => {
      values.set(key, value);
    },
  };
}

export function createInMemorySettingsStore(): InMemorySettingsStore {
  const values = new Map<string, unknown>();

  return {
    get: <T>(key: string): Promise<T | null> => {
      const value = values.get(key);
      return Promise.resolve(value === undefined ? null : value as T);
    },
    set: (key: string, value: unknown): Promise<void> => {
      values.set(key, value);
      return Promise.resolve();
    },
    update: updateIn(values),
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

export function withoutUndefined(value: unknown): unknown {
  return JSON.parse(JSON.stringify(value));
}
