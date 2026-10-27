import "fake-indexeddb/auto";
import { SettingsStore } from "../../src/storage/local/settings-store.ts";
import { IndexedDBStorageBackend } from "../../src/storage/local/indexeddb-storage-backend.ts";
import { getAppStorageConfig } from "../../src/storage/init-app-storage.ts";
import { RecoveryStore } from "../../src/workbook/recovery/log-store.ts";
import type { WorkbookRecoverySnapshot } from "../../src/workbook/recovery-log.ts";

export const RECOVERY_SETTING_KEY = "workbook.recovery-snapshots.v1";

export function createInMemorySettingsStore() {
  const backend = new IndexedDBStorageBackend(getAppStorageConfig(`recovery-test-${crypto.randomUUID()}`));
  const settings = new SettingsStore();
  settings.setBackend(backend);
  return {
    backend,
    recovery: new RecoveryStore(backend),
    get: <T>(key: string): Promise<T | null> => backend.get<T>("settings", key),
    set: (key: string, value: unknown): Promise<void> => settings.set(key, value),
    update: settings.update.bind(settings),
  };
}
export type InMemorySettingsStore = ReturnType<typeof createInMemorySettingsStore>;

export function createRecoveringInMemorySettingsStore() {
  const store = createInMemorySettingsStore();
  let failRead = true;
  const original = store.recovery.migrate.bind(store.recovery);
  store.recovery.migrate = () => {
    if (failRead) {
      failRead = false;
      return Promise.reject(new Error("seeded recovery read failure"));
    }
    return original();
  };
  return { ...store, seed: (key: string, value: unknown) => store.set(key, value) };
}

export function findSnapshotById<T extends { id: string }>(snapshots: T[], id: string): T | null {
  return snapshots.find((snapshot) => snapshot.id === id) ?? null;
}

export async function persistedSnapshots(settings: InMemorySettingsStore): Promise<WorkbookRecoverySnapshot[]> {
  const entries = (await settings.recovery.page()).entries;
  const snapshots = await Promise.all(entries.map((entry) => settings.recovery.read(entry.id)));
  return snapshots.filter((snapshot): snapshot is WorkbookRecoverySnapshot => snapshot !== null);
}

export function withoutUndefined(value: unknown): unknown {
  return JSON.parse(JSON.stringify(value));
}
