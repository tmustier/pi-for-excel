/**
 * Shared app storage initialization for taskpane + dialog.
 */

import { AppStorage, setAppStorage } from "./local/app-storage.js";
import { IndexedDBStorageBackend } from "./local/indexeddb-storage-backend.js";
import { CustomProvidersStore } from "./local/custom-providers-store.js";
import { ModelCatalogsStore } from "./local/model-catalogs-store.js";
import { ProviderKeysStore } from "./local/provider-keys-store.js";
import { SessionsStore } from "./local/sessions-store.js";
import { SettingsStore } from "./local/settings-store.js";
import type { IndexedDBConfig, StorageBackend } from "./local/types.js";

export const APP_STORAGE_DATABASE_NAME = "pi-for-excel";
export const APP_STORAGE_DATABASE_VERSION = 2;

export function getAppStorageConfig(
  dbName = APP_STORAGE_DATABASE_NAME,
): IndexedDBConfig {
  const settings = new SettingsStore();
  const providerKeys = new ProviderKeysStore();
  const sessions = new SessionsStore();
  const customProviders = new CustomProvidersStore();
  const modelCatalogs = new ModelCatalogsStore();

  return {
    dbName,
    version: APP_STORAGE_DATABASE_VERSION,
    stores: [
      settings.getConfig(),
      providerKeys.getConfig(),
      sessions.getConfig(),
      SessionsStore.getMetadataConfig(),
      customProviders.getConfig(),
      modelCatalogs.getConfig(),
    ],
  };
}

type InitializedAppStorage = {
  storage: AppStorage;
  settings: SettingsStore;
  providerKeys: ProviderKeysStore;
  sessions: SessionsStore;
  customProviders: CustomProvidersStore;
  modelCatalogs: ModelCatalogsStore;
  backend: StorageBackend;
};

export function initAppStorage(
  dbName = APP_STORAGE_DATABASE_NAME,
  suppliedBackend?: StorageBackend,
): InitializedAppStorage {
  const settings = new SettingsStore();
  const providerKeys = new ProviderKeysStore();
  const sessions = new SessionsStore();
  const customProviders = new CustomProvidersStore();
  const modelCatalogs = new ModelCatalogsStore();

  const backend = suppliedBackend ?? new IndexedDBStorageBackend(getAppStorageConfig(dbName));

  settings.setBackend(backend);
  providerKeys.setBackend(backend);
  sessions.setBackend(backend);
  customProviders.setBackend(backend);
  modelCatalogs.setBackend(backend);

  const storage = new AppStorage(
    settings,
    providerKeys,
    sessions,
    customProviders,
    modelCatalogs,
    backend,
  );
  setAppStorage(storage);

  return { storage, settings, providerKeys, sessions, customProviders, modelCatalogs, backend };
}
