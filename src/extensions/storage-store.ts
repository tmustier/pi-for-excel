import { isRecord } from "../utils/type-guards.js";

const EXTENSION_STORAGE_KEY = "extensions.storage.v1";
const EXTENSION_STORAGE_VERSION = 1;
const MAX_EXTENSION_STORAGE_BYTES = 1_000_000;

interface ExtensionStorageDocument {
  version: number;
  items: Record<string, Record<string, unknown>>;
}

export interface ExtensionStorageSettings {
  get(key: string): Promise<unknown>;
  set(key: string, value: unknown): Promise<void>;
}

function normalizeStorageDocument(raw: unknown): ExtensionStorageDocument {
  if (!isRecord(raw) || typeof raw.version !== "number" || !isRecord(raw.items)) {
    return {
      version: EXTENSION_STORAGE_VERSION,
      items: {},
    };
  }

  const items: Record<string, Record<string, unknown>> = {};

  for (const [extensionId, extensionRecord] of Object.entries(raw.items)) {
    if (!isRecord(extensionRecord)) {
      continue;
    }

    items[extensionId] = { ...extensionRecord };
  }

  return {
    version: raw.version,
    items,
  };
}

async function loadStorageDocument(settings: ExtensionStorageSettings): Promise<ExtensionStorageDocument> {
  const raw = await settings.get(EXTENSION_STORAGE_KEY);
  return normalizeStorageDocument(raw);
}

async function saveStorageDocument(
  settings: ExtensionStorageSettings,
  document: ExtensionStorageDocument,
): Promise<void> {
  await settings.set(EXTENSION_STORAGE_KEY, {
    version: EXTENSION_STORAGE_VERSION,
    items: document.items,
  });
}

function normalizeStorageKey(key: string): string {
  const trimmed = key.trim();
  if (trimmed.length === 0) {
    throw new Error("Storage key cannot be empty.");
  }

  return trimmed;
}

function calculateSerializedSize(value: unknown): number {
  return new TextEncoder().encode(JSON.stringify(value)).length;
}

export async function getExtensionStorageValue(
  settings: ExtensionStorageSettings,
  extensionId: string,
  key: string,
): Promise<unknown> {
  const document = await loadStorageDocument(settings);
  const extensionStore = document.items[extensionId] ?? {};
  return extensionStore[normalizeStorageKey(key)];
}

export async function setExtensionStorageValue(
  settings: ExtensionStorageSettings,
  extensionId: string,
  key: string,
  value: unknown,
): Promise<void> {
  const normalizedKey = normalizeStorageKey(key);
  const document = await loadStorageDocument(settings);
  const extensionStore = {
    ...(document.items[extensionId] ?? {}),
    [normalizedKey]: value,
  };

  if (calculateSerializedSize(extensionStore) > MAX_EXTENSION_STORAGE_BYTES) {
    throw new Error("Extension storage quota exceeded (1 MB per extension).");
  }

  document.items[extensionId] = extensionStore;
  await saveStorageDocument(settings, document);
}

export async function deleteExtensionStorageValue(
  settings: ExtensionStorageSettings,
  extensionId: string,
  key: string,
): Promise<void> {
  const normalizedKey = normalizeStorageKey(key);
  const document = await loadStorageDocument(settings);
  const extensionStore = { ...(document.items[extensionId] ?? {}) };

  if (!(normalizedKey in extensionStore)) {
    return;
  }

  delete extensionStore[normalizedKey];

  if (Object.keys(extensionStore).length === 0) {
    delete document.items[extensionId];
  } else {
    document.items[extensionId] = extensionStore;
  }

  await saveStorageDocument(settings, document);
}

export async function listExtensionStorageKeys(
  settings: ExtensionStorageSettings,
  extensionId: string,
): Promise<string[]> {
  const document = await loadStorageDocument(settings);
  const extensionStore = document.items[extensionId] ?? {};
  return Object.keys(extensionStore).sort((left, right) => left.localeCompare(right));
}

export async function clearExtensionStorage(
  settings: ExtensionStorageSettings,
  extensionId: string,
): Promise<void> {
  const document = await loadStorageDocument(settings);
  if (!(extensionId in document.items)) {
    return;
  }

  delete document.items[extensionId];
  await saveStorageDocument(settings, document);
}
