/**
 * Persisted extension registry (SettingsStore-backed).
 *
 * We keep this as a versioned document so future migrations stay explicit.
 */

import { isRecord } from "../utils/type-guards.js";

export const EXTENSIONS_REGISTRY_STORAGE_KEY = "extensions.registry.v1";
const EXTENSIONS_REGISTRY_VERSION = 1;

export interface ExtensionSettingsStore {
  get(key: string): Promise<unknown>;
  set(key: string, value: unknown): Promise<void>;
}

export const BUILTIN_SNAKE_EXTENSION_ID = "builtin.snake";
const BUILTIN_SNAKE_EXTENSION_NAME = "Snake";
const BUILTIN_SNAKE_SPECIFIER = "../extensions/snake.js";

export type StoredExtensionSource =
  | { kind: "module"; specifier: string }
  | { kind: "inline"; code: string };

export interface StoredExtensionEntry {
  id: string;
  name: string;
  enabled: boolean;
  source: StoredExtensionSource;
  createdAt: string;
  updatedAt: string;
}

interface StoredExtensionRegistryDocument {
  version: number;
  items: StoredExtensionEntry[];
}

function isValidIsoTimestamp(value: string): boolean {
  return !Number.isNaN(Date.parse(value));
}

function normalizeNonEmptyString(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function normalizeSource(raw: unknown): StoredExtensionSource | null {
  if (!isRecord(raw)) return null;

  const kind = raw.kind;
  if (kind === "module") {
    const specifier = normalizeNonEmptyString(raw.specifier);
    if (!specifier) return null;
    return { kind: "module", specifier };
  }

  if (kind === "inline") {
    const code = typeof raw.code === "string" ? raw.code : null;
    if (code === null) return null;
    return { kind: "inline", code };
  }

  return null;
}

function normalizeEntry(raw: unknown): StoredExtensionEntry | null {
  if (!isRecord(raw)) return null;

  const id = normalizeNonEmptyString(raw.id);
  const name = normalizeNonEmptyString(raw.name);
  const enabled = raw.enabled;
  const source = normalizeSource(raw.source);
  const createdAtRaw = typeof raw.createdAt === "string" ? raw.createdAt : "";
  const updatedAtRaw = typeof raw.updatedAt === "string" ? raw.updatedAt : "";

  if (!id || !name || typeof enabled !== "boolean" || !source) {
    return null;
  }

  const createdAt = isValidIsoTimestamp(createdAtRaw) ? createdAtRaw : new Date().toISOString();
  const updatedAt = isValidIsoTimestamp(updatedAtRaw) ? updatedAtRaw : createdAt;

  return {
    id,
    name,
    enabled,
    source,
    createdAt,
    updatedAt,
  };
}

function normalizeItems(raw: unknown): StoredExtensionEntry[] | null {
  if (!Array.isArray(raw)) return null;

  const byId = new Map<string, StoredExtensionEntry>();
  for (const item of raw) {
    const normalized = normalizeEntry(item);
    if (!normalized) {
      continue;
    }

    if (!byId.has(normalized.id)) {
      byId.set(normalized.id, normalized);
    }
  }

  return Array.from(byId.values());
}

function normalizeDocument(raw: unknown): StoredExtensionRegistryDocument | null {
  if (!isRecord(raw)) return null;

  const version = raw.version;
  const items = normalizeItems(raw.items);
  if (typeof version !== "number" || !items) {
    return null;
  }

  return {
    version,
    items,
  };
}

function createRegistryDocument(items: StoredExtensionEntry[]): StoredExtensionRegistryDocument {
  return {
    version: EXTENSIONS_REGISTRY_VERSION,
    items,
  };
}

export function createDefaultExtensionEntries(
  timestamp: string = new Date().toISOString(),
): StoredExtensionEntry[] {
  return [
    {
      id: BUILTIN_SNAKE_EXTENSION_ID,
      name: BUILTIN_SNAKE_EXTENSION_NAME,
      enabled: true,
      source: {
        kind: "module",
        specifier: BUILTIN_SNAKE_SPECIFIER,
      },
      createdAt: timestamp,
      updatedAt: timestamp,
    },
  ];
}

export async function saveStoredExtensions(
  settings: ExtensionSettingsStore,
  items: StoredExtensionEntry[],
): Promise<void> {
  await settings.set(EXTENSIONS_REGISTRY_STORAGE_KEY, createRegistryDocument(items));
}

/**
 * Load stored extensions from SettingsStore.
 *
 * If nothing is stored (or stored data is invalid), we seed defaults (Snake).
 */
export async function loadStoredExtensions(settings: ExtensionSettingsStore): Promise<StoredExtensionEntry[]> {
  const raw = await settings.get(EXTENSIONS_REGISTRY_STORAGE_KEY);
  const normalized = normalizeDocument(raw);

  if (normalized) {
    return normalized.items;
  }

  const defaults = createDefaultExtensionEntries();
  await saveStoredExtensions(settings, defaults);
  return defaults;
}
