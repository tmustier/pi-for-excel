/**
 * Persisted extension registry (SettingsStore-backed).
 *
 * We keep this as a versioned document so migrations remain explicit.
 */

import { isRecord } from "../utils/type-guards.js";
import {
  deriveStoredExtensionTrust,
  getDefaultPermissionsForTrust,
  normalizeStoredExtensionPermissions,
  type StoredExtensionPermissions,
  type StoredExtensionTrust,
} from "./permissions.js";

export const EXTENSIONS_REGISTRY_STORAGE_KEY = "extensions.registry.v2";
export const LEGACY_EXTENSIONS_REGISTRY_STORAGE_KEY = "extensions.registry.v1";
const EXTENSIONS_REGISTRY_VERSION = 2;

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
  trust: StoredExtensionTrust;
  permissions: StoredExtensionPermissions;
  createdAt: string;
  updatedAt: string;
}

interface StoredExtensionRegistryDocument {
  version: number;
  items: StoredExtensionEntry[];
}

interface LegacyStoredExtensionEntry {
  id: string;
  name: string;
  enabled: boolean;
  source: StoredExtensionSource;
  createdAt: string;
  updatedAt: string;
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

function normalizeBaseEntry(raw: unknown): {
  id: string;
  name: string;
  enabled: boolean;
  source: StoredExtensionSource;
  createdAt: string;
  updatedAt: string;
} | null {
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

function normalizeTrust(raw: unknown): StoredExtensionTrust | null {
  switch (raw) {
    case "builtin":
    case "local-module":
    case "inline-code":
    case "remote-url":
      return raw;
    default:
      return null;
  }
}

function normalizeEntry(raw: unknown): StoredExtensionEntry | null {
  const base = normalizeBaseEntry(raw);
  if (!base) {
    return null;
  }

  const trust = isRecord(raw)
    ? (normalizeTrust(raw.trust) ?? deriveStoredExtensionTrust(base.id, base.source))
    : deriveStoredExtensionTrust(base.id, base.source);

  const permissions = isRecord(raw)
    ? normalizeStoredExtensionPermissions(raw.permissions, trust)
    : getDefaultPermissionsForTrust(trust);

  return {
    ...base,
    trust,
    permissions,
  };
}

function normalizeLegacyEntry(raw: unknown): LegacyStoredExtensionEntry | null {
  const base = normalizeBaseEntry(raw);
  if (!base) {
    return null;
  }

  return base;
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

function normalizeLegacyItems(raw: unknown): LegacyStoredExtensionEntry[] | null {
  if (!Array.isArray(raw)) return null;

  const byId = new Map<string, LegacyStoredExtensionEntry>();
  for (const item of raw) {
    const normalized = normalizeLegacyEntry(item);
    if (!normalized) {
      continue;
    }

    if (!byId.has(normalized.id)) {
      byId.set(normalized.id, normalized);
    }
  }

  return Array.from(byId.values());
}

function migrateLegacyEntry(entry: LegacyStoredExtensionEntry): StoredExtensionEntry {
  const trust = deriveStoredExtensionTrust(entry.id, entry.source);
  return {
    ...entry,
    trust,
    permissions: getDefaultPermissionsForTrust(trust),
  };
}

function normalizeDocument(raw: unknown): StoredExtensionRegistryDocument | null {
  if (!isRecord(raw)) return null;

  const version = raw.version;
  if (typeof version !== "number") {
    return null;
  }

  if (version >= EXTENSIONS_REGISTRY_VERSION) {
    const items = normalizeItems(raw.items);
    if (!items) {
      return null;
    }

    return {
      version,
      items,
    };
  }

  if (version === 1) {
    const legacyItems = normalizeLegacyItems(raw.items);
    if (!legacyItems) {
      return null;
    }

    return {
      version,
      items: legacyItems.map(migrateLegacyEntry),
    };
  }

  return null;
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
  const source: StoredExtensionSource = {
    kind: "module",
    specifier: BUILTIN_SNAKE_SPECIFIER,
  };

  const trust = deriveStoredExtensionTrust(BUILTIN_SNAKE_EXTENSION_ID, source);

  return [
    {
      id: BUILTIN_SNAKE_EXTENSION_ID,
      name: BUILTIN_SNAKE_EXTENSION_NAME,
      enabled: true,
      source,
      trust,
      permissions: getDefaultPermissionsForTrust(trust),
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
 * Legacy `extensions.registry.v1` data is migrated to `extensions.registry.v2`.
 */
export async function loadStoredExtensions(settings: ExtensionSettingsStore): Promise<StoredExtensionEntry[]> {
  const rawCurrent = await settings.get(EXTENSIONS_REGISTRY_STORAGE_KEY);
  const normalizedCurrent = normalizeDocument(rawCurrent);

  if (normalizedCurrent && normalizedCurrent.version >= EXTENSIONS_REGISTRY_VERSION) {
    return normalizedCurrent.items;
  }

  const rawLegacy = await settings.get(LEGACY_EXTENSIONS_REGISTRY_STORAGE_KEY);
  const normalizedLegacy = normalizeDocument(rawLegacy);
  if (normalizedLegacy) {
    await saveStoredExtensions(settings, normalizedLegacy.items);
    return normalizedLegacy.items;
  }

  const defaults = createDefaultExtensionEntries();
  await saveStoredExtensions(settings, defaults);
  return defaults;
}
