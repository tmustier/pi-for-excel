import { isRecord } from "../utils/type-guards.js";
import type { ConnectionSecrets, ConnectionStatus } from "./types.js";

export const CONNECTION_STORE_KEY = "connections.store.v1";
const CONNECTION_STORE_VERSION = 1;

export interface ConnectionSettingsStore {
  get(key: string): Promise<unknown>;
  set(key: string, value: unknown): Promise<void>;
  delete?(key: string): Promise<void>;
}

export interface StoredConnectionRecord {
  status?: ConnectionStatus;
  lastValidatedAt?: string;
  lastError?: string;
  secrets?: ConnectionSecrets;
}

interface ConnectionStoreDocument {
  version: number;
  items: Record<string, StoredConnectionRecord>;
}

function normalizeOptionalString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function normalizeConnectionStatus(value: unknown): ConnectionStatus | undefined {
  return value === "connected" || value === "missing" || value === "invalid" || value === "error"
    ? value
    : undefined;
}

function normalizeSecrets(value: unknown): ConnectionSecrets | undefined {
  if (!isRecord(value)) return undefined;

  const next: ConnectionSecrets = {};
  for (const [key, raw] of Object.entries(value)) {
    if (typeof raw !== "string") continue;
    next[key] = raw;
  }

  return Object.keys(next).length > 0 ? next : undefined;
}

function normalizeConnectionRecord(value: unknown): StoredConnectionRecord | null {
  if (!isRecord(value)) return null;

  return {
    status: normalizeConnectionStatus(value.status),
    lastValidatedAt: normalizeOptionalString(value.lastValidatedAt),
    lastError: normalizeOptionalString(value.lastError),
    secrets: normalizeSecrets(value.secrets),
  };
}

function normalizeDocument(value: unknown): ConnectionStoreDocument {
  if (!isRecord(value) || !isRecord(value.items)) {
    return {
      version: CONNECTION_STORE_VERSION,
      items: {},
    };
  }

  const items: Record<string, StoredConnectionRecord> = {};
  for (const [connectionId, rawRecord] of Object.entries(value.items)) {
    const normalized = normalizeConnectionRecord(rawRecord);
    if (!normalized) continue;
    items[connectionId] = normalized;
  }

  return {
    version: CONNECTION_STORE_VERSION,
    items,
  };
}

export async function loadConnectionStoreDocument(
  settings: ConnectionSettingsStore,
): Promise<Record<string, StoredConnectionRecord>> {
  const raw = await settings.get(CONNECTION_STORE_KEY);
  const normalized = normalizeDocument(raw);
  return normalized.items;
}

export async function saveConnectionStoreDocument(
  settings: ConnectionSettingsStore,
  items: Record<string, StoredConnectionRecord>,
): Promise<void> {
  await settings.set(CONNECTION_STORE_KEY, {
    version: CONNECTION_STORE_VERSION,
    items,
  });
}
