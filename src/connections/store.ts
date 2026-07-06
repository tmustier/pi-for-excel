import type { ConnectionSecrets, ConnectionStatus } from "./types.js";

function isConnectionsStorePayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export const CONNECTION_STORE_KEY = "connections.store.v1";
const CONNECTION_STORE_VERSION = 1;

export interface ConnectionSettingsStore {
  get(key: string): Promise<DynamicValue>;
  set(key: string, value: DynamicValue): Promise<void>;
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

function normalizeOptionalString(value: DynamicValue): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function normalizeConnectionStatus(value: DynamicValue): ConnectionStatus | undefined {
  return value === "connected" || value === "missing" || value === "invalid" || value === "error"
    ? value
    : undefined;
}

function normalizeSecrets(value: DynamicValue): ConnectionSecrets | undefined {
  if (!isConnectionsStorePayloadShape(value)) return undefined;

  const next: ConnectionSecrets = {};
  for (const [key, raw] of Object.entries(value)) {
    if (typeof raw !== "string") continue;
    next[key] = raw;
  }

  return Object.keys(next).length > 0 ? next : undefined;
}

function normalizeConnectionRecord(value: DynamicValue): StoredConnectionRecord | null {
  if (!isConnectionsStorePayloadShape(value)) return null;

  const record: StoredConnectionRecord = {};
  const status = normalizeConnectionStatus(value.status);
  const lastValidatedAt = normalizeOptionalString(value.lastValidatedAt);
  const lastError = normalizeOptionalString(value.lastError);
  const secrets = normalizeSecrets(value.secrets);

  if (status !== undefined) record.status = status;
  if (lastValidatedAt !== undefined) record.lastValidatedAt = lastValidatedAt;
  if (lastError !== undefined) record.lastError = lastError;
  if (secrets !== undefined) record.secrets = secrets;

  return record;
}

function normalizeDocument(value: DynamicValue): ConnectionStoreDocument {
  if (!isConnectionsStorePayloadShape(value) || !isConnectionsStorePayloadShape(value.items)) {
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
