import type { ConnectionSecrets, ConnectionStatus } from "./types.js";
import type { SettingsReader, SettingsWriter } from "../storage/local/settings-store.js";

function isConnectionsStorePayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export const CONNECTION_STORE_KEY = "connections.store.v1";
const CONNECTION_STORE_VERSION = 1;

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
  if (!isConnectionsStorePayloadShape(value)) return undefined;

  const next: ConnectionSecrets = {};
  for (const [key, raw] of Object.entries(value)) {
    if (typeof raw !== "string") continue;
    next[key] = raw;
  }

  return Object.keys(next).length > 0 ? next : undefined;
}

function normalizeConnectionRecord(value: unknown): StoredConnectionRecord | null {
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

function normalizeDocument(value: unknown): ConnectionStoreDocument {
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
  settings: SettingsReader,
): Promise<Record<string, StoredConnectionRecord>> {
  try {
    const raw = await settings.get(CONNECTION_STORE_KEY);
    return normalizeDocument(raw).items;
  } catch {
    return {};
  }
}

/**
 * Reads the connection document before a mutation. Unlike the display reader,
 * storage failures propagate so a write cannot replace unread sibling records.
 */
export async function loadConnectionStoreDocumentForUpdate(
  settings: SettingsWriter,
): Promise<Record<string, StoredConnectionRecord>> {
  const raw = await settings.get(CONNECTION_STORE_KEY);
  return normalizeDocument(raw).items;
}

export async function saveConnectionStoreDocument(
  settings: SettingsWriter,
  items: Record<string, StoredConnectionRecord>,
): Promise<void> {
  await settings.set(CONNECTION_STORE_KEY, {
    version: CONNECTION_STORE_VERSION,
    items,
  });
}
