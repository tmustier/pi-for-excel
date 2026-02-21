import {
  loadConnectionStoreDocument,
  saveConnectionStoreDocument,
  type ConnectionSettingsStore,
  type StoredConnectionRecord,
} from "./store.js";
import type {
  ConnectionDefinition,
  ConnectionPromptEntry,
  ConnectionRuntimeAuthFailure,
  ConnectionSnapshot,
  ConnectionState,
  ConnectionStatus,
} from "./types.js";

interface RegisteredConnectionDefinition extends ConnectionDefinition {
  ownerId: string;
}

type ConnectionManagerListener = () => void;

const CONNECTION_ID_PATTERN = /^[a-z0-9][a-z0-9._-]{1,95}$/;

const ALL_CONNECTION_STATUSES: readonly ConnectionStatus[] = [
  "connected",
  "missing",
  "invalid",
  "error",
];

const STATUS_TRANSITIONS: Readonly<Record<ConnectionStatus, readonly ConnectionStatus[]>> = {
  connected: ALL_CONNECTION_STATUSES,
  missing: ALL_CONNECTION_STATUSES,
  invalid: ALL_CONNECTION_STATUSES,
  error: ALL_CONNECTION_STATUSES,
};

function normalizeNonEmpty(input: string, field: string): string {
  const trimmed = input.trim();
  if (trimmed.length === 0) {
    throw new Error(`${field} cannot be empty.`);
  }

  return trimmed;
}

function normalizeConnectionId(input: string): string {
  const normalized = normalizeNonEmpty(input, "connection id").toLowerCase();
  if (!CONNECTION_ID_PATTERN.test(normalized)) {
    throw new Error(
      "Connection id must start with a letter/number and use only lowercase letters, numbers, dot, underscore, or dash.",
    );
  }

  return normalized;
}

function normalizeOwnerId(ownerId: string): string {
  return normalizeNonEmpty(ownerId, "owner id");
}

function normalizeSecrets(input: Record<string, string>): Record<string, string> {
  const normalized: Record<string, string> = {};

  for (const [fieldId, rawValue] of Object.entries(input)) {
    const normalizedFieldId = normalizeNonEmpty(fieldId, "secret field id");
    const value = rawValue.trim();
    if (value.length === 0) continue;
    normalized[normalizedFieldId] = value;
  }

  return normalized;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function redactSecretsInMessage(message: string, secrets: Record<string, string> | undefined): string {
  if (!secrets) return message;

  let sanitized = message;
  const uniqueValues = new Set<string>();

  for (const value of Object.values(secrets)) {
    const trimmed = value.trim();
    if (trimmed.length < 4) continue;
    uniqueValues.add(trimmed);
  }

  for (const secret of uniqueValues) {
    sanitized = sanitized.replace(new RegExp(escapeRegExp(secret), "g"), "••••");
  }

  return sanitized;
}

function buildDefaultSetupHint(title: string): string {
  return `Open /tools → Connections → ${title}`;
}

function assertStatusTransition(from: ConnectionStatus, to: ConnectionStatus): void {
  const allowed = STATUS_TRANSITIONS[from];
  if (!allowed.includes(to)) {
    throw new Error(`Invalid connection status transition: ${from} → ${to}`);
  }
}

function hasRequiredSecrets(definition: ConnectionDefinition, secrets: Record<string, string> | undefined): boolean {
  const currentSecrets = secrets ?? {};

  for (const field of definition.secretFields) {
    if (!field.required) continue;

    const value = currentSecrets[field.id];
    if (typeof value !== "string" || value.trim().length === 0) {
      return false;
    }
  }

  return true;
}

function resolveEffectiveStatus(
  definition: ConnectionDefinition,
  record: StoredConnectionRecord | undefined,
): ConnectionStatus {
  const requiredPresent = hasRequiredSecrets(definition, record?.secrets);
  if (!requiredPresent) {
    return "missing";
  }

  if (record?.status === "invalid") return "invalid";
  if (record?.status === "error") return "error";
  return "connected";
}

function normalizeConnectionDefinition(definition: ConnectionDefinition): ConnectionDefinition {
  const id = normalizeConnectionId(definition.id);
  const title = normalizeNonEmpty(definition.title, "connection title");
  const capability = normalizeNonEmpty(definition.capability, "connection capability");

  const secretFields = definition.secretFields.map((field) => {
    return {
      id: normalizeNonEmpty(field.id, "secret field id"),
      label: normalizeNonEmpty(field.label, "secret field label"),
      required: field.required,
      maskInUi: field.maskInUi,
    };
  });

  return {
    ...definition,
    id,
    title,
    capability,
    secretFields,
    setupHint: definition.setupHint?.trim().length
      ? definition.setupHint.trim()
      : undefined,
  };
}

async function mutateConnectionRecord(args: {
  settings: ConnectionSettingsStore;
  connectionId: string;
  mutator: (record: StoredConnectionRecord | undefined) => StoredConnectionRecord | null;
}): Promise<boolean> {
  const items = await loadConnectionStoreDocument(args.settings);
  const previous = items[args.connectionId];
  const next = args.mutator(previous);

  if (!next) {
    if (!(args.connectionId in items)) {
      return false;
    }

    delete items[args.connectionId];
    await saveConnectionStoreDocument(args.settings, items);
    return true;
  }

  const previousSerialized = JSON.stringify(previous ?? null);
  const nextSerialized = JSON.stringify(next);
  if (previousSerialized === nextSerialized) {
    return false;
  }

  items[args.connectionId] = next;
  await saveConnectionStoreDocument(args.settings, items);
  return true;
}

export class ConnectionManager {
  private readonly settings: ConnectionSettingsStore;
  private readonly definitions = new Map<string, RegisteredConnectionDefinition>();
  private readonly listeners = new Set<ConnectionManagerListener>();

  constructor(options: { settings: ConnectionSettingsStore }) {
    this.settings = options.settings;
  }

  subscribe(listener: ConnectionManagerListener): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  private notify(): void {
    for (const listener of this.listeners) {
      try {
        listener();
      } catch (error: unknown) {
        console.warn("[pi] Connection manager listener failed:", error);
      }
    }
  }

  registerDefinition(ownerId: string, definition: ConnectionDefinition): string {
    const normalizedOwnerId = normalizeOwnerId(ownerId);
    const normalizedDefinition = normalizeConnectionDefinition(definition);

    const existing = this.definitions.get(normalizedDefinition.id);
    if (existing && existing.ownerId !== normalizedOwnerId) {
      throw new Error(
        `Connection id "${normalizedDefinition.id}" is already registered by another extension/runtime owner.`,
      );
    }

    this.definitions.set(normalizedDefinition.id, {
      ...normalizedDefinition,
      ownerId: normalizedOwnerId,
    });

    this.notify();
    return normalizedDefinition.id;
  }

  unregisterDefinition(ownerId: string, connectionId: string): boolean {
    const normalizedOwnerId = normalizeOwnerId(ownerId);
    const normalizedConnectionId = normalizeConnectionId(connectionId);
    const existing = this.definitions.get(normalizedConnectionId);
    if (!existing) return false;

    if (existing.ownerId !== normalizedOwnerId) {
      throw new Error(
        `Connection "${normalizedConnectionId}" is owned by another extension/runtime owner.`,
      );
    }

    this.definitions.delete(normalizedConnectionId);
    this.notify();
    return true;
  }

  unregisterDefinitionsByOwner(ownerId: string): void {
    const normalizedOwnerId = normalizeOwnerId(ownerId);
    let changed = false;

    for (const [connectionId, definition] of this.definitions.entries()) {
      if (definition.ownerId !== normalizedOwnerId) continue;
      this.definitions.delete(connectionId);
      changed = true;
    }

    if (changed) {
      this.notify();
    }
  }

  hasDefinition(connectionId: string): boolean {
    const normalizedConnectionId = normalizeConnectionId(connectionId);
    return this.definitions.has(normalizedConnectionId);
  }

  listRegisteredConnectionIds(): string[] {
    return Array.from(this.definitions.keys()).sort((left, right) => left.localeCompare(right));
  }

  private getRequiredDefinition(connectionId: string): RegisteredConnectionDefinition {
    const normalizedConnectionId = normalizeConnectionId(connectionId);
    const definition = this.definitions.get(normalizedConnectionId);
    if (!definition) {
      throw new Error(`Connection "${normalizedConnectionId}" is not registered.`);
    }

    return definition;
  }

  assertConnectionOwnedBy(ownerId: string, connectionId: string): void {
    const normalizedOwnerId = normalizeOwnerId(ownerId);
    const definition = this.getRequiredDefinition(connectionId);

    if (definition.ownerId !== normalizedOwnerId) {
      throw new Error(
        `Connection "${definition.id}" is not owned by this extension/runtime owner.`,
      );
    }
  }

  async setSecrets(ownerId: string, connectionId: string, secrets: Record<string, string>): Promise<void> {
    this.assertConnectionOwnedBy(ownerId, connectionId);
    const definition = this.getRequiredDefinition(connectionId);
    const normalizedSecrets = normalizeSecrets(secrets);

    const allowedSecretIds = new Set<string>(definition.secretFields.map((field) => field.id));
    for (const fieldId of Object.keys(normalizedSecrets)) {
      if (!allowedSecretIds.has(fieldId)) {
        throw new Error(`Unknown secret field "${fieldId}" for connection "${definition.id}".`);
      }
    }

    const now = new Date().toISOString();
    const requiredPresent = hasRequiredSecrets(definition, normalizedSecrets);
    const nextStatus: ConnectionStatus = requiredPresent ? "connected" : "missing";

    const changed = await mutateConnectionRecord({
      settings: this.settings,
      connectionId: definition.id,
      mutator: (record) => {
        const previousStatus = record?.status ?? "missing";
        assertStatusTransition(previousStatus, nextStatus);

        return {
          status: nextStatus,
          lastValidatedAt: requiredPresent ? now : undefined,
          lastError: undefined,
          secrets: normalizedSecrets,
        };
      },
    });

    if (changed) {
      this.notify();
    }
  }

  async clearSecrets(ownerId: string, connectionId: string): Promise<void> {
    this.assertConnectionOwnedBy(ownerId, connectionId);
    const definition = this.getRequiredDefinition(connectionId);

    const changed = await mutateConnectionRecord({
      settings: this.settings,
      connectionId: definition.id,
      mutator: (record) => {
        const previousStatus = record?.status ?? "missing";
        assertStatusTransition(previousStatus, "missing");

        return {
          status: "missing",
          lastValidatedAt: undefined,
          lastError: undefined,
          secrets: {},
        };
      },
    });

    if (changed) {
      this.notify();
    }
  }

  async markValidated(ownerId: string, connectionId: string): Promise<void> {
    this.assertConnectionOwnedBy(ownerId, connectionId);
    const definition = this.getRequiredDefinition(connectionId);

    const changed = await mutateConnectionRecord({
      settings: this.settings,
      connectionId: definition.id,
      mutator: (record) => {
        const previousStatus = record?.status ?? "missing";

        if (!hasRequiredSecrets(definition, record?.secrets)) {
          assertStatusTransition(previousStatus, "missing");
          return {
            ...record,
            status: "missing",
            lastValidatedAt: undefined,
            lastError: undefined,
          };
        }

        assertStatusTransition(previousStatus, "connected");
        return {
          ...record,
          status: "connected",
          lastValidatedAt: new Date().toISOString(),
          lastError: undefined,
        };
      },
    });

    if (changed) {
      this.notify();
    }
  }

  async markInvalid(ownerId: string, connectionId: string, reason: string): Promise<void> {
    this.assertConnectionOwnedBy(ownerId, connectionId);
    const definition = this.getRequiredDefinition(connectionId);
    const trimmedReason = normalizeNonEmpty(reason, "reason");

    const changed = await mutateConnectionRecord({
      settings: this.settings,
      connectionId: definition.id,
      mutator: (record) => {
        const previousStatus = record?.status ?? "missing";
        assertStatusTransition(previousStatus, "invalid");

        const redactedReason = redactSecretsInMessage(trimmedReason, record?.secrets);

        return {
          ...record,
          status: "invalid",
          lastValidatedAt: new Date().toISOString(),
          lastError: redactedReason,
        };
      },
    });

    if (changed) {
      this.notify();
    }
  }

  async markRuntimeAuthFailure(connectionId: string, failure: ConnectionRuntimeAuthFailure): Promise<void> {
    const definition = this.getRequiredDefinition(connectionId);
    const message = normalizeNonEmpty(failure.message, "failure message");

    const changed = await mutateConnectionRecord({
      settings: this.settings,
      connectionId: definition.id,
      mutator: (record) => {
        const previousStatus = record?.status ?? "missing";
        assertStatusTransition(previousStatus, "error");

        const redactedMessage = redactSecretsInMessage(message, record?.secrets);

        return {
          ...record,
          status: "error",
          lastValidatedAt: new Date().toISOString(),
          lastError: redactedMessage,
        };
      },
    });

    if (changed) {
      this.notify();
    }
  }

  async getSnapshot(connectionId: string): Promise<ConnectionSnapshot | null> {
    const definition = this.definitions.get(normalizeConnectionId(connectionId));
    if (!definition) {
      return null;
    }

    const items = await loadConnectionStoreDocument(this.settings);
    const record = items[definition.id];

    const status = resolveEffectiveStatus(definition, record);

    const setupHint = definition.setupHint ?? buildDefaultSetupHint(definition.title);

    const lastError = (status === "invalid" || status === "error")
      ? record?.lastError
      : undefined;

    return {
      connectionId: definition.id,
      title: definition.title,
      capability: definition.capability,
      status,
      setupHint,
      lastValidatedAt: record?.lastValidatedAt,
      lastError,
    };
  }

  async getState(connectionId: string): Promise<ConnectionState | null> {
    const snapshot = await this.getSnapshot(connectionId);
    if (!snapshot) return null;

    return {
      connectionId: snapshot.connectionId,
      status: snapshot.status,
      lastValidatedAt: snapshot.lastValidatedAt,
      lastError: snapshot.lastError,
    };
  }

  async listSnapshots(connectionIds?: readonly string[]): Promise<ConnectionSnapshot[]> {
    const ids = connectionIds && connectionIds.length > 0
      ? Array.from(new Set(connectionIds.map((id) => normalizeConnectionId(id))))
      : this.listRegisteredConnectionIds();

    const snapshots: ConnectionSnapshot[] = [];

    for (const connectionId of ids) {
      const snapshot = await this.getSnapshot(connectionId);
      if (!snapshot) continue;
      snapshots.push(snapshot);
    }

    snapshots.sort((left, right) => left.title.localeCompare(right.title));
    return snapshots;
  }

  async listPromptEntries(connectionIds: readonly string[]): Promise<ConnectionPromptEntry[]> {
    const snapshots = await this.listSnapshots(connectionIds);

    return snapshots.map((snapshot) => ({
      id: snapshot.connectionId,
      title: snapshot.title,
      capability: snapshot.capability,
      status: snapshot.status,
      setupHint: snapshot.setupHint,
      lastError: snapshot.lastError,
    }));
  }
}

export function looksLikeConnectionAuthFailure(message: string): boolean {
  const normalized = message.toLowerCase();

  return normalized.includes("401")
    || normalized.includes("403")
    || normalized.includes("unauthorized")
    || normalized.includes("forbidden")
    || normalized.includes("auth failed")
    || normalized.includes("authentication failed")
    || normalized.includes("invalid api key")
    || normalized.includes("invalid token")
    || normalized.includes("token expired")
    || normalized.includes("token revoked")
    || normalized.includes("api key expired")
    || normalized.includes("credential") && normalized.includes("invalid");
}
