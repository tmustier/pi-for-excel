/**
 * MCP server configuration storage.
 */

import { isRecord } from "../utils/type-guards.js";

export const MCP_SERVERS_SETTING_KEY = "mcp.servers.v1";
const MCP_SERVERS_DOC_VERSION = 1;

const CONNECTION_STORE_KEY = "connections.store.v1";
const CONNECTION_STORE_VERSION = 1;

/** Connection-store record for MCP server bearer tokens keyed by server id. */
export const MCP_SERVER_TOKENS_CONNECTION_ID = "builtin.mcp.servers";

export interface McpConfigStore {
  get(key: string): Promise<unknown>;
  set(key: string, value: unknown): Promise<void>;
}

export interface McpServerConfig {
  id: string;
  name: string;
  url: string;
  enabled: boolean;
  token?: string;
}

interface McpServersDocument {
  version: number;
  servers: Array<Omit<McpServerConfig, "token">>;
}

type StoredConnectionStatus = "connected" | "missing" | "invalid" | "error";

type StoredConnectionRecord = {
  status?: StoredConnectionStatus;
  lastValidatedAt?: string;
  lastError?: string;
  secrets?: Record<string, string>;
};

function normalizeOptionalString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function normalizeName(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function normalizeEnabled(value: unknown): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "0" || normalized === "false" || normalized === "off") {
      return false;
    }
  }
  return true;
}

export function validateMcpServerUrl(url: string): string {
  const trimmed = url.trim();
  if (trimmed.length === 0) {
    throw new Error("MCP server URL cannot be empty.");
  }

  let parsed: URL;
  try {
    parsed = new URL(trimmed);
  } catch {
    throw new Error("Invalid MCP server URL.");
  }

  if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
    throw new Error("MCP server URL must use http:// or https://");
  }

  return trimmed.replace(/\/+$/u, "");
}

function normalizeServerId(value: unknown, fallbackName: string, fallbackUrl: string): string {
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed.length > 0) {
      return trimmed;
    }
  }

  const base = `${fallbackName} ${fallbackUrl}`
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 48);

  return base.length > 0 ? `mcp-${base}` : `mcp-${crypto.randomUUID()}`;
}

function normalizeServer(raw: unknown): McpServerConfig | null {
  if (!isRecord(raw)) return null;

  const name = normalizeName(raw.name);
  const rawUrl = normalizeOptionalString(raw.url);
  if (!name || !rawUrl) return null;

  let url: string;
  try {
    url = validateMcpServerUrl(rawUrl);
  } catch {
    return null;
  }

  const id = normalizeServerId(raw.id, name, url);
  const token = normalizeOptionalString(raw.token);

  return {
    id,
    name,
    url,
    enabled: normalizeEnabled(raw.enabled),
    token,
  };
}

function uniqueById(servers: McpServerConfig[]): McpServerConfig[] {
  const used = new Set<string>();
  const out: McpServerConfig[] = [];

  for (const server of servers) {
    let candidate = server.id;
    if (used.has(candidate)) {
      let suffix = 2;
      while (used.has(`${candidate}-${suffix}`)) {
        suffix += 1;
      }
      candidate = `${candidate}-${suffix}`;
    }

    used.add(candidate);
    out.push({
      ...server,
      id: candidate,
    });
  }

  return out;
}

function normalizeServers(raw: unknown): McpServerConfig[] {
  const source = Array.isArray(raw)
    ? raw
    : isRecord(raw) && Array.isArray(raw.servers)
      ? raw.servers
      : [];

  const parsed: McpServerConfig[] = [];
  for (const item of source) {
    const normalized = normalizeServer(item);
    if (!normalized) continue;
    parsed.push(normalized);
  }

  return uniqueById(parsed);
}

function stripServerTokens(servers: readonly McpServerConfig[]): Array<Omit<McpServerConfig, "token">> {
  return servers.map((server) => ({
    id: server.id,
    name: server.name,
    url: server.url,
    enabled: server.enabled,
  }));
}

function createDocument(servers: Array<Omit<McpServerConfig, "token">>): McpServersDocument {
  return {
    version: MCP_SERVERS_DOC_VERSION,
    servers,
  };
}

function normalizeTokenMap(tokens: Readonly<Record<string, string>>): Record<string, string> {
  const normalized: Record<string, string> = {};
  const sortedIds = Object.keys(tokens).sort((left, right) => left.localeCompare(right));

  for (const serverId of sortedIds) {
    const token = normalizeOptionalString(tokens[serverId]);
    if (!token) continue;
    normalized[serverId] = token;
  }

  return normalized;
}

function readTokenMapFromServers(servers: readonly McpServerConfig[]): Record<string, string> {
  const tokens: Record<string, string> = {};

  for (const server of servers) {
    const token = normalizeOptionalString(server.token);
    if (!token) continue;
    tokens[server.id] = token;
  }

  return normalizeTokenMap(tokens);
}

function mergeTokenMaps(args: {
  primary: Readonly<Record<string, string>>;
  fallback: Readonly<Record<string, string>>;
}): Record<string, string> {
  return normalizeTokenMap({
    ...args.fallback,
    ...args.primary,
  });
}

function areTokenMapsEqual(left: Readonly<Record<string, string>>, right: Readonly<Record<string, string>>): boolean {
  return JSON.stringify(normalizeTokenMap(left)) === JSON.stringify(normalizeTokenMap(right));
}

async function loadLegacyMcpServers(settings: McpConfigStore): Promise<McpServerConfig[]> {
  const raw = await settings.get(MCP_SERVERS_SETTING_KEY);
  return normalizeServers(raw);
}

async function loadConnectionStoreItems(
  settings: McpConfigStore,
): Promise<Record<string, StoredConnectionRecord>> {
  const raw = await settings.get(CONNECTION_STORE_KEY);
  if (!isRecord(raw)) return {};

  const rawItems = raw.items;
  if (!isRecord(rawItems)) return {};

  const items: Record<string, StoredConnectionRecord> = {};

  for (const [connectionId, rawRecord] of Object.entries(rawItems)) {
    if (!isRecord(rawRecord)) continue;

    const rawSecrets = rawRecord.secrets;
    const secrets: Record<string, string> = {};
    if (isRecord(rawSecrets)) {
      for (const [fieldId, value] of Object.entries(rawSecrets)) {
        const normalized = normalizeOptionalString(value);
        if (!normalized) continue;
        secrets[fieldId] = normalized;
      }
    }

    const status = rawRecord.status;
    items[connectionId] = {
      status: status === "connected" || status === "missing" || status === "invalid" || status === "error"
        ? status
        : undefined,
      lastValidatedAt: normalizeOptionalString(rawRecord.lastValidatedAt),
      lastError: normalizeOptionalString(rawRecord.lastError),
      secrets,
    };
  }

  return items;
}

async function saveConnectionStoreItems(
  settings: McpConfigStore,
  items: Record<string, StoredConnectionRecord>,
): Promise<void> {
  await settings.set(CONNECTION_STORE_KEY, {
    version: CONNECTION_STORE_VERSION,
    items,
  });
}

async function loadConnectionStoreMcpTokens(
  settings: McpConfigStore,
): Promise<Record<string, string>> {
  const items = await loadConnectionStoreItems(settings);
  const rawTokens = items[MCP_SERVER_TOKENS_CONNECTION_ID]?.secrets ?? {};
  return normalizeTokenMap(rawTokens);
}

async function writeConnectionStoreMcpTokens(
  settings: McpConfigStore,
  tokensByServerId: Readonly<Record<string, string>>,
): Promise<void> {
  const normalizedTokens = normalizeTokenMap(tokensByServerId);
  const items = await loadConnectionStoreItems(settings);
  const previous = items[MCP_SERVER_TOKENS_CONNECTION_ID];

  if (Object.keys(normalizedTokens).length === 0) {
    if (MCP_SERVER_TOKENS_CONNECTION_ID in items) {
      delete items[MCP_SERVER_TOKENS_CONNECTION_ID];
      await saveConnectionStoreItems(settings, items);
    }
    return;
  }

  items[MCP_SERVER_TOKENS_CONNECTION_ID] = {
    status: "connected",
    lastValidatedAt: previous?.lastValidatedAt,
    lastError: undefined,
    secrets: normalizedTokens,
  };

  await saveConnectionStoreItems(settings, items);
}

function mergeServersWithConnectionTokens(args: {
  servers: readonly McpServerConfig[];
  connectionTokens: Readonly<Record<string, string>>;
}): McpServerConfig[] {
  return args.servers.map((server) => ({
    ...server,
    token: normalizeOptionalString(args.connectionTokens[server.id])
      ?? normalizeOptionalString(server.token),
  }));
}

export async function migrateLegacyMcpTokensToConnectionStore(
  settings: McpConfigStore,
): Promise<boolean> {
  const [legacyServers, connectionTokens] = await Promise.all([
    loadLegacyMcpServers(settings),
    loadConnectionStoreMcpTokens(settings),
  ]);

  const legacyTokens = readTokenMapFromServers(legacyServers);
  const mergedTokens = mergeTokenMaps({
    primary: connectionTokens,
    fallback: legacyTokens,
  });

  const shouldWriteConnectionStore = !areTokenMapsEqual(mergedTokens, connectionTokens);
  const hasLegacyTokens = Object.keys(legacyTokens).length > 0;

  if (!shouldWriteConnectionStore && !hasLegacyTokens) {
    return false;
  }

  if (shouldWriteConnectionStore) {
    await writeConnectionStoreMcpTokens(settings, mergedTokens);
  }

  if (hasLegacyTokens) {
    await settings.set(MCP_SERVERS_SETTING_KEY, createDocument(stripServerTokens(legacyServers)));
  }

  return true;
}

export async function loadMcpServers(settings: McpConfigStore): Promise<McpServerConfig[]> {
  const [legacyServers, connectionTokens] = await Promise.all([
    loadLegacyMcpServers(settings),
    loadConnectionStoreMcpTokens(settings),
  ]);

  return mergeServersWithConnectionTokens({
    servers: legacyServers,
    connectionTokens,
  });
}

export async function saveMcpServers(
  settings: McpConfigStore,
  servers: readonly McpServerConfig[],
): Promise<void> {
  const normalized = uniqueById(normalizeServers(servers));
  const tokensByServerId = readTokenMapFromServers(normalized);

  // Write tokens first so a failed connection-store write never strips legacy
  // token fields from mcp.servers.v1 before persistence succeeds.
  const previousTokenMap = await loadConnectionStoreMcpTokens(settings);
  await writeConnectionStoreMcpTokens(settings, tokensByServerId);

  try {
    await settings.set(MCP_SERVERS_SETTING_KEY, createDocument(stripServerTokens(normalized)));
  } catch (error: unknown) {
    try {
      const currentTokenMap = await loadConnectionStoreMcpTokens(settings);
      const rollbackIsSafe = areTokenMapsEqual(currentTokenMap, tokensByServerId);

      if (rollbackIsSafe) {
        await writeConnectionStoreMcpTokens(settings, previousTokenMap);
      }
    } catch {
      // best-effort rollback only; rethrow original failure below.
    }

    throw error;
  }
}

export function createMcpServerConfig(input: {
  name: string;
  url: string;
  token?: string;
  enabled?: boolean;
}): McpServerConfig {
  const name = normalizeName(input.name);
  if (!name) {
    throw new Error("MCP server name cannot be empty.");
  }

  const url = validateMcpServerUrl(input.url);
  const token = normalizeOptionalString(input.token);

  return {
    id: `mcp-${crypto.randomUUID()}`,
    name,
    url,
    enabled: input.enabled ?? true,
    token,
  };
}
