function isToolsMcpConfigPayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * MCP server configuration storage.
 */

import {
  loadConnectionStoreDocument,
  loadConnectionStoreDocumentForUpdate,
  saveConnectionStoreDocument,
  type StoredConnectionRecord,
} from "../connections/store.js";
import type { SettingsWriter } from "../storage/local/settings-store.js";

export const MCP_SERVERS_SETTING_KEY = "mcp.servers.v1";
const MCP_SERVERS_DOC_VERSION = 1;

/** Connection-store record for MCP server bearer tokens keyed by server id. */
export const MCP_SERVER_TOKENS_CONNECTION_ID = "builtin.mcp.servers";

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
  if (!isToolsMcpConfigPayloadShape(raw)) return null;

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
  const server: McpServerConfig = {
    id,
    name,
    url,
    enabled: normalizeEnabled(raw.enabled),
  };

  if (token !== undefined) {
    server.token = token;
  }

  return server;
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
    : isToolsMcpConfigPayloadShape(raw) && Array.isArray(raw.servers)
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

async function loadLegacyMcpServers(settings: SettingsWriter): Promise<McpServerConfig[]> {
  try {
    return normalizeServers(await settings.get(MCP_SERVERS_SETTING_KEY));
  } catch {
    return [];
  }
}

async function loadLegacyMcpServersForUpdate(settings: SettingsWriter): Promise<McpServerConfig[]> {
  return normalizeServers(await settings.get(MCP_SERVERS_SETTING_KEY));
}

function readConnectionStoreMcpTokens(
  items: Record<string, StoredConnectionRecord>,
): Record<string, string> {
  return normalizeTokenMap(items[MCP_SERVER_TOKENS_CONNECTION_ID]?.secrets ?? {});
}

async function loadConnectionStoreMcpTokens(
  settings: SettingsWriter,
): Promise<Record<string, string>> {
  return readConnectionStoreMcpTokens(await loadConnectionStoreDocument(settings));
}

async function loadConnectionStoreMcpTokensForUpdate(
  settings: SettingsWriter,
): Promise<Record<string, string>> {
  return readConnectionStoreMcpTokens(await loadConnectionStoreDocumentForUpdate(settings));
}

async function writeConnectionStoreMcpTokens(
  settings: SettingsWriter,
  tokensByServerId: Readonly<Record<string, string>>,
): Promise<void> {
  const normalizedTokens = normalizeTokenMap(tokensByServerId);
  const items = await loadConnectionStoreDocumentForUpdate(settings);
  const previous = items[MCP_SERVER_TOKENS_CONNECTION_ID];

  if (Object.keys(normalizedTokens).length === 0) {
    if (MCP_SERVER_TOKENS_CONNECTION_ID in items) {
      delete items[MCP_SERVER_TOKENS_CONNECTION_ID];
      await saveConnectionStoreDocument(settings, items);
    }
    return;
  }

  const record: StoredConnectionRecord = {
    status: "connected",
    secrets: normalizedTokens,
  };
  const previousLastValidatedAt = normalizeOptionalString(previous?.lastValidatedAt);
  if (previousLastValidatedAt !== undefined) {
    record.lastValidatedAt = previousLastValidatedAt;
  }
  items[MCP_SERVER_TOKENS_CONNECTION_ID] = record;

  await saveConnectionStoreDocument(settings, items);
}

function mergeServersWithConnectionTokens(args: {
  servers: readonly McpServerConfig[];
  connectionTokens: Readonly<Record<string, string>>;
}): McpServerConfig[] {
  return args.servers.map((server) => {
    const token = normalizeOptionalString(args.connectionTokens[server.id])
      ?? normalizeOptionalString(server.token);
    const merged: McpServerConfig = {
      id: server.id,
      name: server.name,
      url: server.url,
      enabled: server.enabled,
    };
    if (token !== undefined) {
      merged.token = token;
    }
    return merged;
  });
}

export async function migrateLegacyMcpTokensToConnectionStore(
  settings: SettingsWriter,
): Promise<boolean> {
  const [legacyServers, connectionTokens] = await Promise.all([
    loadLegacyMcpServersForUpdate(settings),
    loadConnectionStoreMcpTokensForUpdate(settings),
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

export async function loadMcpServers(settings: SettingsWriter): Promise<McpServerConfig[]> {
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
  settings: SettingsWriter,
  servers: readonly McpServerConfig[],
): Promise<void> {
  const normalized = uniqueById(normalizeServers(servers));
  const tokensByServerId = readTokenMapFromServers(normalized);

  // Write tokens first so a failed connection-store write never strips legacy
  // token fields from mcp.servers.v1 before persistence succeeds.
  const previousTokenMap = await loadConnectionStoreMcpTokensForUpdate(settings);
  await writeConnectionStoreMcpTokens(settings, tokensByServerId);

  try {
    await settings.set(MCP_SERVERS_SETTING_KEY, createDocument(stripServerTokens(normalized)));
  } catch (error) {
    try {
      const currentTokenMap = await loadConnectionStoreMcpTokensForUpdate(settings);
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
  const server: McpServerConfig = {
    id: `mcp-${crypto.randomUUID()}`,
    name,
    url,
    enabled: input.enabled ?? true,
  };

  if (token !== undefined) {
    server.token = token;
  }

  return server;
}
