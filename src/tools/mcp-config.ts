/**
 * MCP server configuration storage.
 */

import { isRecord } from "../utils/type-guards.js";

export const MCP_SERVERS_SETTING_KEY = "mcp.servers.v1";
const MCP_SERVERS_DOC_VERSION = 1;

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
  servers: McpServerConfig[];
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

function createDocument(servers: McpServerConfig[]): McpServersDocument {
  return {
    version: MCP_SERVERS_DOC_VERSION,
    servers,
  };
}

export async function loadMcpServers(settings: McpConfigStore): Promise<McpServerConfig[]> {
  const raw = await settings.get(MCP_SERVERS_SETTING_KEY);
  return normalizeServers(raw);
}

export async function saveMcpServers(
  settings: McpConfigStore,
  servers: readonly McpServerConfig[],
): Promise<void> {
  const normalized = uniqueById(normalizeServers(servers));
  await settings.set(MCP_SERVERS_SETTING_KEY, createDocument(normalized));
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
