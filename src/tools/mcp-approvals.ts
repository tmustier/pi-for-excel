/** Remembered consent is scoped to a server identity, endpoint and tool. */
import { Type } from "typebox";
import { Value } from "typebox/value";
import type { SettingsReader, SettingsUpdater } from "../storage/local/settings-store.js";
import type { McpServerConfig } from "./mcp-config.js";

const GrantsSchema = Type.Array(Type.String());

function key(server: McpServerConfig): string {
  return `mcp.tool-approvals.v1:${JSON.stringify([server.id, server.url])}`;
}

export async function getMcpToolGrants(settings: SettingsReader, server: McpServerConfig): Promise<string[]> {
  const raw = await settings.get(key(server));
  return Value.Check(GrantsSchema, raw) ? raw : [];
}

export async function setMcpToolGrant(settings: SettingsUpdater, server: McpServerConfig, tool: string, allowed: boolean): Promise<void> {
  await settings.update(key(server), (raw) => Value.Check(GrantsSchema, raw) ? raw : [], (current) => {
    const grants = new Set(current);
    if (allowed) grants.add(tool);
    else grants.delete(tool);
    return [...grants].sort();
  });
}
