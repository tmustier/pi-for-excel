import { APP_NAME, APP_VERSION } from "../../app/metadata.js";
import { getEnabledProxyBaseUrl } from "../../tools/external-fetch.js";
import type { McpServerConfig } from "../../tools/mcp-config.js";
import {
  MCP_PROTOCOL_VERSION,
  postMcpJsonRpc,
  type McpSession,
} from "../../tools/mcp-transport.js";
import type { SettingsWriter } from "../../storage/local/settings-store.js";

const MCP_PROBE_TIMEOUT_MS = 8_000;

export async function probeMcpServer(
  server: McpServerConfig,
  settings: SettingsWriter,
): Promise<{ toolCount: number; proxied: boolean; proxyBaseUrl?: string }> {
  const session: McpSession = {};
  const proxyBaseUrl = await getEnabledProxyBaseUrl(settings);
  const connection = {
    server,
    session,
    signal: undefined,
    timeoutMs: MCP_PROBE_TIMEOUT_MS,
    ...(proxyBaseUrl !== undefined ? { proxyBaseUrl } : {}),
  };
  await postMcpJsonRpc({
    ...connection,
    method: "initialize",
    params: {
      protocolVersion: MCP_PROTOCOL_VERSION,
      capabilities: {},
      clientInfo: { name: APP_NAME, version: APP_VERSION },
    },
  });
  await postMcpJsonRpc({
    ...connection,
    method: "notifications/initialized",
    expectResponse: false,
  });
  const list = await postMcpJsonRpc({
    ...connection,
    method: "tools/list",
    params: {},
  });
  const tools = list.result?.result?.tools;
  if (!Array.isArray(tools)) throw new Error("Invalid MCP tools/list result.");
  return {
    toolCount: tools.length,
    proxied: list.proxied,
    ...(list.proxyBaseUrl !== undefined ? { proxyBaseUrl: list.proxyBaseUrl } : {}),
  };
}
