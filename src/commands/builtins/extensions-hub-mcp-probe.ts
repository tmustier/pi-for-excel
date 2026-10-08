import { APP_NAME, APP_VERSION } from "../../app/metadata.js";
import { getEnabledProxyBaseUrl, resolveOutboundRequestUrl } from "../../tools/external-fetch.js";
import { type McpServerConfig } from "../../tools/mcp-config.js";
import {
  getHttpErrorReason,
  runWithTimeoutAbort,
} from "../../utils/network.js";
import type { SettingsWriter } from "../../storage/local/settings-store.js";
import {
  buildMcpRequestHeaders,
  MCP_SESSION_ID_HEADER,
  parseMcpResponseBody,
} from "../../tools/mcp-transport.js";
function isExtensionsHubMcpProbePayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}


const MCP_PROBE_TIMEOUT_MS = 8_000;
const MCP_PROTOCOL_VERSION = "2025-03-26";

function parseToolCountFromListResponse(value: unknown): number {
  if (!isExtensionsHubMcpProbePayloadShape(value)) return 0;
  if (!isExtensionsHubMcpProbePayloadShape(value.result)) return 0;
  const tools = value.result.tools;
  return Array.isArray(tools) ? tools.length : 0;
}

async function postJsonRpc(args: {
  server: McpServerConfig;
  method: string;
  params?: unknown;
  settings: SettingsWriter;
  session: { id?: string };
  expectResponse?: boolean;
}): Promise<{ response: Record<string, unknown> | null; proxied: boolean; proxyBaseUrl?: string } | null> {
  const { server, method, params, settings, session, expectResponse = true } = args;

  const proxyBaseUrl = await getEnabledProxyBaseUrl(settings);
  const resolved = resolveOutboundRequestUrl({
    targetUrl: server.url,
    ...(proxyBaseUrl !== undefined ? { proxyBaseUrl } : {}),
  });

  const body: Record<string, unknown> = {
    jsonrpc: "2.0",
    method,
  };

  if (params !== undefined) {
    body.params = params;
  }

  const requestId = expectResponse ? crypto.randomUUID() : undefined;
  if (requestId !== undefined) {
    body.id = requestId;
  }

  const headers = buildMcpRequestHeaders({
    ...(server.token ? { token: server.token } : {}),
    ...(session.id !== undefined ? { sessionId: session.id } : {}),
  });

  return runWithTimeoutAbort({
    signal: undefined,
    timeoutMs: MCP_PROBE_TIMEOUT_MS,
    timeoutErrorMessage: `MCP request timed out after ${MCP_PROBE_TIMEOUT_MS}ms.`,
    run: async (requestSignal) => {
      const response = await fetch(resolved.requestUrl, {
        method: "POST",
        headers,
        body: JSON.stringify(body),
        signal: requestSignal,
      });

      if (!response.ok) {
        const text = await response.text();
        const reason = getHttpErrorReason(response.status, text);
        throw new Error(`MCP request failed (${response.status}): ${reason}`);
      }

      const assignedSessionId = response.headers.get(MCP_SESSION_ID_HEADER);
      if (assignedSessionId) {
        session.id = assignedSessionId;
      }

      if (!expectResponse) {
        return {
          response: null,
          proxied: resolved.proxied,
          ...(resolved.proxyBaseUrl !== undefined ? { proxyBaseUrl: resolved.proxyBaseUrl } : {}),
        };
      }

      const text = await response.text();
      const payload = parseMcpResponseBody(text, response.headers.get("Content-Type"), requestId);

      return {
        response: payload,
        proxied: resolved.proxied,
        ...(resolved.proxyBaseUrl !== undefined ? { proxyBaseUrl: resolved.proxyBaseUrl } : {}),
      };
    },
  });
}

export async function probeMcpServer(
  server: McpServerConfig,
  settings: SettingsWriter,
): Promise<{ toolCount: number; proxied: boolean; proxyBaseUrl?: string }> {
  const session: { id?: string } = {};

  await postJsonRpc({
    server,
    method: "initialize",
    params: {
      protocolVersion: MCP_PROTOCOL_VERSION,
      capabilities: {},
      clientInfo: {
        name: APP_NAME,
        version: APP_VERSION,
      },
    },
    settings,
    session,
  });

  await postJsonRpc({
    server,
    method: "notifications/initialized",
    settings,
    session,
    expectResponse: false,
  });

  const list = await postJsonRpc({
    server,
    method: "tools/list",
    params: {},
    settings,
    session,
  });

  if (!list) {
    throw new Error("MCP tools/list returned no response.");
  }

  return {
    toolCount: parseToolCountFromListResponse(list.response),
    proxied: list.proxied,
    ...(list.proxyBaseUrl !== undefined ? { proxyBaseUrl: list.proxyBaseUrl } : {}),
  };
}
