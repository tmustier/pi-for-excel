import { APP_NAME, APP_VERSION } from "../../app/metadata.js";
import { type IntegrationSettingsStore } from "../../integrations/store.js";
import { getEnabledProxyBaseUrl, resolveOutboundRequestUrl } from "../../tools/external-fetch.js";
import { type McpServerConfig } from "../../tools/mcp-config.js";
import {
  getHttpErrorReason,
  runWithTimeoutAbort,
} from "../../utils/network.js";
import { isRecord } from "../../utils/type-guards.js";

const MCP_PROBE_TIMEOUT_MS = 8_000;
const MCP_PROTOCOL_VERSION = "2025-03-26";

function parseToolCountFromListResponse(value: unknown): number {
  if (!isRecord(value)) return 0;
  if (!isRecord(value.result)) return 0;
  const tools = value.result.tools;
  return Array.isArray(tools) ? tools.length : 0;
}

async function postJsonRpc(args: {
  server: McpServerConfig;
  method: string;
  params?: unknown;
  settings: IntegrationSettingsStore;
  expectResponse?: boolean;
}): Promise<{ response: unknown; proxied: boolean; proxyBaseUrl?: string } | null> {
  const { server, method, params, settings, expectResponse = true } = args;

  const proxyBaseUrl = await getEnabledProxyBaseUrl(settings);
  const resolved = resolveOutboundRequestUrl({
    targetUrl: server.url,
    proxyBaseUrl,
  });

  const body: Record<string, unknown> = {
    jsonrpc: "2.0",
    method,
  };

  if (params !== undefined) {
    body.params = params;
  }

  if (expectResponse) {
    body.id = crypto.randomUUID();
  }

  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    Accept: "application/json",
  };

  if (server.token) {
    headers.Authorization = `Bearer ${server.token}`;
  }

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

      if (!expectResponse) {
        return {
          response: null,
          proxied: resolved.proxied,
          proxyBaseUrl: resolved.proxyBaseUrl,
        };
      }

      const text = await response.text();
      const payload: unknown = text.trim().length > 0 ? JSON.parse(text) : null;

      return {
        response: payload,
        proxied: resolved.proxied,
        proxyBaseUrl: resolved.proxyBaseUrl,
      };
    },
  });
}

export async function probeMcpServer(
  server: McpServerConfig,
  settings: IntegrationSettingsStore,
): Promise<{ toolCount: number; proxied: boolean; proxyBaseUrl?: string }> {
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
  });

  await postJsonRpc({
    server,
    method: "notifications/initialized",
    settings,
    expectResponse: false,
  });

  const list = await postJsonRpc({
    server,
    method: "tools/list",
    params: {},
    settings,
  });

  if (!list) {
    throw new Error("MCP tools/list returned no response.");
  }

  return {
    toolCount: parseToolCountFromListResponse(list.response),
    proxied: list.proxied,
    proxyBaseUrl: list.proxyBaseUrl,
  };
}
