function isToolsMcpPayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * mcp — Model Context Protocol gateway for configured HTTP servers.
 */

import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { Type, type Static, type TSchema } from "typebox";

import { APP_NAME, APP_VERSION } from "../app/metadata.js";
import { integrationsCommandHint } from "../integrations/naming.js";
import { getErrorMessage } from "../utils/errors.js";
import { parseJson, type JsonValue } from "../utils/json.js";
import {
  buildProxyDownErrorMessage,
  getEnabledProxyBaseUrl,
  isLikelyProxyConnectionError,
} from "./external-fetch.js";
import type { SettingsReader, SettingsWriter } from "../storage/local/settings-store.js";
import {
  loadMcpServers,
  type McpServerConfig,
} from "./mcp-config.js";
import {
  MCP_PROTOCOL_VERSION,
  postMcpJsonRpc,
  type McpSession,
} from "./mcp-transport.js";

const MCP_CLIENT_NAME = APP_NAME;
const MCP_CLIENT_VERSION = APP_VERSION;
const MCP_TIMEOUT_MS = 15_000;

const schema = Type.Object({
  tool: Type.Optional(Type.String({
    description: "Tool name to call.",
  })),
  args: Type.Optional(Type.String({
    description: "Tool arguments as a JSON string.",
  })),
  connect: Type.Optional(Type.String({
    description: "Server name/id to connect and refresh.",
  })),
  describe: Type.Optional(Type.String({
    description: "Tool name to describe.",
  })),
  search: Type.Optional(Type.String({
    description: "Search query for MCP tools.",
  })),
  server: Type.Optional(Type.String({
    description: "Optional server name/id filter.",
  })),
});

type Params = Static<typeof schema>;

interface McpToolDescriptor {
  serverId: string;
  serverName: string;
  serverUrl: string;
  name: string;
  description?: string;
  inputSchema?: unknown;
}

interface ServerToolList {
  server: McpServerConfig;
  tools: McpToolDescriptor[];
  proxied: boolean;
  proxyBaseUrl?: string;
  session: McpSession;
  callJsonRpc: NonNullable<McpToolDependencies["callJsonRpc"]>;
}

interface PendingServerTools {
  server: McpServerConfig;
  proxyBaseUrl?: string;
  promise: Promise<ServerToolList>;
}

interface RpcCallResult {
  result: unknown;
  proxied: boolean;
  proxyBaseUrl?: string;
}

export interface McpGatewayDetails {
  kind: "mcp_gateway";
  ok: boolean;
  operation: string;
  server?: string;
  tool?: string;
  proxied?: boolean;
  proxyBaseUrl?: string;
  resultPreview?: string;
  error?: string;
  /** `true` when the failure is due to the local CORS proxy being unreachable. */
  proxyDown?: boolean;
}

export interface McpRuntimeConfig {
  servers: McpServerConfig[];
  proxyBaseUrl?: string;
}

export interface McpToolDependencies {
  getRuntimeConfig?: () => Promise<McpRuntimeConfig>;
  callJsonRpc?: (args: {
    server: McpServerConfig;
    method: string;
    params?: JsonValue;
    signal: AbortSignal | undefined;
    proxyBaseUrl?: string;
    expectResponse?: boolean;
  }) => Promise<RpcCallResult | null>;
}

function normalizeOptionalString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function proxyBaseUrlArg(proxyBaseUrl: string | undefined): { proxyBaseUrl?: string } {
  return proxyBaseUrl === undefined ? {} : { proxyBaseUrl };
}

function createMcpGatewayDetails(args: {
  ok: boolean;
  operation: string;
  server?: string | undefined;
  tool?: string | undefined;
  proxied?: boolean | undefined;
  proxyBaseUrl?: string | undefined;
  resultPreview?: string | undefined;
  error?: string | undefined;
  proxyDown?: boolean | undefined;
}): McpGatewayDetails {
  const details: McpGatewayDetails = {
    kind: "mcp_gateway",
    ok: args.ok,
    operation: args.operation,
  };
  if (args.server !== undefined) details.server = args.server;
  if (args.tool !== undefined) details.tool = args.tool;
  if (args.proxied !== undefined) details.proxied = args.proxied;
  if (args.proxyBaseUrl !== undefined) details.proxyBaseUrl = args.proxyBaseUrl;
  if (args.resultPreview !== undefined) details.resultPreview = args.resultPreview;
  if (args.error !== undefined) details.error = args.error;
  if (args.proxyDown !== undefined) details.proxyDown = args.proxyDown;
  return details;
}

function parseParams(raw: unknown): Params {
  if (!isToolsMcpPayloadShape(raw)) {
    return {};
  }

  const params: Params = {};

  const tool = normalizeOptionalString(raw.tool);
  if (tool) params.tool = tool;

  const args = normalizeOptionalString(raw.args);
  if (args) params.args = args;

  const connect = normalizeOptionalString(raw.connect);
  if (connect) params.connect = connect;

  const describe = normalizeOptionalString(raw.describe);
  if (describe) params.describe = describe;

  const search = normalizeOptionalString(raw.search);
  if (search) params.search = search;

  const server = normalizeOptionalString(raw.server);
  if (server) params.server = server;

  return params;
}

function normalizeServerToken(value: string): string {
  return value.trim().toLowerCase();
}

function findServerByToken(servers: readonly McpServerConfig[], token: string): McpServerConfig | null {
  const normalized = normalizeServerToken(token);

  for (const server of servers) {
    if (normalizeServerToken(server.id) === normalized) return server;
    if (normalizeServerToken(server.name) === normalized) return server;
  }

  return null;
}

function matchesServerToken(tool: McpToolDescriptor, token: string): boolean {
  const normalized = normalizeServerToken(token);
  return (
    normalizeServerToken(tool.serverId) === normalized ||
    normalizeServerToken(tool.serverName) === normalized
  );
}

function parseToolListResult(server: McpServerConfig, value: unknown): McpToolDescriptor[] {
  if (!isToolsMcpPayloadShape(value)) return [];
  if (!isToolsMcpPayloadShape(value.result)) return [];
  const result = value.result;

  const tools = result.tools;
  if (!Array.isArray(tools)) return [];

  const out: McpToolDescriptor[] = [];

  for (const item of tools) {
    if (!isToolsMcpPayloadShape(item)) continue;

    const name = normalizeOptionalString(item.name);
    if (!name) continue;

    const descriptor: McpToolDescriptor = {
      serverId: server.id,
      serverName: server.name,
      serverUrl: server.url,
      name,
    };
    const description = normalizeOptionalString(item.description);
    if (description !== undefined) {
      descriptor.description = description;
    }
    if (item.inputSchema !== undefined) {
      descriptor.inputSchema = item.inputSchema;
    }
    out.push(descriptor);
  }

  return out;
}

function extractTextContentBlocks(value: unknown): string[] {
  if (!Array.isArray(value)) return [];

  const lines: string[] = [];
  for (const item of value) {
    if (!isToolsMcpPayloadShape(item)) continue;
    if (item.type !== "text") continue;
    const text = normalizeOptionalString(item.text);
    if (!text) continue;
    lines.push(text);
  }

  return lines;
}

function formatJson(value: unknown): string {
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value);
  }
}

function parseCallArgs(rawArgs: string | undefined): JsonValue {
  if (!rawArgs) return {};

  try {
    return parseJson(rawArgs);
  } catch {
    throw new Error("mcp.args must be valid JSON.");
  }
}

function buildServerStatusLine(server: McpServerConfig, tools: McpToolDescriptor[] | null): string {
  if (!server.enabled) {
    return `- ${server.name} (${server.url}) — disabled`;
  }

  if (!tools) {
    return `- ${server.name} (${server.url}) — enabled`;
  }

  return `- ${server.name} (${server.url}) — enabled, ${tools.length} tool${tools.length === 1 ? "" : "s"}`;
}

function buildToolPreview(tools: readonly McpToolDescriptor[], max = 20): string {
  if (tools.length === 0) {
    return "No tools found.";
  }

  const lines: string[] = [];
  const count = Math.min(max, tools.length);

  for (const tool of tools.slice(0, count)) {
    const desc = tool.description ? ` — ${tool.description}` : "";
    lines.push(`- ${tool.name}${desc}`);
  }

  if (tools.length > count) {
    lines.push(`- … ${tools.length - count} more`);
  }

  return lines.join("\n");
}

function firstLine(value: string): string {
  const line = value.split("\n")[0] ?? value;
  return line.length > 220 ? `${line.slice(0, 217)}…` : line;
}

function toSearchTokens(query: string): string[] {
  return query
    .toLowerCase()
    .split(/\s+/)
    .map((token) => token.trim())
    .filter((token) => token.length > 0);
}

function matchesSearch(tool: McpToolDescriptor, query: string): boolean {
  const haystack = `${tool.name} ${tool.description ?? ""}`.toLowerCase();
  const tokens = toSearchTokens(query);
  if (tokens.length === 0) return true;

  return tokens.some((token) => haystack.includes(token));
}

async function defaultGetRuntimeConfig(): Promise<McpRuntimeConfig> {
  const storageModule = await import("../storage/local/app-storage.js");
  const settingsStore = storageModule.getAppStorage().settings;
  const configStore: SettingsWriter = settingsStore;
  const proxyStore: SettingsReader = settingsStore;

  const [servers, proxyBaseUrl] = await Promise.all([
    loadMcpServers(configStore),
    getEnabledProxyBaseUrl(proxyStore),
  ]);

  return {
    servers,
    ...proxyBaseUrlArg(proxyBaseUrl),
  };
}

function matchesConnection(
  connection: { server: McpServerConfig; proxyBaseUrl?: string },
  server: McpServerConfig,
  proxyBaseUrl: string | undefined,
): boolean {
  return connection.server.url === server.url
    && connection.server.token === server.token
    && connection.server.name === server.name
    && connection.proxyBaseUrl === proxyBaseUrl;
}

/** One caller's cancellation must not abort another caller's shared handshake. */
async function waitForServerTools(
  promise: Promise<ServerToolList>,
  signal: AbortSignal | undefined,
): Promise<ServerToolList> {
  signal?.throwIfAborted();
  if (!signal) return promise;
  let onAbort = (): void => {};
  const aborted = new Promise<never>((_resolve, reject) => {
    onAbort = () => reject(new Error("Aborted"));
    signal.addEventListener("abort", onAbort, { once: true });
  });
  try {
    return await Promise.race([promise, aborted]);
  } finally {
    signal.removeEventListener("abort", onAbort);
  }
}

export function createMcpTool(
  dependencies: McpToolDependencies = {},
): AgentTool<TSchema, McpGatewayDetails> {
  const getRuntimeConfig = dependencies.getRuntimeConfig ?? defaultGetRuntimeConfig;
  const toolCache = new Map<string, ServerToolList>();
  const pendingTools = new Map<string, PendingServerTools>();

  const discoverServerTools = async (
    server: McpServerConfig,
    proxyBaseUrl: string | undefined,
  ): Promise<ServerToolList> => {
    const session: McpSession = {};
    const callJsonRpc = dependencies.callJsonRpc ?? ((args) => postMcpJsonRpc({
      ...args,
      session,
      timeoutMs: MCP_TIMEOUT_MS,
    }));
    // Each request has a bounded lifetime; waiters cancel independently.
    const signal = undefined;
    await callJsonRpc({
      server,
      method: "initialize",
      params: {
        protocolVersion: MCP_PROTOCOL_VERSION,
        capabilities: {},
        clientInfo: {
          name: MCP_CLIENT_NAME,
          version: MCP_CLIENT_VERSION,
        },
      },
      signal,
      ...proxyBaseUrlArg(proxyBaseUrl),
    });

    await callJsonRpc({
      server,
      method: "notifications/initialized",
      signal,
      ...proxyBaseUrlArg(proxyBaseUrl),
      expectResponse: false,
    });

    const listResult = await callJsonRpc({
      server,
      method: "tools/list",
      params: {},
      signal,
      ...proxyBaseUrlArg(proxyBaseUrl),
    });

    if (!listResult) {
      throw new Error("MCP tools/list returned no response.");
    }

    const tools = parseToolListResult(server, listResult.result);
    const entry: ServerToolList = {
      server,
      tools,
      proxied: listResult.proxied,
      ...proxyBaseUrlArg(listResult.proxyBaseUrl),
      session,
      callJsonRpc,
    };
    return entry;
  };

  const ensureServerTools = async (args: {
    server: McpServerConfig;
    proxyBaseUrl?: string;
    signal: AbortSignal | undefined;
    refresh?: boolean;
  }): Promise<ServerToolList> => {
    const { server, proxyBaseUrl, signal, refresh = false } = args;
    signal?.throwIfAborted();
    const pending = pendingTools.get(server.id);
    if (pending && matchesConnection(pending, server, proxyBaseUrl)) {
      return waitForServerTools(pending.promise, signal);
    }
    const cached = toolCache.get(server.id);
    if (!refresh && cached && !cached.session.expired && matchesConnection(cached, server, proxyBaseUrl)) {
      return cached;
    }
    toolCache.delete(server.id);
    const connectionServer = { ...server };
    const request: PendingServerTools = {
      server: connectionServer,
      ...proxyBaseUrlArg(proxyBaseUrl),
      promise: discoverServerTools(connectionServer, proxyBaseUrl),
    };
    pendingTools.set(server.id, request);
    void request.promise.then((entry) => {
      if (pendingTools.get(server.id) === request) {
        toolCache.set(server.id, entry);
        pendingTools.delete(server.id);
      }
    }, () => {
      if (pendingTools.get(server.id) === request) pendingTools.delete(server.id);
    });
    return waitForServerTools(request.promise, signal);
  };

  const listToolsAcrossServers = async (args: {
    servers: readonly McpServerConfig[];
    proxyBaseUrl?: string;
    signal: AbortSignal | undefined;
  }): Promise<ServerToolList[]> => {
    const out: ServerToolList[] = [];

    for (const server of args.servers) {
      if (!server.enabled) continue;
      const list = await ensureServerTools({
        server,
        ...proxyBaseUrlArg(args.proxyBaseUrl),
        signal: args.signal,
      });
      out.push(list);
    }

    return out;
  };

  return {
    name: "mcp",
    label: "MCP Gateway",
    description:
      "Connect to configured MCP servers, discover tools, and call MCP tools.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      rawParams: unknown,
      signal: AbortSignal | undefined,
    ): Promise<AgentToolResult<McpGatewayDetails>> => {
      const params = parseParams(rawParams);
      let usedProxyBaseUrl: string | undefined;

      try {
        const runtimeConfig = await getRuntimeConfig();
        usedProxyBaseUrl = runtimeConfig.proxyBaseUrl;
        const enabledServers = runtimeConfig.servers.filter((server) => server.enabled);

        if (runtimeConfig.servers.length === 0) {
          throw new Error(`No MCP servers configured. Open ${integrationsCommandHint()} to add one.`);
        }

        const resolveSingleServer = (token: string): McpServerConfig => {
          const server = findServerByToken(runtimeConfig.servers, token);
          if (!server) {
            throw new Error(`Unknown MCP server: ${token}`);
          }
          if (!server.enabled) {
            throw new Error(`MCP server "${server.name}" is disabled.`);
          }
          return server;
        };

        if (params.connect) {
          const server = resolveSingleServer(params.connect);
          const list = await ensureServerTools({
            server,
            ...proxyBaseUrlArg(runtimeConfig.proxyBaseUrl),
            signal,
            refresh: true,
          });

          const text = [
            `Connected to MCP server \"${server.name}\" (${server.url}).`,
            `Discovered ${list.tools.length} tool${list.tools.length === 1 ? "" : "s"}.`,
            "",
            buildToolPreview(list.tools),
          ].join("\n");

          return {
            content: [{ type: "text", text }],
            details: {
              kind: "mcp_gateway",
              ok: true,
              operation: "connect",
              server: server.name,
              proxied: list.proxied,
              ...proxyBaseUrlArg(list.proxyBaseUrl),
              resultPreview: firstLine(text),
            },
          };
        }

        if (params.tool) {
          const candidateLists = params.server
            ? [await ensureServerTools({
              server: resolveSingleServer(params.server),
              ...proxyBaseUrlArg(runtimeConfig.proxyBaseUrl),
              signal,
            })]
            : await listToolsAcrossServers({
              servers: enabledServers,
              ...proxyBaseUrlArg(runtimeConfig.proxyBaseUrl),
              signal,
            });

          const allTools = candidateLists.flatMap((list) => list.tools);
          const matched = allTools.filter((tool) => tool.name === params.tool);

          if (matched.length === 0) {
            throw new Error(`MCP tool not found: ${params.tool}`);
          }

          if (!params.server) {
            const matchedServerIds = new Set(matched.map((tool) => tool.serverId));
            if (matchedServerIds.size > 1) {
              const serverNames = Array.from(new Set(matched.map((tool) => tool.serverName))).sort();
              throw new Error(
                `MCP tool "${params.tool}" is available on multiple servers (${serverNames.join(", ")}). Specify the server parameter.`,
              );
            }
          }

          const targetTool = matched[0];
          if (!targetTool) {
            throw new Error(`MCP tool not found: ${params.tool}`);
          }
          const targetList = candidateLists.find((list) => list.server.id === targetTool.serverId);
          if (!targetList) throw new Error(`MCP server not found: ${targetTool.serverId}`);
          const targetServer = targetList.server;
          const parsedArgs = parseCallArgs(params.args);
          signal?.throwIfAborted();
          const callResult = await targetList.callJsonRpc({
            server: targetServer,
            method: "tools/call",
            params: {
              name: targetTool.name,
              arguments: parsedArgs,
            },
            signal,
            ...proxyBaseUrlArg(runtimeConfig.proxyBaseUrl),
          });

          if (!callResult) {
            throw new Error("MCP tools/call returned no response.");
          }

          const payload = callResult.result;
          let resultText = "";

          if (isToolsMcpPayloadShape(payload) && isToolsMcpPayloadShape(payload.result)) {
            const rpcResult = payload.result;
            const contentBlocks = extractTextContentBlocks(rpcResult.content);
            if (contentBlocks.length > 0) {
              resultText = contentBlocks.join("\n\n");
            } else if (rpcResult.structuredContent !== undefined) {
              resultText = `\`\`\`json\n${formatJson(rpcResult.structuredContent)}\n\`\`\``;
            } else {
              resultText = `\`\`\`json\n${formatJson(rpcResult)}\n\`\`\``;
            }
          } else {
            resultText = `\`\`\`json\n${formatJson(payload)}\n\`\`\``;
          }

          const text = [
            "MCP tool call",
            `- server: ${targetServer.name} (${targetServer.url})`,
            `- tool: ${targetTool.name}`,
            "- arguments sent:",
            "```json",
            formatJson(parsedArgs),
            "```",
            "",
            "Result:",
            resultText,
          ].join("\n");

          return {
            content: [{ type: "text", text }],
            details: {
              kind: "mcp_gateway",
              ok: true,
              operation: "tool",
              server: targetServer.name,
              tool: targetTool.name,
              proxied: callResult.proxied,
              ...proxyBaseUrlArg(callResult.proxyBaseUrl),
              resultPreview: firstLine(resultText),
            },
          };
        }

        if (params.describe) {
          const lists = await listToolsAcrossServers({
            servers: enabledServers,
            ...proxyBaseUrlArg(runtimeConfig.proxyBaseUrl),
            signal,
          });

          const allTools = lists.flatMap((list) => list.tools);
          const candidates = params.server
            ? allTools.filter((tool) => matchesServerToken(tool, params.server ?? ""))
            : allTools;

          const target = candidates.find((tool) => tool.name === params.describe) ?? null;
          if (!target) {
            throw new Error(`MCP tool not found: ${params.describe}`);
          }

          const schemaBlock = target.inputSchema === undefined
            ? "(no input schema provided)"
            : `\`\`\`json\n${formatJson(target.inputSchema)}\n\`\`\``;

          const description = target.description ?? "(no description provided)";

          const text = [
            `MCP tool: ${target.name}`,
            `- server: ${target.serverName} (${target.serverUrl})`,
            `- description: ${description}`,
            "",
            "Input schema:",
            schemaBlock,
          ].join("\n");

          return {
            content: [{ type: "text", text }],
            details: {
              kind: "mcp_gateway",
              ok: true,
              operation: "describe",
              server: target.serverName,
              tool: target.name,
              resultPreview: firstLine(description),
            },
          };
        }

        if (params.search) {
          const lists = await listToolsAcrossServers({
            servers: enabledServers,
            ...proxyBaseUrlArg(runtimeConfig.proxyBaseUrl),
            signal,
          });

          const allTools = lists.flatMap((list) => list.tools);
          const matches = allTools.filter((tool) => matchesSearch(tool, params.search ?? ""));

          const lines: string[] = [];
          lines.push(`MCP search \"${params.search}\"`);
          lines.push("");

          if (matches.length === 0) {
            lines.push("No matching tools.");
          } else {
            const limit = Math.min(matches.length, 30);
            for (const tool of matches.slice(0, limit)) {
              const desc = tool.description ? ` — ${tool.description}` : "";
              lines.push(`- ${tool.name} (${tool.serverName})${desc}`);
            }

            if (matches.length > limit) {
              lines.push(`- … ${matches.length - limit} more`);
            }
          }

          const text = lines.join("\n");
          return {
            content: [{ type: "text", text }],
            details: {
              kind: "mcp_gateway",
              ok: true,
              operation: "search",
              resultPreview: firstLine(text),
            },
          };
        }

        if (params.server) {
          const server = resolveSingleServer(params.server);
          const list = await ensureServerTools({
            server,
            ...proxyBaseUrlArg(runtimeConfig.proxyBaseUrl),
            signal,
          });

          const text = [
            `MCP tools on ${server.name} (${server.url}):`,
            "",
            buildToolPreview(list.tools),
          ].join("\n");

          return {
            content: [{ type: "text", text }],
            details: {
              kind: "mcp_gateway",
              ok: true,
              operation: "server",
              server: server.name,
              proxied: list.proxied,
              ...proxyBaseUrlArg(list.proxyBaseUrl),
              resultPreview: firstLine(text),
            },
          };
        }

        const lines: string[] = [];
        lines.push("MCP server status:");

        for (const server of runtimeConfig.servers) {
          const cached = toolCache.get(server.id);
          const tools = cached && !cached.session.expired
            && matchesConnection(cached, server, runtimeConfig.proxyBaseUrl) ? cached.tools : null;
          lines.push(buildServerStatusLine(server, tools));
        }

        lines.push("");
        lines.push("Tip: use `connect`, `server`, `search`, `describe`, or `tool`.");

        const statusText = lines.join("\n");

        return {
          content: [{ type: "text", text: statusText }],
          details: {
            kind: "mcp_gateway",
            ok: true,
            operation: "status",
            resultPreview: firstLine(statusText),
          },
        };
      } catch (error) {
        const message = getErrorMessage(error);
        const proxyDown = isLikelyProxyConnectionError(message, usedProxyBaseUrl);
        const displayMessage = proxyDown
          ? buildProxyDownErrorMessage("MCP gateway", message)
          : `Error: ${message}`;

        const operation = params.tool
          ? "tool"
          : params.connect
            ? "connect"
            : params.describe
              ? "describe"
              : params.search
                ? "search"
                : params.server
                  ? "server"
                  : "status";

        return {
          content: [{ type: "text", text: displayMessage }],
          details: createMcpGatewayDetails({
            ok: false,
            operation,
            server: params.server ?? params.connect,
            tool: params.tool ?? params.describe,
            proxyBaseUrl: usedProxyBaseUrl,
            error: message,
            proxyDown,
          }),
        };
      }
    },
  };
}
