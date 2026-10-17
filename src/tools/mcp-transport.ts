/** Streamable HTTP transport shared by the MCP gateway and connection probe. */
import { Type, type Static } from "typebox";
import { Value } from "typebox/value";

import { parseJson, type JsonValue } from "../utils/json.js";
import { getHttpErrorReason, runWithTimeoutAbort } from "../utils/network.js";
import { resolveOutboundRequestUrl } from "./external-fetch.js";
import type { McpServerConfig } from "./mcp-config.js";

export const MCP_ACCEPT_HEADER = "application/json, text/event-stream";
export const MCP_SESSION_ID_HEADER = "Mcp-Session-Id";
export const MCP_PROTOCOL_VERSION = "2025-03-26";

const jsonValueSchema = Type.Cyclic({
  JsonValue: Type.Union([
    Type.String(), Type.Number(), Type.Boolean(), Type.Null(),
    Type.Array(Type.Ref("JsonValue")),
    Type.Record(Type.String(), Type.Ref("JsonValue")),
  ]),
}, "JsonValue");
const resultSchema = Type.Record(Type.String(), jsonValueSchema);
const responseIdSchema = Type.Union([Type.String(), Type.Number(), Type.Null()]);
const responseSchema = Type.Union([
  Type.Object({
    jsonrpc: Type.Literal("2.0"),
    id: responseIdSchema,
    result: resultSchema,
    error: Type.Optional(Type.Never()),
    method: Type.Optional(Type.Never()),
  }),
  Type.Object({
    jsonrpc: Type.Literal("2.0"),
    id: responseIdSchema,
    error: Type.Object({ code: Type.Integer(), message: Type.String() }),
    result: Type.Optional(Type.Never()),
    method: Type.Optional(Type.Never()),
  }),
]);
type McpResponse = Static<typeof responseSchema>;

const initializeResultSchema = Type.Object({ protocolVersion: Type.Literal(MCP_PROTOCOL_VERSION) });
const toolListResultSchema = Type.Object({
  tools: Type.Array(Type.Object({ name: Type.String() })),
});

/** Owned by one connection, never shared between chats or probes. */
export interface McpSession {
  id?: string;
  protocolVersion?: string;
  expired?: boolean;
}

export function buildMcpRequestHeaders(args: {
  token?: string;
  sessionId?: string;
  protocolVersion?: string;
}): Record<string, string> {
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    Accept: MCP_ACCEPT_HEADER,
  };
  if (args.token) headers.Authorization = `Bearer ${args.token}`;
  if (args.sessionId) headers[MCP_SESSION_ID_HEADER] = args.sessionId;
  if (args.protocolVersion) headers["MCP-Protocol-Version"] = args.protocolVersion;
  return headers;
}

/** Notifications and server requests (even with the same id) are not answers. */
function selectResponse(payload: JsonValue, requestId: string): McpResponse | null {
  const messages = Array.isArray(payload) ? payload : [payload];
  for (const message of messages) {
    if (Value.Check(responseSchema, message) && message.id === requestId) return message;
  }
  return null;
}

/** Keeps line and CRLF state across byte chunks; dispatches only complete events. */
class McpEventStream {
  private line = "";
  private skipLf = false;
  private data: string[] = [];

  private readonly requestId: string;

  constructor(requestId: string) {
    this.requestId = requestId;
  }

  push(text: string): McpResponse | null {
    for (const character of text) {
      if (this.skipLf) {
        this.skipLf = false;
        if (character === "\n") continue;
      }
      if (character !== "\r" && character !== "\n") {
        this.line += character;
        continue;
      }
      this.skipLf = character === "\r";
      const line = this.line;
      this.line = "";
      if (line === "") {
        const data = this.data.join("\n");
        this.data = [];
        if (data === "") continue;
        let payload: JsonValue;
        try {
          payload = parseJson(data);
        } catch {
          // Ignore non-JSON events; only a matching JSON-RPC response completes a POST.
          continue;
        }
        const response = selectResponse(payload, this.requestId);
        if (response) return response;
      } else if (line === "data" || line.startsWith("data:")) {
        this.data.push(line.slice(5).replace(/^ /, ""));
      }
    }
    return null;
  }
}

/** Return at the matching response, even if the server keeps its SSE stream open. */
export async function readMcpResponse(response: Response, requestId: string): Promise<McpResponse> {
  const contentType = response.headers.get("Content-Type")?.split(";")[0]?.trim().toLowerCase();
  if (contentType !== "text/event-stream") {
    const body = await response.text();
    const payload = body.trim() ? selectResponse(parseJson(body), requestId) : null;
    if (!payload) throw new Error("Invalid MCP JSON-RPC response to this request.");
    return payload;
  }

  const reader = response.body?.getReader();
  if (!reader) throw new Error("MCP server sent an empty event stream.");
  const decoder = new TextDecoder();
  const stream = new McpEventStream(requestId);
  try {
    while (true) {
      const { value, done } = await reader.read();
      const payload = stream.push(done ? decoder.decode() : decoder.decode(value, { stream: true }));
      if (payload) return payload;
      if (done) throw new Error("MCP server sent an event stream without a JSON-RPC response to this request.");
    }
  } finally {
    // Do not wait for a remote stream to close, or let cleanup replace its result/error.
    void reader.cancel().catch(() => {});
    reader.releaseLock();
  }
}

export interface McpRpcResult {
  result: McpResponse | null;
  proxied: boolean;
  proxyBaseUrl?: string;
}

export async function postMcpJsonRpc(args: {
  server: McpServerConfig;
  method: string;
  params?: JsonValue;
  signal: AbortSignal | undefined;
  proxyBaseUrl?: string;
  expectResponse?: boolean;
  session: McpSession;
  timeoutMs: number;
}): Promise<McpRpcResult> {
  const { server, method, params, signal, session, timeoutMs, expectResponse = true } = args;
  const resolved = resolveOutboundRequestUrl({
    targetUrl: server.url,
    ...(args.proxyBaseUrl !== undefined ? { proxyBaseUrl: args.proxyBaseUrl } : {}),
  });
  const sessionId = session.id;
  const headers = buildMcpRequestHeaders({
    ...(server.token ? { token: server.token } : {}),
    ...(sessionId !== undefined ? { sessionId } : {}),
    ...(session.protocolVersion !== undefined ? { protocolVersion: session.protocolVersion } : {}),
  });
  const requestId = expectResponse ? crypto.randomUUID() : undefined;
  return runWithTimeoutAbort({
    signal,
    timeoutMs,
    timeoutErrorMessage: `MCP request timed out after ${timeoutMs}ms.`,
    run: async (requestSignal) => {
      requestSignal.throwIfAborted();
      const response = await fetch(resolved.requestUrl, {
        method: "POST",
        headers,
        body: JSON.stringify({ jsonrpc: "2.0", method, params, id: requestId }),
        signal: requestSignal,
      });
      if (!response.ok) {
        if (response.status === 404 && sessionId !== undefined) session.expired = true;
        const reason = getHttpErrorReason(response.status, await response.text());
        throw new Error(`MCP request failed (${response.status}): ${reason}`);
      }

      let payload: McpResponse | null = null;
      if (requestId === undefined) {
        void response.body?.cancel().catch(() => {});
      } else {
        payload = await readMcpResponse(response, requestId);
        if (payload.error) throw new Error(payload.error.message);
        if (method === "initialize") {
          if (!Value.Check(initializeResultSchema, payload.result)) {
            throw new Error(`MCP server did not negotiate supported protocol ${MCP_PROTOCOL_VERSION}.`);
          }
          session.protocolVersion = payload.result.protocolVersion;
          const assignedSessionId = response.headers.get(MCP_SESSION_ID_HEADER);
          if (assignedSessionId) session.id = assignedSessionId;
        }
        if (method === "tools/list" && !Value.Check(toolListResultSchema, payload.result)) {
          throw new Error("Invalid MCP tools/list result.");
        }
      }
      return {
        result: payload,
        proxied: resolved.proxied,
        ...(resolved.proxyBaseUrl !== undefined ? { proxyBaseUrl: resolved.proxyBaseUrl } : {}),
      };
    },
  });
}
