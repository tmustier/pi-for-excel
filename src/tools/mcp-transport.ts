/**
 * Streamable HTTP request and response handling shared by the MCP tool and
 * the Extensions Hub probe.
 *
 * The MCP Streamable HTTP transport requires every POST to accept both
 * `application/json` and `text/event-stream`, and lets the server answer with
 * either. Servers built on the official SDKs reject `Accept: application/json`
 * alone with HTTP 406, and many answer a JSON accepting client with an SSE
 * stream, so the client must read both.
 */

/** Accept header value the Streamable HTTP transport requires on every POST. */
export const MCP_ACCEPT_HEADER = "application/json, text/event-stream";

/** Response header a server uses to assign a session; the client echoes it on later requests. */
export const MCP_SESSION_ID_HEADER = "Mcp-Session-Id";

/** Build the headers for one MCP JSON-RPC POST. */
export function buildMcpRequestHeaders(args: {
  token?: string;
  sessionId?: string;
}): Record<string, string> {
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    Accept: MCP_ACCEPT_HEADER,
  };

  if (args.token) {
    headers.Authorization = `Bearer ${args.token}`;
  }

  if (args.sessionId) {
    headers[MCP_SESSION_ID_HEADER] = args.sessionId;
  }

  return headers;
}

function isEventStream(contentType: string | null): boolean {
  return (contentType ?? "").toLowerCase().includes("text/event-stream");
}

function readSseDataFields(body: string): string[] {
  const events = body.replace(/\r\n?/g, "\n").split(/\n\n+/);
  const fields: string[] = [];

  for (const event of events) {
    const dataLines = event
      .split("\n")
      .filter((line) => line.startsWith("data:"))
      .map((line) => line.slice(5).replace(/^ /, ""));

    if (dataLines.length > 0) {
      fields.push(dataLines.join("\n"));
    }
  }

  return fields;
}

function isJsonRpcMessageFor(
  value: Record<string, unknown>,
  requestId: string | undefined,
): boolean {
  if (requestId !== undefined) {
    return value.id === requestId;
  }
  return "result" in value || "error" in value;
}

/**
 * Parse an MCP JSON-RPC response body sent as JSON or as an SSE stream.
 * For SSE, returns the message that answers `requestId` and skips server
 * notifications sent ahead of it.
 */
export function parseMcpResponseBody(
  body: string,
  contentType: string | null,
  requestId: string | undefined,
): Record<string, unknown> | null {
  if (body.trim().length === 0) {
    return null;
  }

  if (!isEventStream(contentType)) {
    const payload: unknown = JSON.parse(body);
    return typeof payload === "object" && payload !== null && !Array.isArray(payload)
      ? { ...payload }
      : null;
  }

  for (const field of readSseDataFields(body)) {
    let message: unknown;
    try {
      message = JSON.parse(field);
    } catch {
      continue;
    }

    if (typeof message !== "object" || message === null || Array.isArray(message)) {
      continue;
    }

    const record = { ...message };
    if (isJsonRpcMessageFor(record, requestId)) {
      return record;
    }
  }

  throw new Error("MCP server sent an event stream without a JSON-RPC response to this request.");
}
