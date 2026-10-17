/** Real HTTP fixture. SSE replies can deliberately remain open after their answer. */
import { createServer } from "node:http";
import { once } from "node:events";

export async function startMcpServer(options = {}) {
  const requests = [];
  const sessions = new Map();
  const streams = new Set();
  let closedStreams = 0;
  const server = createServer((request, response) => {
    void handle(request, response).catch((error) => {
      response.writeHead(500).end(String(error));
    });
  });
  async function handle(request, response) {
    let text = "";
    for await (const chunk of request) text += chunk;
    const message = JSON.parse(text);
    const record = {
      ...message,
      path: request.url,
      sessionId: request.headers["mcp-session-id"],
      protocolVersion: request.headers["mcp-protocol-version"],
      authorization: request.headers.authorization,
      accept: request.headers.accept,
    };
    requests.push(record);
    const override = await options.reply?.(record) ?? {};
    if (response.destroyed) return;
    if (override.status) {
      response.writeHead(override.status).end("fixture HTTP error");
      return;
    }
    let result = {};
    const headers = { ...override.headers };
    if (message.method === "initialize") {
      if (record.sessionId) throw new Error("initialize must not reuse a session");
      const sessionId = `fixture-session-${sessions.size + 1}`;
      sessions.set(sessionId, { authorization: record.authorization, path: record.path, initialized: false });
      record.assignedSessionId = sessionId;
      headers["Mcp-Session-Id"] = sessionId;
      result = { protocolVersion: "2025-03-26", capabilities: { tools: {} }, serverInfo: { name: "fixture", version: "1" } };
    } else {
      const session = sessions.get(record.sessionId);
      if (!session || session.authorization !== record.authorization || session.path !== record.path) {
        response.writeHead(404).end("unknown session");
        return;
      }
      if (message.method === "notifications/initialized") {
        session.initialized = true;
        response.writeHead(202, headers).end();
        return;
      }
      if (!session.initialized) throw new Error("request before initialized notification");
      if (message.method === "tools/list") {
        result = { tools: [{ name: "echo", description: "Echo the supplied text unchanged.", inputSchema: {
          type: "object", properties: { text: { type: "string" } }, required: ["text"],
        } }] };
      } else if (message.method === "tools/call") {
        result = { content: [{ type: "text", text: message.params.arguments.text }] };
      } else {
        throw new Error(`Unexpected method ${message.method}`);
      }
    }
    const answer = override.error
      ? { jsonrpc: "2.0", id: message.id, error: override.error }
      : { jsonrpc: "2.0", id: message.id, result: override.result ?? result };
    if (options.format === "sse-batch") {
      response.writeHead(200, { ...headers, "Content-Type": "text/event-stream" });
      response.write('event: message\r\ndata: {"jsonrpc":"2.0","method":"notifications/message"}\r\n\r\n');
      response.write(`data: ${JSON.stringify({ jsonrpc: "2.0", id: message.id, method: "sampling/createMessage" })}\r\n\r\n`);
      response.write(`data: ${JSON.stringify([{ jsonrpc: "2.0", id: "other", result: {} }, answer])}\r\n\r\n`);
      streams.add(response);
      response.on("close", () => { streams.delete(response); closedStreams++; });
      if (!options.keepOpen) response.end();
    } else {
      response.writeHead(200, { ...headers, "Content-Type": "application/json" }).end(JSON.stringify(answer));
    }
  }
  server.listen(options.port ?? 0, "127.0.0.1");
  await once(server, "listening");
  const address = server.address();
  if (!address || typeof address === "string") throw new Error("Missing fixture TCP address");
  return {
    url: `http://127.0.0.1:${address.port}/mcp`,
    requests,
    sessions,
    get openStreams() { return streams.size; },
    get closedStreams() { return closedStreams; },
    async close() {
      server.closeAllConnections();
      await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
    },
  };
}
