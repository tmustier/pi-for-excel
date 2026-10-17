import assert from "node:assert/strict";
import { test } from "node:test";

import {
  buildMcpRequestHeaders,
  MCP_ACCEPT_HEADER,
  readMcpResponse,
} from "../src/tools/mcp-transport.ts";

const answer = { jsonrpc: "2.0", id: "req", result: { tools: [{ name: "搜索🦊" }] } };
const json = JSON.stringify(answer);

void test("MCP requests accept JSON and SSE with optional auth, session and protocol headers", () => {
  assert.deepEqual(buildMcpRequestHeaders({}), {
    "Content-Type": "application/json",
    Accept: "application/json, text/event-stream",
  });
  const headers = buildMcpRequestHeaders({ token: "secret", sessionId: "session-1", protocolVersion: "2025-03-26" });
  assert.equal(headers.Accept, MCP_ACCEPT_HEADER);
  assert.equal(headers.Authorization, "Bearer secret");
  assert.equal(headers["Mcp-Session-Id"], "session-1");
  assert.equal(headers["MCP-Protocol-Version"], "2025-03-26");
});

void test("a plain JSON response body is parsed as before", async () => {
  assert.deepEqual(await readMcpResponse(new Response(json), "req"), answer);
});

void test("JSON batches select only the matching response envelope", async () => {
  const messages = [
    { jsonrpc: "2.0", id: "req", method: "sampling/createMessage", params: {} },
    { jsonrpc: "2.0", id: "other", result: {} },
    answer,
  ];
  assert.deepEqual(await readMcpResponse(new Response(JSON.stringify(messages)), "req"), answer);
});

void test("empty, mismatched and invalid JSON responses fail rather than reporting success", async () => {
  for (const body of ["", "  ", "null", "[]", '{"jsonrpc":"2.0","id":"other","result":{}}',
    '{"jsonrpc":"1.0","id":"req","result":{}}',
    '{"jsonrpc":"2.0","id":"req","result":{},"error":{"code":1,"message":"bad"}}',
    '{"jsonrpc":"2.0","id":"req","method":"request","result":{}}',
    '{"jsonrpc":"2.0","id":"req","error":{"message":"missing code"}}']) {
    await assert.rejects(readMcpResponse(new Response(body), "req"), /Invalid MCP JSON-RPC response/);
  }
});

for (const newline of ["\n", "\r\n", "\r"]) {
  void test(`SSE handles ${JSON.stringify(newline)} and UTF-8 split at every byte`, async () => {
    const body = [
      ": heartbeat", "event: message", "data: not-json", "", "data", "",
      'data: {"jsonrpc":"2.0","method":"notifications/progress","params":{}}', "",
      'data: {"jsonrpc":"2.0","id":"req","method":"sampling/createMessage"}', "",
      'data: [{"jsonrpc":"2.0","id":"wrong","result":{}},',
      `data: ${json}]`, "", "",
    ].join(newline);
    const bytes = new TextEncoder().encode(body);
    let offset = 0;
    let cancelled = false;
    const stream = new ReadableStream<Uint8Array>({
      pull(controller) {
        if (offset < bytes.length) controller.enqueue(bytes.slice(offset, ++offset));
        // Intentionally never close: completing must not require EOF.
      },
      cancel() { cancelled = true; },
    });
    const response = new Response(stream, { headers: { "Content-Type": "text/event-stream; charset=utf-8" } });
    assert.deepEqual(await readMcpResponse(response, "req"), answer);
    assert.equal(cancelled, true);
    assert.equal(stream.locked, false);
  });
}

void test("SSE error responses are selected after leading notifications", async () => {
  const body = 'data: {"jsonrpc":"2.0","method":"notifications/message"}\n\n'
    + 'data: {"jsonrpc":"2.0","id":"req",\ndata: "error":{"code":-32601,"message":"nope"}}\n\n';
  assert.deepEqual(await readMcpResponse(new Response(body, { headers: { "Content-Type": "text/event-stream" } }), "req"), {
    jsonrpc: "2.0", id: "req", error: { code: -32601, message: "nope" },
  });
});

void test("SSE EOF without a matching complete event is an error", async () => {
  for (const body of ["", `data: ${json}\n`, 'data: {"jsonrpc":"2.0","method":"notifications/progress"}\n\n']) {
    await assert.rejects(readMcpResponse(new Response(body, { headers: { "Content-Type": "text/event-stream" } }), "req"),
      /event stream without a JSON-RPC response/);
  }
});

void test("a hanging stream cancellation cannot delay a successful answer", async () => {
  const stream = new ReadableStream<Uint8Array>({
    start(controller) { controller.enqueue(new TextEncoder().encode(`data: ${json}\n\n`)); },
    cancel() { return new Promise<void>(() => {}); },
  });
  assert.deepEqual(await readMcpResponse(new Response(stream, { headers: { "Content-Type": "text/event-stream" } }), "req"), answer);
  assert.equal(stream.locked, false);
});

void test("stream read errors survive cleanup and release the reader lock", async () => {
  const stream = new ReadableStream<Uint8Array>({
    start(controller) { controller.error(new Error("broken stream")); },
  });
  await assert.rejects(readMcpResponse(new Response(stream, { headers: { "Content-Type": "text/event-stream" } }), "req"), /broken stream/);
  assert.equal(stream.locked, false);
});
