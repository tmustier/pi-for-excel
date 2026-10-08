import assert from "node:assert/strict";
import { test } from "node:test";

import {
  buildMcpRequestHeaders,
  MCP_ACCEPT_HEADER,
  parseMcpResponseBody,
} from "../src/tools/mcp-transport.ts";

void test("MCP requests accept both JSON and SSE, as Streamable HTTP requires", () => {
  const headers = buildMcpRequestHeaders({});

  assert.equal(headers.Accept, "application/json, text/event-stream");
  assert.equal(headers.Accept, MCP_ACCEPT_HEADER);
  assert.equal(headers["Content-Type"], "application/json");
  assert.equal("Authorization" in headers, false);
  assert.equal("Mcp-Session-Id" in headers, false);
});

void test("MCP request headers carry the bearer token and session id when known", () => {
  const headers = buildMcpRequestHeaders({ token: "secret", sessionId: "session-1" });

  assert.equal(headers.Authorization, "Bearer secret");
  assert.equal(headers["Mcp-Session-Id"], "session-1");
});

void test("a plain JSON response body is parsed as before", () => {
  const payload = parseMcpResponseBody(
    '{"jsonrpc":"2.0","id":"a","result":{"tools":[]}}',
    "application/json",
    "a",
  );

  assert.deepEqual(payload, { jsonrpc: "2.0", id: "a", result: { tools: [] } });
});

void test("an empty response body parses to null", () => {
  assert.equal(parseMcpResponseBody("  ", "application/json", "a"), null);
});

void test("an SSE response yields the JSON-RPC message that answers the request", () => {
  const body = [
    "event: message",
    'data: {"jsonrpc":"2.0","method":"notifications/message","params":{"level":"info"}}',
    "",
    "event: message",
    'data: {"jsonrpc":"2.0","id":"req-7","result":{"tools":[{"name":"search"}]}}',
    "",
    "",
  ].join("\n");

  const payload = parseMcpResponseBody(body, "text/event-stream; charset=utf-8", "req-7");

  assert.deepEqual(payload, { jsonrpc: "2.0", id: "req-7", result: { tools: [{ name: "search" }] } });
});

void test("an SSE data field split over several lines is joined before parsing", () => {
  const body = 'id: 1\r\ndata: {"jsonrpc":"2.0","id":"x",\r\ndata: "error":{"code":-32601,"message":"nope"}}\r\n\r\n';

  const payload = parseMcpResponseBody(body, "text/event-stream", "x");

  assert.deepEqual(payload, { jsonrpc: "2.0", id: "x", error: { code: -32601, message: "nope" } });
});

void test("an SSE stream with no answer to the request is reported, not returned as empty", () => {
  const body = 'data: {"jsonrpc":"2.0","method":"notifications/progress","params":{}}\n\n';

  assert.throws(
    () => parseMcpResponseBody(body, "text/event-stream", "missing"),
    /event stream without a JSON-RPC response/,
  );
});
