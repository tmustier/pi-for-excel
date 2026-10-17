import assert from "node:assert/strict";
import { setTimeout as delay } from "node:timers/promises";
import { test } from "node:test";

import { createMcpTool } from "../src/tools/mcp.ts";
import { probeMcpServer } from "../src/commands/builtins/extensions-hub-mcp-probe.ts";
import { startMcpServer } from "./fixtures/mcp-streamable-server.mjs";

// agent-default (not a user rule): 2026-10-17 — bound deliberately unclosed-stream regressions.
const TEST_TIMEOUT_MS = 3_000;
const settings = { get: async () => undefined, set: async () => {}, delete: async () => {} };
const config = (url) => ({ id: "shared-id", name: "fixture", url, enabled: true });
const gateway = (server) => createMcpTool({ getRuntimeConfig: async () => ({ servers: [server] }) });
const call = (tool, text = "搜索🦊") => tool.execute("test-call", { tool: "echo", args: JSON.stringify({ text }) });
const connect = (tool) => tool.execute("test-connect", { connect: "fixture" });
function assertOk(result) {
  assert.equal(result.details.ok, true, JSON.stringify(result));
}
function deferred() {
  let resolve;
  const promise = new Promise((done) => { resolve = done; });
  return { promise, resolve };
}
async function fixture(t, options = {}) {
  const server = await startMcpServer(options);
  t.after(() => server.close());
  return server;
}

void test("gateway and probe finish batched SSE replies without waiting for EOF", { timeout: TEST_TIMEOUT_MS }, async (t) => {
  const server = await fixture(t, { format: "sse-batch", keepOpen: true });
  const result = await call(gateway(config(server.url)));
  assertOk(result);
  assert.match(result.content[0].text, /搜索🦊/);
  assert.deepEqual(await probeMcpServer(config(server.url), settings), { toolCount: 1, proxied: false });
  // All five response streams should be cancelled by the client, not by fixture teardown.
  for (let i = 0; server.closedStreams < 5 && i < 100; i++) await delay(1);
  assert.equal(server.closedStreams, 5);
  assert.equal(server.openStreams, 0);
  for (const request of server.requests) {
    assert.equal(request.accept, "application/json, text/event-stream");
    if (request.method !== "initialize") assert.equal(request.protocolVersion, "2025-03-26");
  }
});

void test("two chat runtimes never share sessions for the same server ID", async (t) => {
  const server = await fixture(t);
  const first = gateway(config(server.url));
  const second = gateway(config(server.url));
  assertOk(await connect(first));
  assertOk(await connect(second));
  assertOk(await call(first, "first"));
  assertOk(await call(second, "second"));
  const initializations = server.requests.filter((request) => request.method === "initialize");
  const calls = server.requests.filter((request) => request.method === "tools/call");
  assert.equal(calls[0].sessionId, initializations[0].assignedSessionId);
  assert.equal(calls[1].sessionId, initializations[1].assignedSessionId);
  assert.notEqual(calls[0].sessionId, calls[1].sessionId);
});

void test("parallel cold calls and refresh coalesce the complete handshake", async (t) => {
  const started = deferred();
  const release = deferred();
  const server = await fixture(t, { reply: async (request) => {
    if (request.method === "initialize") { started.resolve(); await release.promise; }
  } });
  t.after(() => release.resolve());
  const tool = gateway(config(server.url));
  const first = call(tool, "first");
  await started.promise;
  const second = call(tool, "second");
  const refresh = connect(tool);
  release.resolve();
  for (const result of await Promise.all([first, second, refresh])) assertOk(result);
  for (const method of ["initialize", "notifications/initialized", "tools/list"]) {
    assert.equal(server.requests.filter((request) => request.method === method).length, 1, method);
  }
  assert.equal(server.requests.filter((request) => request.method === "tools/call").length, 2);
});

void test("aborting one handshake waiter does not abort another or dispatch its tool", async (t) => {
  const started = deferred();
  const release = deferred();
  const server = await fixture(t, { reply: async (request) => {
    if (request.method === "initialize") { started.resolve(); await release.promise; }
  } });
  t.after(() => release.resolve());
  const tool = gateway(config(server.url));
  const controller = new AbortController();
  const first = tool.execute("cancel-me", { tool: "echo", args: '{"text":"cancelled"}' }, controller.signal);
  await started.promise;
  const second = call(tool, "survivor");
  controller.abort();
  const cancelled = await first;
  assert.equal(cancelled.details.ok, false);
  assert.match(cancelled.details.error, /abort/i);
  release.resolve();
  assertOk(await second);
  assert.equal(server.sessions.size, 1);
  const calls = server.requests.filter((request) => request.method === "tools/call");
  assert.deepEqual(calls.map((request) => request.params.arguments.text), ["survivor"]);
});

for (const changed of ["url", "token", "proxy"]) {
  void test(`changing ${changed} creates a new connection and refreshes cached tools`, async (t) => {
    const server = await fixture(t);
    const mutableServer = config(server.url);
    const runtime = { servers: [mutableServer] };
    const tool = createMcpTool({ getRuntimeConfig: async () => runtime });
    assertOk(await call(tool, "before"));
    if (changed === "url") mutableServer.url += "-new";
    if (changed === "token") mutableServer.token = "fixture-token";
    if (changed === "proxy") runtime.proxyBaseUrl = server.url.replace("/mcp", "");
    assertOk(await call(tool, "after"));
    assert.equal(server.sessions.size, 2);
    assert.equal(server.requests.filter((request) => request.method === "tools/list").length, 2);
    const calls = server.requests.filter((request) => request.method === "tools/call");
    assert.notEqual(calls[0].sessionId, calls[1].sessionId);
  });
}

void test("an older handshake cannot replace a newer configuration's cache or session", async (t) => {
  const started = deferred();
  const release = deferred();
  const server = await fixture(t, { reply: async (request) => {
    if (request.method === "initialize" && request.path === "/mcp") { started.resolve(); await release.promise; }
  } });
  t.after(() => release.resolve());
  let current = config(server.url);
  const tool = createMcpTool({ getRuntimeConfig: async () => ({ servers: [current] }) });
  const old = connect(tool);
  await started.promise;
  current = { ...current, url: `${server.url}-new` };
  assertOk(await connect(tool));
  release.resolve();
  assertOk(await old);
  assertOk(await call(tool));
  assert.equal(server.sessions.size, 2);
  const lastCall = server.requests.findLast((request) => request.method === "tools/call");
  const newInit = server.requests.find((request) => request.method === "initialize" && request.path === "/mcp-new");
  assert.equal(lastCall.sessionId, newInit.assignedSessionId);
});

for (const method of ["initialize", "tools/list"]) {
  void test(`probe surfaces ${method} JSON-RPC errors instead of a successful zero-tool result`, async (t) => {
    const server = await fixture(t, { reply: (request) => request.method === method
      ? { error: { code: -32603, message: "fixture RPC failure" } } : {} });
    await assert.rejects(probeMcpServer(config(server.url), settings), /fixture RPC failure/);
    assert.equal(server.requests.at(-1).method, method);
  });
}

void test("probe rejects malformed tool lists and unsupported initialization versions", async (t) => {
  const options = { reply: (request) => request.method === "tools/list" ? { result: { tools: "not-an-array" } } : {} };
  const server = await fixture(t, options);
  await assert.rejects(probeMcpServer(config(server.url), settings), /Invalid MCP tools\/list result/);
  options.reply = (request) => request.method === "initialize" ? { result: { protocolVersion: "unsupported" } } : {};
  await assert.rejects(probeMcpServer(config(server.url), settings), /did not negotiate supported protocol/);
});

void test("failed discovery is not cached and a subsequent connection can succeed", async (t) => {
  let fail = true;
  const server = await fixture(t, { reply: (request) => request.method === "tools/list" && fail
    ? { error: { code: -32603, message: "discovery failed" } } : {} });
  const tool = gateway(config(server.url));
  assert.equal((await call(tool)).details.ok, false);
  fail = false;
  assertOk(await call(tool));
  assert.equal(server.sessions.size, 2);
});

void test("only initialize can assign the session header", async (t) => {
  const server = await fixture(t, { reply: (request) => request.method !== "initialize"
    ? { headers: { "Mcp-Session-Id": "wrong-session" } } : {} });
  assertOk(await call(gateway(config(server.url))));
  const initialization = server.requests[0];
  assert.ok(server.requests.slice(1).every((request) => request.sessionId === initialization.assignedSessionId));
});

void test("an expired session is invalidated without automatically replaying tools/call", async (t) => {
  let expire = true;
  const server = await fixture(t, { reply: (request) => request.method === "tools/call" && expire ? { status: 404 } : {} });
  const tool = gateway(config(server.url));
  const failed = await call(tool, "first");
  assert.equal(failed.details.ok, false);
  assert.match(failed.details.error, /404/);
  assert.equal(server.requests.filter((request) => request.method === "tools/call").length, 1);
  expire = false;
  assertOk(await call(tool, "explicit next call"));
  assert.equal(server.sessions.size, 2);
});

void test("a stale request's 404 cannot invalidate a replacement connection", async (t) => {
  const started = deferred();
  const release = deferred();
  const server = await fixture(t, { reply: async (request) => {
    if (request.method === "tools/call" && request.params.arguments.text === "old") {
      started.resolve(); await release.promise;
      return { status: 404 };
    }
  } });
  t.after(() => release.resolve());
  const tool = gateway(config(server.url));
  const old = call(tool, "old");
  await started.promise;
  assertOk(await connect(tool));
  release.resolve();
  assert.equal((await old).details.ok, false);
  assertOk(await call(tool, "new"));
  assert.equal(server.sessions.size, 2);
});

void test("parallel probes each own a separate session", async (t) => {
  const server = await fixture(t);
  const results = await Promise.all([probeMcpServer(config(server.url), settings), probeMcpServer(config(server.url), settings)]);
  for (const result of results) assert.equal(result.toolCount, 1);
  const lists = server.requests.filter((request) => request.method === "tools/list");
  assert.notEqual(lists[0].sessionId, lists[1].sessionId);
});
