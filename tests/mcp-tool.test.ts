import assert from "node:assert/strict";
import { test } from "node:test";

import { createMcpTool, type McpToolDependencies } from "../src/tools/mcp.ts";

const TEST_SERVER = {
  id: "srv.local",
  name: "local",
  url: "https://localhost:4010/mcp",
  enabled: true,
} as const;

function createMockMcpTool() {
  const calls: Array<{ method: string; params?: unknown }> = [];

  const tool = createMcpTool({
    requestToolApproval: () => Promise.resolve(true),
    getRuntimeConfig: () => Promise.resolve({
      servers: [TEST_SERVER],
      proxyBaseUrl: undefined,
    }),
    callJsonRpc: ({ method, params }) => {
      calls.push({ method, params });

      if (method === "initialize") {
        return Promise.resolve({
          result: { result: { protocolVersion: "2025-03-26" } },
          proxied: false,
        });
      }

      if (method === "notifications/initialized") {
        return Promise.resolve({
          result: null,
          proxied: false,
        });
      }

      if (method === "tools/list") {
        return Promise.resolve({
          result: {
            result: {
              tools: [
                {
                  name: "echo",
                  description: "Echo input text",
                  inputSchema: {
                    type: "object",
                    properties: {
                      text: { type: "string" },
                    },
                  },
                },
              ],
            },
          },
          proxied: false,
        });
      }

      if (method === "tools/call") {
        return Promise.resolve({
          result: {
            result: {
              content: [
                {
                  type: "text",
                  text: "echo: hello",
                },
              ],
            },
          },
          proxied: false,
        });
      }

      throw new Error(`Unexpected method: ${method}`);
    },
  });

  return { tool, calls };
}

void test("mcp connect refreshes and lists server tools", async () => {
  const { tool, calls } = createMockMcpTool();

  const result = await tool.execute("call-1", { connect: "local" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /Connected to MCP server/);
  assert.match(text, /Discovered 1 tool/);
  assert.ok(calls.some((call) => call.method === "initialize"));
  assert.ok(calls.some((call) => call.method === "tools\/list"));
});

void test("mcp tool call includes attribution and arguments", async () => {
  const { tool, calls } = createMockMcpTool();

  const result = await tool.execute("call-2", {
    tool: "echo",
    args: JSON.stringify({ text: "hello" }),
  });

  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /MCP tool call/);
  assert.match(text, /server: local/);
  assert.match(text, /tool: echo/);
  assert.match(text, /"text": "hello"/);
  assert.match(text, /echo: hello/);

  const toolCall = calls.find((call) => call.method === "tools/call");
  assert.ok(toolCall);
});

void test("mcp tool call requires server when tool name is ambiguous", async () => {
  const serverA = {
    id: "srv.a",
    name: "alpha",
    url: "https://alpha.example/mcp",
    enabled: true,
  } as const;

  const serverB = {
    id: "srv.b",
    name: "beta",
    url: "https://beta.example/mcp",
    enabled: true,
  } as const;

  let callInvoked = false;

  const tool = createMcpTool({
    getRuntimeConfig: () => Promise.resolve({
      servers: [serverA, serverB],
      proxyBaseUrl: undefined,
    }),
    callJsonRpc: ({ server, method }) => {
      if (method === "initialize") {
        return Promise.resolve({
          result: { result: { protocolVersion: "2025-03-26" } },
          proxied: false,
        });
      }

      if (method === "notifications/initialized") {
        return Promise.resolve({
          result: null,
          proxied: false,
        });
      }

      if (method === "tools/list") {
        return Promise.resolve({
          result: {
            result: {
              tools: [
                {
                  name: "echo",
                  description: `Echo from ${server.name}`,
                  inputSchema: { type: "object" },
                },
              ],
            },
          },
          proxied: false,
        });
      }

      if (method === "tools/call") {
        callInvoked = true;
      }

      throw new Error(`Unexpected method: ${method}`);
    },
  });

  const result = await tool.execute("call-3", { tool: "echo" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /available on multiple servers.*specify the server parameter/i);
  assert.equal(callInvoked, false);
});

void test("mcp reports proxy-down error when proxy transport is unreachable", async () => {
  const tool = createMcpTool({
    getRuntimeConfig: () => Promise.resolve({
      servers: [TEST_SERVER],
      proxyBaseUrl: "https://localhost:3003",
    }),
    callJsonRpc: () => Promise.reject(new TypeError("Load failed")),
  });

  const result = await tool.execute("call-4", { connect: "local" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /local CORS proxy is not running/i);
  assert.match(text, /npx pi-for-excel-proxy/i);
  assert.match(text, /Do not retry/i);

  const details = result.details as {
    ok?: boolean;
    operation?: string;
    server?: string;
    proxyBaseUrl?: string;
    proxyDown?: boolean;
  };
  assert.equal(details.ok, false);
  assert.equal(details.operation, "connect");
  assert.equal(details.server, "local");
  assert.equal(details.proxyBaseUrl, "https://localhost:3003");
  assert.equal(details.proxyDown, true);
});

void test("mcp does not flag proxyDown for upstream JSON-RPC errors containing fetch failed", async () => {
  const tool = createMcpTool({
    getRuntimeConfig: () => Promise.resolve({
      servers: [TEST_SERVER],
      proxyBaseUrl: "https://localhost:3003",
    }),
    callJsonRpc: () => Promise.reject(new Error("JSON-RPC error: upstream fetch failed while calling backend service")),
  });

  const result = await tool.execute("call-5", { connect: "local" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /^Error: JSON-RPC error: upstream fetch failed while calling backend service$/);

  const details = result.details as { ok?: boolean; proxyDown?: boolean };
  assert.equal(details.ok, false);
  assert.equal(details.proxyDown, false);
});

void test("MCP consent is per server/tool, once or remembered in Auto, and per call in Confirm", async () => {
  let mode: "yolo" | "safe" = "yolo";
  let remember = false;
  let allow = true;
  const grants = new Map<string, string[]>();
  const asked: string[] = [];
  const sent: string[] = [];
  const serverB = { ...TEST_SERVER, id: "other", name: "other", url: "https://localhost:4011/mcp" };
  const dependencies: McpToolDependencies = {
    getExecutionMode: () => Promise.resolve(mode),
    getRuntimeConfig: () => Promise.resolve({ servers: [TEST_SERVER, serverB] }),
    getToolGrants: (server) => Promise.resolve(grants.get(server.id) ?? []),
    grantTool: (server, tool) => { grants.set(server.id, [...(grants.get(server.id) ?? []), tool]); return Promise.resolve(); },
    requestToolApproval: async (request) => {
      asked.push(request.message);
      if (remember && allow) await request.secondaryApproval?.onApprove();
      if (mode === "safe") assert.equal(request.secondaryApproval, undefined);
      return allow;
    },
    callJsonRpc: ({ method, server, params }) => {
      if (method === "tools/list") return Promise.resolve({ proxied: false, result: { result: { tools: [
        { name: "send" }, { name: "remove", annotations: { readOnlyHint: false } },
        { name: "read", annotations: { readOnlyHint: true } },
      ] } } });
      if (method === "tools/call") sent.push(`${server.id}:${JSON.stringify(params)}`);
      return Promise.resolve({ proxied: false, result: { result: { content: [{ type: "text", text: "ok" }] } } });
    },
  };
  let tool = createMcpTool(dependencies);
  const run = (name: string, server = "local", signal?: AbortSignal) => tool.execute("call", { tool: name, server, args: '{"amount":17}' }, signal);
  await run("read");
  assert.equal(asked.length, 0);
  await run("send");
  await run("send");
  assert.equal(asked.length, 2, "Allow once does not remember consent");
  assert.match(asked[0] ?? "", /local[\s\S]*send[\s\S]*17/);
  remember = true;
  await run("send");
  tool = createMcpTool(dependencies);
  await run("send");
  assert.equal(asked.length, 3, "remembered consent survives tool recreation");
  await run("send", "other");
  await run("remove");
  assert.equal(asked.length, 5, "another server and another tool require independent consent");
  mode = "safe";
  await run("send");
  await run("send");
  await run("read");
  assert.equal(asked.length, 8, "Confirm ignores remembered grants and read-only hints");
  allow = false;
  const before = sent.length;
  assert.equal((await run("send")).details.ok, false);
  assert.equal(sent.length, before);
  const controller = new AbortController();
  dependencies.requestToolApproval = () => { controller.abort(); return Promise.resolve(true); };
  tool = createMcpTool(dependencies);
  assert.equal((await run("send", "local", controller.signal)).details.ok, false);
  assert.equal(sent.length, before, "Stop prevents dispatch even if approval resolves true");
  delete dependencies.requestToolApproval;
  tool = createMcpTool(dependencies);
  assert.equal((await run("send")).details.ok, false);
  assert.equal(sent.length, before, "missing approval handler fails closed");
});
