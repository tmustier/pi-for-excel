import assert from "node:assert/strict";
import { test } from "node:test";

import { createMcpTool } from "../src/tools/mcp.ts";

const TEST_SERVER = {
  id: "srv.local",
  name: "local",
  url: "https://localhost:4010/mcp",
  enabled: true,
} as const;

function createMockMcpTool() {
  const calls: Array<{ method: string; params?: unknown }> = [];

  const tool = createMcpTool({
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
