import assert from "node:assert/strict";
import { test } from "node:test";

import { createFetchPageTool } from "../src/tools/fetch-page.ts";
import { createMcpTool } from "../src/tools/mcp.ts";
import { createModifyStructureTool } from "../src/tools/modify-structure.ts";
import { createWebSearchTool } from "../src/tools/web-search.ts";
import { OfficeError } from "./fixtures/excel-write-host.ts";

/** An Excel host whose structural edits fail with `failure`, as Excel on the web does on a protected sheet. */
async function withFailingStructureHost<T>(failure: OfficeError, action: () => Promise<T>): Promise<T> {
  let failNextSync = false;
  const sheet = {
    id: "sheet-1",
    name: "Protected",
    load: (): void => {},
    protection: { protected: failure.code === "AccessDenied", load: (): void => {} },
    getRange: () => ({ insert: (): void => { failNextSync = true; } }),
  };
  const context = {
    workbook: {
      worksheets: { getItem: () => sheet, getActiveWorksheet: () => sheet },
      protection: { protected: false, load: (): void => {} },
    },
    sync: (): Promise<void> => {
      if (!failNextSync) return Promise.resolve();
      failNextSync = false;
      return Promise.reject(failure);
    },
  };
  const previous = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", { run: <R>(callback: (ctx: unknown) => Promise<R>): Promise<R> => callback(context) });
  try {
    return await action();
  } finally {
    Reflect.set(globalThis, "Excel", previous);
  }
}

void test("error-path matrix: wrong web_search API key falls back to Jina with warning", async () => {
  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ provider: "serper", apiKey: "bad-key", jinaApiKey: "jina-fallback-key" }),
    executeSearch: (_params, config) => {
      if (config.provider === "serper") {
        return Promise.reject(new Error("Serper.dev search request failed (401): invalid API key"));
      }

      return Promise.resolve({
        sentQuery: "latest cpi",
        proxied: false,
        hits: [{
          title: "CPI release",
          url: "https://example.com/cpi",
          snippet: "Fallback search result.",
        }],
      });
    },
  });

  const result = await tool.execute("call-1", { query: "latest cpi" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.doesNotMatch(text, /^Error: /);
  assert.match(text, /used Jina Search/i);
  assert.match(text, /Web search via Jina Search/i);

  const details = result.details as {
    ok?: boolean;
    provider?: string;
    fallback?: { fromProvider?: string; toProvider?: string; reason?: string };
  };
  assert.equal(details.ok, true);
  assert.equal(details.provider, "jina");
  assert.equal(details.fallback?.fromProvider, "serper");
  assert.equal(details.fallback?.toProvider, "jina");
  assert.match(details.fallback?.reason ?? "", /401/i);
});

void test("error-path matrix: web_search rate-limit failures fall back to Jina", async () => {
  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ provider: "tavily", apiKey: "tv-key", jinaApiKey: "jina-fallback-key" }),
    executeSearch: (_params, config) => {
      if (config.provider === "tavily") {
        return Promise.reject(new Error("429 Too Many Requests: rate limit exceeded"));
      }

      return Promise.resolve({
        sentQuery: "fx rates",
        proxied: false,
        hits: [{
          title: "FX rates overview",
          url: "https://example.com/fx",
          snippet: "Fallback result.",
        }],
      });
    },
  });

  const result = await tool.execute("call-2", { query: "fx rates" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.doesNotMatch(text, /^Error: /);
  assert.match(text, /used Jina Search/i);
  assert.match(text, /rate limit/i);

  const details = result.details as {
    ok?: boolean;
    provider?: string;
    fallback?: { fromProvider?: string; toProvider?: string; reason?: string };
  };
  assert.equal(details.ok, true);
  assert.equal(details.provider, "jina");
  assert.equal(details.fallback?.fromProvider, "tavily");
  assert.equal(details.fallback?.toProvider, "jina");
  assert.match(details.fallback?.reason ?? "", /rate limit/i);
});

void test("error-path matrix: fetch_page reports proxy-down transport errors", async () => {
  const tool = createFetchPageTool({
    getConfig: () => Promise.resolve({ proxyBaseUrl: "https://localhost:3003" }),
    executeFetch: () => {
      return Promise.reject(new TypeError("fetch failed"));
    },
    now: () => 1_000,
  });

  const result = await tool.execute("call-3", { url: "https://proxy-down.example/docs" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /^Error: /);
  assert.match(text, /fetch failed/i);
  assert.equal(result.isError, true);

  const details = result.details as { ok?: boolean; url?: string; error?: string };
  assert.equal(details.ok, false);
  assert.equal(details.url, "https://proxy-down.example/docs");
  assert.match(details.error ?? "", /fetch failed/i);
});

void test("error-path matrix: mcp connect surfaces expired-token auth failures", async () => {
  const tool = createMcpTool({
    getRuntimeConfig: () => Promise.resolve({
      servers: [{
        id: "srv.local",
        name: "local",
        url: "https://localhost:4010/mcp",
        enabled: true,
      }],
      proxyBaseUrl: undefined,
    }),
    callJsonRpc: () => {
      return Promise.reject(new Error("401 Unauthorized: token expired"));
    },
  });

  const result = await tool.execute("call-4", { connect: "local" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /^Error: /);
  assert.match(text, /401/i);
  assert.match(text, /token expired/i);

  const details = result.details as { ok?: boolean; operation?: string; server?: string; error?: string };
  assert.equal(details.ok, false);
  assert.equal(details.operation, "connect");
  assert.equal(details.server, "local");
  assert.match(details.error ?? "", /token expired/i);
});

void test("error-path matrix: mcp tool call handles mid-call network disconnect", async () => {
  const server = {
    id: "srv.local",
    name: "local",
    url: "https://localhost:4010/mcp",
    enabled: true,
  } as const;

  const tool = createMcpTool({
    getRuntimeConfig: () => Promise.resolve({
      servers: [server],
      proxyBaseUrl: undefined,
    }),
    callJsonRpc: ({ method }) => {
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
              tools: [{
                name: "echo",
                description: "Echo input",
                inputSchema: { type: "object" },
              }],
            },
          },
          proxied: false,
        });
      }

      if (method === "tools/call") {
        return Promise.reject(new TypeError("NetworkError when attempting to fetch resource."));
      }

      return Promise.reject(new Error(`Unexpected method: ${method}`));
    },
  });

  const result = await tool.execute("call-5", {
    tool: "echo",
    args: JSON.stringify({ text: "hello" }),
  });

  const text = result.content[0]?.type === "text" ? result.content[0].text : "";
  assert.match(text, /^Error: /);
  assert.match(text, /networkerror/i);

  const details = result.details as { ok?: boolean; operation?: string; tool?: string; error?: string };
  assert.equal(details.ok, false);
  assert.equal(details.operation, "tool");
  assert.equal(details.tool, "echo");
  assert.match(details.error ?? "", /networkerror/i);
});

void test("error-path matrix: a protected sheet is named as the reason, as a tool error", async () => {
  const result = await withFailingStructureHost(
    new OfficeError("AccessDenied", "You cannot perform the requested operation.", "Range.insert"),
    () => createModifyStructureTool().execute("call-protected", { action: "insert_rows", sheet: "Protected", position: 3 }),
  );
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.equal(result.isError, true);
  assert.match(text, /the workbook was not changed/u);
  assert.match(text, /Sheet "Protected" is protected/u);
  assert.doesNotMatch(text, /You cannot perform the requested operation/u);
});

void test("error-path matrix: cell-edit mode tells the agent to ask the user to leave the cell", async () => {
  const result = await withFailingStructureHost(
    new OfficeError("InvalidOperationInCellEditMode", "Excel is in cell-editing mode.", ""),
    () => createModifyStructureTool().execute("call-edit-mode", { action: "insert_rows", sheet: "Protected", position: 3 }),
  );
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.equal(result.isError, true);
  assert.match(text, /cell-edit mode/u);
  assert.match(text, /press Enter or Esc/u);
});
