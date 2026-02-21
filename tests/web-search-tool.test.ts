import assert from "node:assert/strict";
import { test } from "node:test";

import { createWebSearchTool } from "../src/tools/web-search.ts";

void test("web_search falls back to Jina when key-required provider is missing an API key", async () => {
  const calledProviders: string[] = [];

  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ provider: "serper", apiKey: undefined }),
    executeSearch: (_params, config) => {
      calledProviders.push(config.provider);
      return Promise.resolve({
        sentQuery: "latest inflation data",
        proxied: false,
        hits: [
          {
            title: "Inflation summary",
            url: "https://example.com/inflation",
            snippet: "Fallback result via Jina.",
          },
        ],
      });
    },
  });

  const result = await tool.execute("call-1", { query: "latest inflation data" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /used Jina Search/i);
  assert.match(text, /Web search via Jina Search/);
  assert.deepEqual(calledProviders, ["jina"]);

  const details = result.details as {
    ok?: boolean;
    provider?: string;
    fallback?: { fromProvider?: string; toProvider?: string; reason?: string };
  };

  assert.equal(details.ok, true);
  assert.equal(details.provider, "jina");
  assert.equal(details.fallback?.fromProvider, "serper");
  assert.equal(details.fallback?.toProvider, "jina");
  assert.match(details.fallback?.reason ?? "", /api key is missing/i);
});

void test("web_search renders compact cited results for serper", async () => {
  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ provider: "serper", apiKey: "token" }),
    executeSearch: () => {
      return Promise.resolve({
        sentQuery: "latest cpi (site:bls.gov)",
        proxied: false,
        hits: [
          {
            title: "Consumer Price Index Summary",
            url: "https://www.bls.gov/news.release/cpi.nr0.htm",
            snippet: "Monthly CPI release from U.S. BLS.",
          },
          {
            title: "Inflation data explorer",
            url: "https://example.com/cpi",
            snippet: "Interactive inflation explorer.",
          },
        ],
      });
    },
  });

  const result = await tool.execute("call-2", {
    query: "latest cpi",
    site: ["bls.gov"],
    max_results: 2,
  });

  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /Web search via Serper\.dev/);
  assert.match(text, /\[1\] \[Consumer Price Index Summary\]/);
  assert.match(text, /\[2\] \[Inflation data explorer\]/);
  assert.ok(result.details);

  const details = result.details as { ok?: boolean; resultCount?: number; maxResults?: number; provider?: string };
  assert.equal(details.ok, true);
  assert.equal(details.provider, "serper");
  assert.equal(details.resultCount, 2);
  assert.equal(details.maxResults, 2);
});

void test("web_search falls back to Jina when configured provider returns auth or rate-limit errors", async () => {
  const calledProviders: string[] = [];

  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ provider: "tavily", apiKey: "tv-key" }),
    executeSearch: (_params, config) => {
      calledProviders.push(config.provider);

      if (config.provider === "tavily") {
        return Promise.reject(new Error("429 Too Many Requests: rate limit exceeded"));
      }

      return Promise.resolve({
        sentQuery: "excel volatility functions",
        proxied: false,
        hits: [
          {
            title: "Volatile functions in Excel",
            url: "https://example.com/volatile",
            snippet: "NOW, TODAY, RAND and more.",
          },
        ],
      });
    },
  });

  const result = await tool.execute("call-fallback-429", { query: "excel volatility functions" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /Tavily search failed/i);
  assert.match(text, /used Jina Search/i);
  assert.deepEqual(calledProviders, ["tavily", "jina"]);

  const details = result.details as {
    ok?: boolean;
    provider?: string;
    fallback?: { fromProvider?: string; toProvider?: string; reason?: string };
  };

  assert.equal(details.ok, true);
  assert.equal(details.provider, "jina");
  assert.equal(details.fallback?.fromProvider, "tavily");
  assert.equal(details.fallback?.toProvider, "jina");
  assert.match(details.fallback?.reason ?? "", /429/i);
});

void test("web_search fallback uses configured Jina API key when available", async () => {
  const calls: Array<{ provider: string; apiKey: string }> = [];

  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({
      provider: "serper",
      apiKey: "serper-key",
      jinaApiKey: "jina-fallback-key",
    }),
    executeSearch: (_params, config) => {
      calls.push({ provider: config.provider, apiKey: config.apiKey });

      if (config.provider === "serper") {
        return Promise.reject(new Error("Serper.dev search request failed (401): invalid API key"));
      }

      return Promise.resolve({
        sentQuery: "excel shortcuts",
        proxied: false,
        hits: [{
          title: "Excel keyboard shortcuts",
          url: "https://example.com/shortcuts",
          snippet: "Fallback via authenticated Jina.",
        }],
      });
    },
  });

  const result = await tool.execute("call-fallback-jina-key", { query: "excel shortcuts" });
  const details = result.details as {
    ok?: boolean;
    provider?: string;
    fallback?: { fromProvider?: string; toProvider?: string };
  };

  assert.equal(details.ok, true);
  assert.equal(details.provider, "jina");
  assert.equal(details.fallback?.fromProvider, "serper");
  assert.equal(details.fallback?.toProvider, "jina");
  assert.deepEqual(calls, [
    { provider: "serper", apiKey: "serper-key" },
    { provider: "jina", apiKey: "jina-fallback-key" },
  ]);
});

void test("web_search does not fall back for malformed-provider requests", async () => {
  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ provider: "serper", apiKey: "token" }),
    executeSearch: () => Promise.reject(new Error("Serper.dev search request failed (400): invalid request body")),
  });

  const result = await tool.execute("call-no-fallback-400", { query: "excel formula" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /^Error: /);
  assert.match(text, /400/i);
  assert.doesNotMatch(text, /used Jina Search/i);

  const details = result.details as { ok?: boolean; provider?: string; fallback?: unknown };
  assert.equal(details.ok, false);
  assert.equal(details.provider, "serper");
  assert.equal(details.fallback, undefined);
});

void test("web_search works without API key for jina (zero-config)", async () => {
  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ provider: "jina", apiKey: undefined }),
    executeSearch: () => Promise.resolve({
      sentQuery: "excel vlookup",
      proxied: false,
      hits: [
        {
          title: "VLOOKUP function",
          url: "https://support.microsoft.com/vlookup",
          snippet: "Use VLOOKUP to find things in a table.",
        },
      ],
    }),
  });

  const result = await tool.execute("call-jina-1", { query: "excel vlookup", max_results: 1 });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /Web search via Jina Search/);
  assert.match(text, /\[1\] \[VLOOKUP function\]/);

  const details = result.details as { ok?: boolean; provider?: string; resultCount?: number };
  assert.equal(details.ok, true);
  assert.equal(details.provider, "jina");
  assert.equal(details.resultCount, 1);
});

void test("web_search keeps provider metadata for brave responses", async () => {
  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ provider: "brave", apiKey: "token", proxyBaseUrl: "https://localhost:3003" }),
    executeSearch: () => Promise.resolve({
      sentQuery: "excel shortcuts",
      proxied: true,
      proxyBaseUrl: "https://localhost:3003",
      hits: [
        {
          title: "Excel keyboard shortcuts",
          url: "https://support.microsoft.com/shortcuts",
          snippet: "Official shortcut list.",
        },
      ],
    }),
  });

  const result = await tool.execute("call-3", { query: "excel shortcuts", max_results: 1 });
  const details = result.details as { provider?: string; proxied?: boolean; proxyBaseUrl?: string };

  assert.equal(details.provider, "brave");
  assert.equal(details.proxied, true);
  assert.equal(details.proxyBaseUrl, "https://localhost:3003");
});

void test("web_search reports proxy-down error when proxy is unreachable", async () => {
  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ provider: "jina", apiKey: undefined, proxyBaseUrl: "https://localhost:3003" }),
    executeSearch: () => Promise.reject(new TypeError("Load failed")),
  });

  const result = await tool.execute("call-proxy-down", { query: "test query" });
  const text = result.content[0];
  assert.equal(text.type, "text");
  assert.ok((text as { text: string }).text.includes("local CORS proxy is not running"));
  assert.ok((text as { text: string }).text.includes("npx pi-for-excel-proxy"));
  assert.ok((text as { text: string }).text.includes("Do not retry"));

  const details = result.details as { ok?: boolean; proxyDown?: boolean; error?: string };
  assert.equal(details.ok, false);
  assert.equal(details.proxyDown, true);
});

void test("web_search does not flag proxyDown when proxy is not configured", async () => {
  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ provider: "jina", apiKey: undefined }),
    executeSearch: () => Promise.reject(new TypeError("Load failed")),
  });

  const result = await tool.execute("call-no-proxy", { query: "test query" });
  const text = result.content[0];
  assert.equal(text.type, "text");
  assert.ok((text as { text: string }).text.startsWith("Error:"));

  const details = result.details as { ok?: boolean; proxyDown?: boolean };
  assert.equal(details.ok, false);
  assert.equal(details.proxyDown, false);
});
