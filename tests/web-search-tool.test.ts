import assert from "node:assert/strict";
import { test } from "node:test";

import { createWebSearchTool } from "../src/tools/web-search.ts";

void test("web_search reports missing API key for key-required provider", async () => {
  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ provider: "serper", apiKey: undefined }),
  });

  const result = await tool.execute("call-1", { query: "latest inflation data" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /API key is missing/i);
  assert.match(text, /Serper API key/i);
  assert.ok(result.details);
  assert.equal((result.details as { ok?: boolean }).ok, false);
  assert.equal((result.details as { provider?: string }).provider, "serper");
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
