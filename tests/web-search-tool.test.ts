import assert from "node:assert/strict";
import { test } from "node:test";

import { createWebSearchTool } from "../src/tools/web-search.ts";

void test("web_search reports missing API key", async () => {
  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ apiKey: undefined }),
  });

  const result = await tool.execute("call-1", { query: "latest inflation data" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /API key is missing/i);
  assert.ok(result.details);
  assert.equal((result.details as { ok?: boolean }).ok, false);
});

void test("web_search renders compact cited results", async () => {
  const tool = createWebSearchTool({
    getConfig: () => Promise.resolve({ apiKey: "token" }),
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

  assert.match(text, /Web search via Brave Search/);
  assert.match(text, /\[1\] \[Consumer Price Index Summary\]/);
  assert.match(text, /\[2\] \[Inflation data explorer\]/);
  assert.ok(result.details);

  const details = result.details as { ok?: boolean; resultCount?: number; maxResults?: number };
  assert.equal(details.ok, true);
  assert.equal(details.resultCount, 2);
  assert.equal(details.maxResults, 2);
});
