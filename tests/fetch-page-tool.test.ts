import assert from "node:assert/strict";
import { test } from "node:test";

import { createFetchPageTool } from "../src/tools/fetch-page.ts";

void test("fetch_page returns readable content and metadata", async () => {
  const tool = createFetchPageTool({
    getConfig: () => Promise.resolve({ proxyBaseUrl: "https://localhost:3003" }),
    executeFetch: () => Promise.resolve({
      status: 200,
      ok: true,
      contentType: "text/plain; charset=utf-8",
      body: "Line one\nLine two",
    }),
    now: () => 1_000,
  });

  const result = await tool.execute("call-1", { url: "https://example.com/article", max_chars: 2000 });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /Fetched page content/);
  assert.match(text, /url: https:\/\/example.com\/article/);
  assert.match(text, /Line one Line two/);

  const details = result.details as { ok?: boolean; proxied?: boolean; proxyBaseUrl?: string; chars?: number };
  assert.equal(details.ok, true);
  assert.equal(details.proxied, true);
  assert.equal(details.proxyBaseUrl, "https://localhost:3003");
  assert.ok(typeof details.chars === "number" && details.chars > 0);
});

void test("fetch_page rejects non-http urls", async () => {
  const tool = createFetchPageTool({
    getConfig: () => Promise.resolve({}),
    executeFetch: () => Promise.resolve({
      status: 200,
      ok: true,
      contentType: "text/plain",
      body: "ok",
    }),
  });

  const result = await tool.execute("call-2", { url: "file:///tmp/test.txt" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /Only http\(s\) URLs are supported/i);
  const details = result.details as { ok?: boolean };
  assert.equal(details.ok, false);
});

void test("fetch_page enforces per-domain rate limit", async () => {
  const tool = createFetchPageTool({
    getConfig: () => Promise.resolve({}),
    executeFetch: () => Promise.resolve({
      status: 200,
      ok: true,
      contentType: "text/plain",
      body: "hello",
    }),
    now: () => 5_000,
  });

  const first = await tool.execute("call-a", { url: "https://example.com/a" });
  assert.equal((first.details as { ok?: boolean }).ok, true);

  const second = await tool.execute("call-b", { url: "https://example.com/b" });
  const text = second.content[0]?.type === "text" ? second.content[0].text : "";
  assert.match(text, /Rate limited/i);
  assert.equal((second.details as { ok?: boolean }).ok, false);
});
