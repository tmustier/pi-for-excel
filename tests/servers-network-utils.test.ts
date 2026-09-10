import assert from "node:assert/strict";
import { test } from "node:test";

import { createFetchPageTool, type FetchPageToolDependencies } from "../src/tools/fetch-page.ts";

function resultText(content: Awaited<ReturnType<ReturnType<typeof createFetchPageTool>["execute"]>>["content"]): string {
  const first = content[0];
  return first?.type === "text" ? first.text : "";
}

void test("fetch_page preserves HTTP, completion, timeout, and caller-abort semantics", async () => {
  const rows: ReadonlyArray<{
    name: string;
    dependencies: FetchPageToolDependencies;
    signal?: AbortSignal;
    expected: RegExp;
    ok: boolean;
  }> = [
    {
      name: "returns the network result",
      dependencies: {
        executeFetch: () => Promise.resolve({ status: 200, ok: true, contentType: "text/plain", body: "Network result" }),
      },
      expected: /Network result/u,
      ok: true,
    },
    {
      name: "prefers a non-empty HTTP response body",
      dependencies: {
        executeFetch: () => Promise.resolve({ status: 502, ok: false, contentType: "text/plain", body: " upstream failed " }),
      },
      expected: /fetch_page request failed \(502\): upstream failed/u,
      ok: false,
    },
    {
      name: "times out a pending request",
      dependencies: {
        timeoutMs: 5,
        executeFetch: () => new Promise(() => {}),
      },
      expected: /fetch_page timed out after 5ms/u,
      ok: false,
    },
    {
      name: "preserves caller abort semantics",
      dependencies: {
        executeFetch: () => new Promise(() => {}),
      },
      signal: AbortSignal.abort(),
      expected: /^Error: Aborted$/u,
      ok: false,
    },
  ];

  for (const [index, row] of rows.entries()) {
    const tool = createFetchPageTool({
      getConfig: () => Promise.resolve({}),
      now: () => 10_000,
      ...row.dependencies,
    });
    const result = await tool.execute(
      `network-${index}`,
      { url: `https://network-${index}.example/page` },
      row.signal,
    );

    assert.match(resultText(result.content), row.expected, row.name);
    assert.equal(result.details?.ok, row.ok, row.name);
  }
});
