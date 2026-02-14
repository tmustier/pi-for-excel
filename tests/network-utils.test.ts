import assert from "node:assert/strict";
import { test } from "node:test";

import { getHttpErrorReason, runWithTimeoutAbort } from "../src/utils/network.ts";

void test("getHttpErrorReason prefers non-empty response body", () => {
  assert.equal(getHttpErrorReason(502, " upstream failed "), "upstream failed");
  assert.equal(getHttpErrorReason(404, "   \n  "), "HTTP 404");
});

void test("runWithTimeoutAbort returns run result", async () => {
  const value = await runWithTimeoutAbort({
    signal: undefined,
    timeoutMs: 50,
    timeoutErrorMessage: "timed out",
    run: () => Promise.resolve("ok"),
  });

  assert.equal(value, "ok");
});

void test("runWithTimeoutAbort throws timeout error", async () => {
  await assert.rejects(
    runWithTimeoutAbort({
      signal: undefined,
      timeoutMs: 5,
      timeoutErrorMessage: "request timed out",
      run: async (_signal) => {
        return new Promise<string>(() => {
          // Never resolves; timeout drives completion.
        });
      },
    }),
    /request timed out/,
  );
});

void test("runWithTimeoutAbort preserves caller abort semantics", async () => {
  const callerController = new AbortController();

  const pending = runWithTimeoutAbort({
    signal: callerController.signal,
    timeoutMs: 200,
    timeoutErrorMessage: "request timed out",
    run: async (requestSignal) => {
      return new Promise<string>((_resolve, reject) => {
        requestSignal.addEventListener("abort", () => {
          reject(new DOMException("aborted", "AbortError"));
        }, { once: true });
      });
    },
  });

  callerController.abort();

  await assert.rejects(pending, /^Error: Aborted$/);
});
