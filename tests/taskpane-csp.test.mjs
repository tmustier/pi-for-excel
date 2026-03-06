import assert from "node:assert/strict";
import test from "node:test";

import {
  readTaskpaneConnectSrcTokens,
  readTaskpaneScriptSrcTokens,
} from "./helpers/taskpane-csp.mjs";

test("taskpane CSP allows Pyodide CDN host in script-src and connect-src", async () => {
  const connectTokens = await readTaskpaneConnectSrcTokens();
  const scriptTokens = await readTaskpaneScriptSrcTokens();

  assert.ok(connectTokens.has("https://cdn.jsdelivr.net"), "Missing jsDelivr in CSP connect-src");
  assert.ok(scriptTokens.has("https://cdn.jsdelivr.net"), "Missing jsDelivr in CSP script-src");
});

test("taskpane CSP allows blob module imports in script-src", async () => {
  const scriptTokens = await readTaskpaneScriptSrcTokens();

  assert.ok(scriptTokens.has("blob:"), "Missing blob: in CSP script-src");
});
