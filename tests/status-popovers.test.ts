import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

// Static policy: the browser tests cover computed placement; these declarations encode
// the separate CSS ownership contract for which element clips and which one scrolls.
void test("static contract: status popover shell constrains height and delegates scrolling to its list", () => {
  const css = readFileSync("src/ui/theme/components/status-bar.css", "utf8");

  assert.match(
    css,
    /\.pi-status-popover\s*\{[^}]*max-height:\s*calc\(100vh - 16px\);[^}]*overflow:\s*hidden;/su,
  );
  assert.match(
    css,
    /\.pi-status-popover__list,[^}]*min-height:\s*0;[^}]*overflow-y:\s*auto;/su,
  );
});
