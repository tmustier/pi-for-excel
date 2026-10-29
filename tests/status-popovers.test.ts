import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

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
