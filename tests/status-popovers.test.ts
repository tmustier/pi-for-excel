import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

import { resolveStatusPopoverLayout } from "../src/taskpane/status-popovers.ts";

void test("thinking popover stays inside a 290x400 short taskpane", () => {
  const layout = resolveStatusPopoverLayout({
    anchor: { top: 350, bottom: 380, right: 280 },
    viewportWidth: 290,
    viewportHeight: 400,
    popoverWidth: 274,
    popoverHeight: 488,
  });

  assert.deepEqual(layout, {
    left: 8,
    top: 8,
    maxHeight: 384,
  });
  assert.ok(layout.top + layout.maxHeight <= 400 - 8);
});

void test("thinking popover remains above its anchor when it fits", () => {
  const layout = resolveStatusPopoverLayout({
    anchor: { top: 800, bottom: 830, right: 780 },
    viewportWidth: 800,
    viewportHeight: 900,
    popoverWidth: 290,
    popoverHeight: 300,
  });

  assert.deepEqual(layout, {
    left: 490,
    top: 492,
    maxHeight: 884,
  });
});

void test("status popover CSS constrains the shell and scrolls its choices", () => {
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
