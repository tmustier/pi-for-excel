import assert from "node:assert/strict";
import { test } from "node:test";

import {
  getStatusContextHealth,
  parseStatusContextWarningSeverity,
} from "../src/taskpane/status-context.ts";

void test("returns no warning at or below 40% usage", () => {
  const health = getStatusContextHealth(40);

  assert.equal(health.colorClass, "");
  assert.equal(health.warning, null);
});

void test("returns yellow warning above 40% usage", () => {
  const health = getStatusContextHealth(41);

  assert.equal(health.colorClass, "pi-status-ctx--yellow");
  assert.equal(health.warning?.severity, "yellow");
  assert.equal(health.warning?.text, "Context 41% full.");
  assert.equal(
    health.warning?.actionText,
    "Consider using /compact to free space or /new to start fresh.",
  );
});

void test("returns red warning above 60% usage", () => {
  const health = getStatusContextHealth(61);

  assert.equal(health.colorClass, "pi-status-ctx--red");
  assert.equal(health.warning?.severity, "red");
  assert.equal(health.warning?.text, "Context 61% full — responses may become less reliable.");
  assert.equal(
    health.warning?.actionText,
    "Use /compact to free space or /new to start fresh.",
  );
});

void test("returns full-context warning above 100% usage", () => {
  const health = getStatusContextHealth(101);

  assert.equal(health.colorClass, "pi-status-ctx--red");
  assert.equal(health.warning?.severity, "red");
  assert.equal(health.warning?.text, "Context is full — the next message will fail.");
});

void test("normalizes warning severity from context attributes", () => {
  assert.equal(parseStatusContextWarningSeverity("red"), "red");
  assert.equal(parseStatusContextWarningSeverity("yellow"), "yellow");
  assert.equal(parseStatusContextWarningSeverity("unknown"), "yellow");
  assert.equal(parseStatusContextWarningSeverity(null), "yellow");
});
