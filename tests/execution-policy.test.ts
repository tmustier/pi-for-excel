import assert from "node:assert/strict";
import { test } from "node:test";

import { getToolContextImpact, getToolExecutionMode } from "../src/tools/execution-policy.ts";

void test("classifies read_range as read with no workbook-context impact", () => {
  assert.equal(getToolExecutionMode("read_range", {}), "read");
  assert.equal(getToolContextImpact("read_range", {}), "none");
});

void test("classifies modify_structure as structure-impact mutate", () => {
  assert.equal(getToolExecutionMode("modify_structure", { action: "add_sheet" }), "mutate");
  assert.equal(getToolContextImpact("modify_structure", { action: "add_sheet" }), "structure");
});

void test("classifies comments read vs mutate actions", () => {
  assert.equal(getToolExecutionMode("comments", { action: "read" }), "read");
  assert.equal(getToolContextImpact("comments", { action: "read" }), "none");

  assert.equal(getToolExecutionMode("comments", { action: "delete" }), "mutate");
  assert.equal(getToolContextImpact("comments", { action: "delete" }), "content");
});

void test("classifies view_settings actions by mode and context impact", () => {
  assert.equal(getToolExecutionMode("view_settings", { action: "get" }), "read");
  assert.equal(getToolContextImpact("view_settings", { action: "get" }), "none");

  assert.equal(getToolExecutionMode("view_settings", { action: "activate" }), "mutate");
  assert.equal(getToolContextImpact("view_settings", { action: "activate" }), "content");

  assert.equal(getToolExecutionMode("view_settings", { action: "hide_sheet" }), "mutate");
  assert.equal(getToolContextImpact("view_settings", { action: "hide_sheet" }), "structure");

  assert.equal(getToolExecutionMode("view_settings", { action: "set_standard_width" }), "mutate");
  assert.equal(getToolContextImpact("view_settings", { action: "set_standard_width" }), "content");
});

void test("classifies instructions as non-workbook read traffic", () => {
  assert.equal(getToolExecutionMode("instructions", { action: "append", level: "user" }), "read");
  assert.equal(getToolContextImpact("instructions", { action: "append", level: "user" }), "none");
});

void test("classifies workbook_history list/delete as read and restore as mutate", () => {
  assert.equal(getToolExecutionMode("workbook_history", { action: "list" }), "read");
  assert.equal(getToolContextImpact("workbook_history", { action: "list" }), "none");

  assert.equal(getToolExecutionMode("workbook_history", { action: "delete", snapshot_id: "abc" }), "read");
  assert.equal(getToolContextImpact("workbook_history", { action: "delete", snapshot_id: "abc" }), "none");

  assert.equal(getToolExecutionMode("workbook_history", { action: "restore", snapshot_id: "abc" }), "mutate");
  assert.equal(getToolContextImpact("workbook_history", { action: "restore", snapshot_id: "abc" }), "content");
});

void test("classifies bridge and external tools as read-only non-workbook traffic", () => {
  assert.equal(getToolExecutionMode("tmux", { action: "list_sessions" }), "read");
  assert.equal(getToolContextImpact("tmux", { action: "list_sessions" }), "none");

  assert.equal(getToolExecutionMode("python_run", { code: "print(1)" }), "read");
  assert.equal(getToolContextImpact("python_run", { code: "print(1)" }), "none");

  assert.equal(
    getToolExecutionMode("libreoffice_convert", { input_path: "/tmp/a.xlsx", target_format: "csv" }),
    "read",
  );
  assert.equal(
    getToolContextImpact("libreoffice_convert", { input_path: "/tmp/a.xlsx", target_format: "csv" }),
    "none",
  );

  assert.equal(getToolExecutionMode("web_search", { query: "latest CPI" }), "read");
  assert.equal(getToolContextImpact("web_search", { query: "latest CPI" }), "none");

  assert.equal(getToolExecutionMode("mcp", { server: "local" }), "read");
  assert.equal(getToolContextImpact("mcp", { server: "local" }), "none");

  assert.equal(getToolExecutionMode("files", { action: "list" }), "read");
  assert.equal(getToolContextImpact("files", { action: "list" }), "none");
});

void test("classifies python_transform_range as workbook content mutation", () => {
  assert.equal(
    getToolExecutionMode("python_transform_range", {
      range: "Sheet1!A1:B10",
      code: "result = input_data['values']",
    }),
    "mutate",
  );

  assert.equal(
    getToolContextImpact("python_transform_range", {
      range: "Sheet1!A1:B10",
      code: "result = input_data['values']",
    }),
    "content",
  );
});

void test("unknown tools default to mutate with content impact", () => {
  assert.equal(getToolExecutionMode("extension_tool", { any: true }), "mutate");
  assert.equal(getToolContextImpact("extension_tool", { any: true }), "content");
});
