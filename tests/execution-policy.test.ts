import assert from "node:assert/strict";
import { test } from "node:test";

import { getToolApprovalRequirement, getToolContextImpact, getToolExecutionMode } from "../src/tools/execution-policy.ts";

void test("execution policy classifies every tool and action", () => {
  const rows: Array<{
    name: string;
    params: unknown;
    mode: "read" | "mutate";
    impact: "none" | "content" | "structure";
  }> = [
    { name: "read_range", params: {}, mode: "read", impact: "none" },
    { name: "trace_dependencies", params: { cell: "Sheet1!D10" }, mode: "read", impact: "none" },
    { name: "trace_dependencies", params: { cell: "Sheet1!D10", mode: "dependents" }, mode: "read", impact: "none" },
    { name: "explain_formula", params: { cell: "Sheet1!D10" }, mode: "read", impact: "none" },
    { name: "modify_structure", params: { action: "add_sheet" }, mode: "mutate", impact: "structure" },
    { name: "comments", params: { action: "read" }, mode: "read", impact: "none" },
    { name: "comments", params: { action: "delete" }, mode: "mutate", impact: "content" },
    { name: "view_settings", params: { action: "get" }, mode: "read", impact: "none" },
    { name: "view_settings", params: { action: "activate" }, mode: "mutate", impact: "content" },
    { name: "view_settings", params: { action: "hide_sheet" }, mode: "mutate", impact: "structure" },
    { name: "view_settings", params: { action: "set_standard_width" }, mode: "mutate", impact: "content" },
    { name: "instructions", params: { action: "append", level: "user" }, mode: "read", impact: "none" },
    { name: "workbook_history", params: { action: "list" }, mode: "read", impact: "none" },
    { name: "workbook_history", params: { action: "delete", snapshot_id: "abc" }, mode: "read", impact: "none" },
    { name: "workbook_history", params: { action: "restore", snapshot_id: "abc" }, mode: "mutate", impact: "content" },
    { name: "tmux", params: { action: "list_sessions" }, mode: "read", impact: "none" },
    { name: "python_run", params: { code: "print(1)" }, mode: "read", impact: "none" },
    { name: "libreoffice_convert", params: { input_path: "/tmp/a.xlsx", target_format: "csv" }, mode: "read", impact: "none" },
    { name: "web_search", params: { query: "latest CPI" }, mode: "read", impact: "none" },
    { name: "fetch_page", params: { url: "https://example.com" }, mode: "read", impact: "none" },
    { name: "mcp", params: { server: "local" }, mode: "read", impact: "none" },
    { name: "files", params: { action: "list" }, mode: "read", impact: "none" },
    { name: "extensions_manager", params: { action: "list" }, mode: "read", impact: "none" },
    // Skills read/write local skill files, never the workbook.
    { name: "skills", params: { action: "list" }, mode: "read", impact: "none" },
    { name: "skills", params: { action: "install" }, mode: "read", impact: "none" },
    { name: "python_transform_range", params: { range: "Sheet1!A1:B10", code: "result = input_data['values']" }, mode: "mutate", impact: "content" },
    { name: "execute_office_js", params: { explanation: "Update workbook settings", code: "return { ok: true };" }, mode: "mutate", impact: "structure" },
    { name: "skills", params: { action: "read", name: "x" }, mode: "read", impact: "none" },
    { name: "extension_tool", params: { custom: true }, mode: "mutate", impact: "content" },
  ];

  for (const row of rows) {
    assert.equal(getToolExecutionMode(row.name, row.params), row.mode, `${row.name} mode`);
    assert.equal(getToolContextImpact(row.name, row.params), row.impact, `${row.name} impact`);
  }
});

void test("approval requirement: Confirm mode asks for workbook changes; deleting backups always asks", () => {
  const rows: Array<{ name: string; params: unknown; requirement: "none" | "confirm_mode" | "always" }> = [
    { name: "read_range", params: {}, requirement: "none" },
    { name: "skills", params: { action: "list" }, requirement: "none" },
    { name: "write_cells", params: { start_cell: "A1" }, requirement: "confirm_mode" },
    { name: "workbook_history", params: { action: "list" }, requirement: "none" },
    { name: "workbook_history", params: { action: "restore", snapshot_id: "abc" }, requirement: "confirm_mode" },
    { name: "workbook_history", params: { action: "delete", snapshot_id: "abc" }, requirement: "always" },
    { name: "workbook_history", params: { action: "clear" }, requirement: "always" },
    // Its own gate asks, so the coordinator must not ask a second time.
    { name: "execute_office_js", params: { code: "" }, requirement: "none" },
    { name: "extension_tool", params: {}, requirement: "confirm_mode" },
  ];

  for (const row of rows) {
    assert.equal(getToolApprovalRequirement(row.name, row.params), row.requirement, `${row.name} ${JSON.stringify(row.params)}`);
  }
});
