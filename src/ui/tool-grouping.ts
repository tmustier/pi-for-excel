/**
 * Tool card grouping rules. <assistant-message> renders consecutive finished
 * calls of the same tool as one collapsible group; this module names the
 * group ("3 reads"). Whether a group is open comes from transcript-state.ts.
 */

import { t } from "../language/index.js";

/** Tool name → locale key for the group header ("3 reads"). */
const TOOL_GROUP_LABEL_KEYS: Record<string, string> = {
  fill_formula: "tool-group.fill_formula",
  write_cells: "tool-group.write_cells",
  read_range: "tool-group.read_range",
  format_cells: "tool-group.format_cells",
  conditional_format: "tool-group.conditional_format",
  view_settings: "tool-group.view_settings",
  charts: "tool-group.charts",
  execute_office_js: "tool-group.script",
  execute_wps_js: "tool-group.script",
};

export function describeToolGroup(toolName: string, count: number): string {
  const key = TOOL_GROUP_LABEL_KEYS[toolName];
  return key ? t(key, { count }) : t("tool-group.other", { count, tool: toolName });
}
