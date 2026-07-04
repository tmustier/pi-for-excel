/** Canonical list of built-in core tool names. */
export const CORE_TOOL_NAMES = [
  "get_workbook_overview",
  "read_range",
  "write_cells",
  "fill_formula",
  "search_workbook",
  "modify_structure",
  "format_cells",
  "conditional_format",
  "charts",
  "trace_dependencies",
  "explain_formula",
  "view_settings",
  "comments",
  "instructions",
  "conventions",
  "workbook_history",
  "skills",
] as const;

export type CoreToolName = (typeof CORE_TOOL_NAMES)[number];
