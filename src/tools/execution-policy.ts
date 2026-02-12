/**
 * Tool execution policy.
 *
 * Classifies tool calls as read-only vs workbook-mutating,
 * and whether a successful mutation should refresh workbook structure context.
 */

export type ToolExecutionMode = "read" | "mutate";
export type ToolContextImpact = "none" | "content" | "structure";

const ALWAYS_READ_TOOLS = new Set<string>([
  "get_workbook_overview",
  "read_range",
  "search_workbook",
  "trace_dependencies",
  // Instructions and conventions mutate local prompt/config state, not workbook cells/structure.
  "instructions",
  "conventions",
  // External bridge traffic does not mutate workbook state directly.
  "tmux",
  "python_run",
  "libreoffice_convert",
  "web_search",
  "mcp",
  // Workspace file operations do not mutate the workbook.
  "files",
  // Extension registry operations mutate local settings/runtime, not workbook content.
  "extensions_manager",
]);

const ALWAYS_MUTATE_TOOLS = new Set<string>([
  "write_cells",
  "fill_formula",
  "modify_structure",
  "format_cells",
  "conditional_format",
  // Bridge-assisted transform writes values back into the workbook.
  "python_transform_range",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function getActionParam(params: unknown): string | null {
  if (!isRecord(params)) return null;
  const action = params.action;
  return typeof action === "string" ? action : null;
}

function classifyViewSettings(params: unknown): ToolExecutionMode {
  const action = getActionParam(params);
  return action === "get" ? "read" : "mutate";
}

function classifyComments(params: unknown): ToolExecutionMode {
  const action = getActionParam(params);
  return action === "read" ? "read" : "mutate";
}

function classifyWorkbookHistory(params: unknown): ToolExecutionMode {
  const action = getActionParam(params);
  return action === "restore" ? "mutate" : "read";
}

function isViewSettingsStructureAction(params: unknown): boolean {
  const action = getActionParam(params);
  return action === "hide_sheet" || action === "show_sheet" || action === "very_hide_sheet";
}

/**
 * Return execution mode for a tool call.
 *
 * Unknown tools default to `mutate` as a safe fallback.
 */
export function getToolExecutionMode(toolName: string, params: unknown): ToolExecutionMode {
  if (ALWAYS_READ_TOOLS.has(toolName)) return "read";
  if (ALWAYS_MUTATE_TOOLS.has(toolName)) return "mutate";

  if (toolName === "view_settings") {
    return classifyViewSettings(params);
  }

  if (toolName === "comments") {
    return classifyComments(params);
  }

  if (toolName === "workbook_history") {
    return classifyWorkbookHistory(params);
  }

  return "mutate";
}

/**
 * Return context impact for a tool call.
 *
 * Lean default:
 * - only clearly structural mutations trigger workbook blueprint invalidation
 * - data/format/comment/view mutations are treated as content-only
 */
export function getToolContextImpact(toolName: string, params: unknown): ToolContextImpact {
  const mode = getToolExecutionMode(toolName, params);
  if (mode === "read") return "none";

  if (toolName === "modify_structure") {
    return "structure";
  }

  if (toolName === "view_settings" && isViewSettingsStructureAction(params)) {
    return "structure";
  }

  return "content";
}
