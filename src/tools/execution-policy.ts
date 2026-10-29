export type ToolExecutionMode = "read" | "mutate";
export type ToolContextImpact = "none" | "content" | "structure";

const ALWAYS_READ_TOOLS = new Set<string>([
  "get_workbook_overview",
  "read_range",
  "search_workbook",
  "trace_dependencies",
  "explain_formula",
  "instructions",
  "conventions",
  "skills",
  "tmux",
  "python_run",
  "libreoffice_convert",
  "web_search",
  "fetch_page",
  "mcp",
  "files",
  "extensions_manager",
]);

const ALWAYS_MUTATE_TOOLS = new Set<string>([
  "write_cells",
  "fill_formula",
  "modify_structure",
  "format_cells",
  "conditional_format",
  "python_transform_range",
  "execute_office_js",
  "execute_wps_js",
]);

function isToolsExecutionPolicyPayloadShape(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function getActionParam(params: unknown): string | null {
  if (!isToolsExecutionPolicyPayloadShape(params)) return null;
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

function classifyCharts(params: unknown): ToolExecutionMode {
  const action = getActionParam(params);
  return action === "list" || action === "get_image" ? "read" : "mutate";
}

function classifyWorkbookHistory(params: unknown): ToolExecutionMode {
  const action = getActionParam(params);
  return action === "restore" ? "mutate" : "read";
}

function isViewSettingsStructureAction(params: unknown): boolean {
  const action = getActionParam(params);
  return action === "hide_sheet" || action === "show_sheet" || action === "very_hide_sheet";
}

function isChartsStructureAction(params: unknown): boolean {
  const action = getActionParam(params);
  return action === "create" || action === "delete";
}

export function getToolExecutionMode(toolName: string, params: unknown): ToolExecutionMode {
  if (ALWAYS_READ_TOOLS.has(toolName)) return "read";
  if (ALWAYS_MUTATE_TOOLS.has(toolName)) return "mutate";

  if (toolName === "view_settings") {
    return classifyViewSettings(params);
  }

  if (toolName === "comments") {
    return classifyComments(params);
  }

  if (toolName === "charts") {
    return classifyCharts(params);
  }

  if (toolName === "workbook_history") {
    return classifyWorkbookHistory(params);
  }

  return "mutate";
}

export type ToolApprovalRequirement = "none" | "confirm_mode" | "always" | "irreversible";

const TOOLS_WITH_OWN_APPROVAL_GATE = new Set<string>([
  "execute_office_js",
  "execute_wps_js",
]);

export function getToolApprovalRequirement(toolName: string, params: unknown): ToolApprovalRequirement {
  if (toolName === "charts" && getActionParam(params) === "delete") return "irreversible";
  if (toolName === "workbook_history") {
    const action = getActionParam(params);
    if (action === "delete" || action === "clear") return "always";
  }

  if (TOOLS_WITH_OWN_APPROVAL_GATE.has(toolName)) return "none";

  return getToolExecutionMode(toolName, params) === "mutate" ? "confirm_mode" : "none";
}

export function getToolContextImpact(toolName: string, params: unknown): ToolContextImpact {
  const mode = getToolExecutionMode(toolName, params);
  if (mode === "read") return "none";

  if (toolName === "modify_structure") {
    return "structure";
  }

  if (toolName === "execute_office_js" || toolName === "execute_wps_js") {
    return "structure";
  }

  if (toolName === "view_settings" && isViewSettingsStructureAction(params)) {
    return "structure";
  }

  if (toolName === "charts" && isChartsStructureAction(params)) {
    return "structure";
  }

  return "content";
}
