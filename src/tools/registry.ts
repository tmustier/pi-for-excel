/**
 * Capability registry (core)
 *
 * Canonical source of truth for built-in Excel tool names + construction.
 *
 * Note: extensions will later register additional tools at runtime, but this
 * module only covers the built-in (core) tools.
 */

import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { TSchema } from "typebox";

import type { SpreadsheetHostKind } from "../host/index.js";
import { createGetWorkbookOverviewTool } from "./get-workbook-overview.js";
import { createReadRangeTool } from "./read-range.js";
import { createWriteCellsTool } from "./write-cells.js";
import { createFillFormulaTool } from "./fill-formula.js";
import { createSearchWorkbookTool } from "./search-workbook.js";
import { createModifyStructureTool } from "./modify-structure.js";
import { createFormatCellsTool } from "./format-cells.js";
import { createConditionalFormatTool } from "./conditional-format.js";
import { createChartsTool } from "./charts.js";
import { createTraceDependenciesTool } from "./trace-dependencies.js";
import { createExplainFormulaTool } from "./explain-formula.js";
import { createViewSettingsTool } from "./view-settings.js";
import { createCommentsTool } from "./comments.js";
import { createInstructionsTool } from "./instructions.js";
import { createConventionsTool } from "./conventions.js";
import { createWorkbookHistoryTool } from "./workbook-history.js";
import {
  createSkillsTool,
  type SkillsToolDependencies,
} from "./skills.js";
import { composeCoreToolsForHost } from "./host-selection.js";
import type { CoreToolName } from "./names.js";

export { CORE_TOOL_NAMES } from "./names.js";
export type { CoreToolName } from "./names.js";

// We intentionally erase per-tool parameter typing at the list boundary.
// Each tool still validates its own schema at runtime.
export type AnyCoreTool = AgentTool<TSchema, DynamicValue>;

export interface CreateCoreToolsOptions {
  hostKind?: SpreadsheetHostKind;
  skills?: SkillsToolDependencies;
}

type CoreToolFactory = (options: CreateCoreToolsOptions) => AnyCoreTool;

const CORE_TOOL_FACTORIES = {
  get_workbook_overview: () => createGetWorkbookOverviewTool(),
  read_range: () => createReadRangeTool(),
  write_cells: () => createWriteCellsTool(),
  fill_formula: () => createFillFormulaTool(),
  search_workbook: () => createSearchWorkbookTool(),
  modify_structure: () => createModifyStructureTool(),
  format_cells: () => createFormatCellsTool(),
  conditional_format: () => createConditionalFormatTool(),
  charts: () => createChartsTool(),
  trace_dependencies: () => createTraceDependenciesTool(),
  explain_formula: () => createExplainFormulaTool(),
  view_settings: () => createViewSettingsTool(),
  comments: () => createCommentsTool(),
  instructions: () => createInstructionsTool(),
  conventions: () => createConventionsTool(),
  workbook_history: () => createWorkbookHistoryTool(),
  skills: (options) => createSkillsTool(options.skills),
} satisfies Record<CoreToolName, CoreToolFactory>;

export { isCoreToolUnsupportedOnWps } from "./host-selection.js";

/** Create all core (built-in) tools for the agent. */
export function createCoreTools(options: CreateCoreToolsOptions = {}): AnyCoreTool[] {
  const hostKind = options.hostKind ?? "office";

  return composeCoreToolsForHost((name) => CORE_TOOL_FACTORIES[name](options), hostKind);
}
