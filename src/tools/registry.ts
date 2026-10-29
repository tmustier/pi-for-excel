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
import type { RecoveryAttribution } from "./recovery-metadata.js";
import type { GuidanceChangeDependencies } from "./guidance-change.js";

export { CORE_TOOL_NAMES } from "./names.js";
export type { CoreToolName } from "./names.js";

export type AnyCoreTool = AgentTool<TSchema, unknown>;

export interface CreateCoreToolsOptions {
  hostKind?: SpreadsheetHostKind;
  skills?: SkillsToolDependencies;
  recovery?: RecoveryAttribution;
  guidance?: GuidanceChangeDependencies;
}

type CoreToolFactory = (options: CreateCoreToolsOptions) => AnyCoreTool;

const CORE_TOOL_FACTORIES = {
  get_workbook_overview: () => createGetWorkbookOverviewTool(),
  read_range: () => createReadRangeTool(),
  write_cells: (options) => createWriteCellsTool({}, options.recovery),
  fill_formula: (options) => createFillFormulaTool(options.recovery),
  search_workbook: () => createSearchWorkbookTool(),
  modify_structure: (options) => createModifyStructureTool(options.recovery),
  format_cells: (options) => createFormatCellsTool(options.recovery),
  conditional_format: (options) => createConditionalFormatTool(options.recovery),
  charts: (options) => createChartsTool({}, options.recovery),
  trace_dependencies: () => createTraceDependenciesTool(),
  explain_formula: () => createExplainFormulaTool(),
  view_settings: () => createViewSettingsTool(),
  comments: (options) => createCommentsTool({}, options.recovery),
  instructions: (options) => createInstructionsTool(options.guidance),
  conventions: (options) => createConventionsTool(options.guidance),
  workbook_history: (options) => createWorkbookHistoryTool({}, options.recovery),
  skills: (options) => createSkillsTool({ ...options.skills, ...options.guidance }),
} satisfies Record<CoreToolName, CoreToolFactory>;

export { isCoreToolUnsupportedOnWps } from "./host-selection.js";

export function createCoreTools(options: CreateCoreToolsOptions = {}): AnyCoreTool[] {
  const hostKind = options.hostKind ?? "office";

  return composeCoreToolsForHost((name) => CORE_TOOL_FACTORIES[name](options), hostKind);
}
