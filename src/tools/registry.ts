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

import { createGetWorkbookOverviewTool } from "./get-workbook-overview.js";
import { createReadRangeTool } from "./read-range.js";
import { createWriteCellsTool } from "./write-cells.js";
import { createFillFormulaTool } from "./fill-formula.js";
import { createSearchWorkbookTool } from "./search-workbook.js";
import { createModifyStructureTool } from "./modify-structure.js";
import { createFormatCellsTool } from "./format-cells.js";
import { createConditionalFormatTool } from "./conditional-format.js";
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

export { CORE_TOOL_NAMES } from "./names.js";
export type { CoreToolName } from "./names.js";

// We intentionally erase per-tool parameter typing at the list boundary.
// Each tool still validates its own schema at runtime.
export type AnyCoreTool = AgentTool<TSchema, unknown>;

export interface CreateCoreToolsOptions {
  skills?: SkillsToolDependencies;
}

/** Create all core (built-in) Excel tools for the agent. */
export function createCoreTools(options: CreateCoreToolsOptions = {}): AnyCoreTool[] {
  return [
    createGetWorkbookOverviewTool(),
    createReadRangeTool(),
    createWriteCellsTool(),
    createFillFormulaTool(),
    createSearchWorkbookTool(),
    createModifyStructureTool(),
    createFormatCellsTool(),
    createConditionalFormatTool(),
    createTraceDependenciesTool(),
    createExplainFormulaTool(),
    createViewSettingsTool(),
    createCommentsTool(),
    createInstructionsTool(),
    createConventionsTool(),
    createWorkbookHistoryTool(),
    createSkillsTool(options.skills),
  ] as unknown as AnyCoreTool[];
}
