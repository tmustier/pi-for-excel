/**
 * Capability registry (core)
 *
 * Canonical source of truth for built-in Excel tool names + construction.
 *
 * Note: extensions will later register additional tools at runtime, but this
 * module only covers the built-in (core) tools.
 */

import type { AgentTool } from "@mariozechner/pi-agent-core";
import type { TSchema } from "@sinclair/typebox";

import { createGetWorkbookOverviewTool } from "./get-workbook-overview.js";
import { createReadRangeTool } from "./read-range.js";
import { createWriteCellsTool } from "./write-cells.js";
import { createFillFormulaTool } from "./fill-formula.js";
import { createSearchWorkbookTool } from "./search-workbook.js";
import { createModifyStructureTool } from "./modify-structure.js";
import { createFormatCellsTool } from "./format-cells.js";
import { createConditionalFormatTool } from "./conditional-format.js";
import { createTraceDependenciesTool } from "./trace-dependencies.js";
import { createViewSettingsTool } from "./view-settings.js";
import { createCommentsTool } from "./comments.js";
import { createInstructionsTool } from "./instructions.js";
import { createConventionsTool } from "./conventions.js";
import { createWorkbookHistoryTool } from "./workbook-history.js";

/** Canonical list of core tool names (single source of truth). */
export const CORE_TOOL_NAMES = [
  "get_workbook_overview",
  "read_range",
  "write_cells",
  "fill_formula",
  "search_workbook",
  "modify_structure",
  "format_cells",
  "conditional_format",
  "trace_dependencies",
  "view_settings",
  "comments",
  "instructions",
  "conventions",
  "workbook_history",
] as const;

export type CoreToolName = (typeof CORE_TOOL_NAMES)[number];

// We intentionally erase per-tool parameter typing at the list boundary.
// Each tool still validates its own schema at runtime.
export type AnyCoreTool = AgentTool<TSchema, unknown>;

/** Create all core (built-in) Excel tools for the agent. */
export function createCoreTools(): AnyCoreTool[] {
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
    createViewSettingsTool(),
    createCommentsTool(),
    createInstructionsTool(),
    createConventionsTool(),
    createWorkbookHistoryTool(),
  ] as unknown as AnyCoreTool[];
}
