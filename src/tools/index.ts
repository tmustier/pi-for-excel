/**
 * Tool registry — creates all Excel tools for the agent.
 */

import type { AgentTool } from "@mariozechner/pi-agent-core";
import type { TSchema } from "@sinclair/typebox";
import { createGetWorkbookOverviewTool } from "./get-workbook-overview.js";
import { createReadRangeTool } from "./read-range.js";
import { createWriteCellsTool } from "./write-cells.js";
import { createSearchWorkbookTool } from "./search-workbook.js";
import { createModifyStructureTool } from "./modify-structure.js";
import { createFormatCellsTool } from "./format-cells.js";
import { createTraceDependenciesTool } from "./trace-dependencies.js";
import { createConditionalFormatTool } from "./conditional-format.js";
import { createFillFormulaTool } from "./fill-formula.js";
import { createReadSelectionTool } from "./read-selection.js";
import { createGetRecentChangesTool } from "./get-recent-changes.js";
import { createGetRangeAsCsvTool } from "./get-range-as-csv.js";
import { createGetAllObjectsTool } from "./get-all-objects.js";
import type { ChangeTracker } from "../context/change-tracker.js";

type AnyTool = AgentTool<TSchema, unknown>;

/** Create all Excel tools */
export function createAllTools(opts?: { changeTracker?: ChangeTracker }): AnyTool[] {
  const tools = [
    createGetWorkbookOverviewTool(),
    createReadRangeTool(),
    createReadSelectionTool(),
    createGetRangeAsCsvTool(),
    createGetAllObjectsTool(),
    createWriteCellsTool(),
    createFillFormulaTool(),
    createSearchWorkbookTool(),
    createModifyStructureTool(),
    createFormatCellsTool(),
    createConditionalFormatTool(),
    createTraceDependenciesTool(),
  ] as unknown as AnyTool[];

  if (opts?.changeTracker) {
    tools.push(createGetRecentChangesTool(opts.changeTracker) as unknown as AnyTool);
  }

  return tools;
}
