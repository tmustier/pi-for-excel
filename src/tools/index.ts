/**
 * Tool registry — creates all Excel tools for the agent.
 */

import type { AgentTool } from "@mariozechner/pi-agent-core";
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
import type { ChangeTracker } from "../context/change-tracker.js";

/** Create all Excel tools */
export function createAllTools(opts?: { changeTracker?: ChangeTracker }): AgentTool<any>[] {
  const tools: AgentTool<any>[] = [
    createGetWorkbookOverviewTool(),
    createReadRangeTool(),
    createReadSelectionTool(),
    createWriteCellsTool(),
    createFillFormulaTool(),
    createSearchWorkbookTool(),
    createModifyStructureTool(),
    createFormatCellsTool(),
    createConditionalFormatTool(),
    createTraceDependenciesTool(),
  ];

  if (opts?.changeTracker) {
    tools.push(createGetRecentChangesTool(opts.changeTracker));
  }

  return tools;
}
