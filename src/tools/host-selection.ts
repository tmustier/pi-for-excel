import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { TSchema } from "typebox";

import type { SpreadsheetHostKind } from "../host/index.js";
import { CORE_TOOL_NAMES, type CoreToolName } from "./names.js";
import { createUnsupportedHostTool } from "./unsupported-host-tool.js";
import {
  executeWpsGetWorkbookOverview,
  executeWpsReadRange,
  executeWpsWriteCells,
} from "./wps/workbook-tools.js";

export type AnyHostSelectableTool = AgentTool<TSchema, unknown>;

const WPS_SUPPORTED_LOCAL_CORE_TOOL_NAMES = new Set<CoreToolName>([
  "instructions",
  "conventions",
  "skills",
]);

type HostExecuteOverride = AnyHostSelectableTool["execute"];

export const WPS_CORE_TOOL_EXECUTE_OVERRIDES: Partial<Record<CoreToolName, HostExecuteOverride>> = {
  get_workbook_overview: executeWpsGetWorkbookOverview,
  read_range: executeWpsReadRange,
  write_cells: executeWpsWriteCells,
};

export function isCoreToolUnsupportedOnWps(name: CoreToolName): boolean {
  return !WPS_SUPPORTED_LOCAL_CORE_TOOL_NAMES.has(name) && !WPS_CORE_TOOL_EXECUTE_OVERRIDES[name];
}

export function selectCoreToolForHost(
  name: CoreToolName,
  tool: AnyHostSelectableTool,
  hostKind: SpreadsheetHostKind,
): AnyHostSelectableTool {
  if (hostKind === "wps") {
    const executeOverride = WPS_CORE_TOOL_EXECUTE_OVERRIDES[name];
    if (executeOverride) {
      return { ...tool, execute: executeOverride };
    }

    if (isCoreToolUnsupportedOnWps(name)) {
      return createUnsupportedHostTool(tool, hostKind);
    }
  }

  return tool;
}

export function composeCoreToolsForHost(
  createTool: (name: CoreToolName) => AnyHostSelectableTool,
  hostKind: SpreadsheetHostKind,
): AnyHostSelectableTool[] {
  return CORE_TOOL_NAMES.map((name) => selectCoreToolForHost(name, createTool(name), hostKind));
}

export function selectOfficeCoupledToolForHost(
  tool: AnyHostSelectableTool,
  hostKind: SpreadsheetHostKind,
): AnyHostSelectableTool {
  if (hostKind === "wps") {
    return createUnsupportedHostTool(tool, hostKind);
  }

  return tool;
}
