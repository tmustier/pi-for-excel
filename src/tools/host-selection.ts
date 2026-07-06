/** Host-specific core tool implementation selection. */

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

export type AnyHostSelectableTool = AgentTool<TSchema, DynamicValue>;

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

/**
 * Compose the core tool list for a host: `CORE_TOOL_NAMES` drives ordering,
 * `selectCoreToolForHost` swaps in fail-fast handlers where a host has no
 * implementation. Kept free of tool-implementation imports so the composition
 * behavior is directly unit-testable.
 */
export function composeCoreToolsForHost(
  createTool: (name: CoreToolName) => AnyHostSelectableTool,
  hostKind: SpreadsheetHostKind,
): AnyHostSelectableTool[] {
  return CORE_TOOL_NAMES.map((name) => selectCoreToolForHost(name, createTool(name), hostKind));
}

/**
 * Non-core tools that drive Office.js/Excel directly (e.g. `execute_office_js`,
 * `python_transform_range`) must also fail fast on hosts without an Office.js
 * runtime instead of reaching Office/Excel helper paths.
 */
export function selectOfficeCoupledToolForHost(
  tool: AnyHostSelectableTool,
  hostKind: SpreadsheetHostKind,
): AnyHostSelectableTool {
  if (hostKind === "wps") {
    return createUnsupportedHostTool(tool, hostKind);
  }

  return tool;
}
