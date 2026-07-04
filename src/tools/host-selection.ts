/** Host-specific core tool implementation selection. */

import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { TSchema } from "typebox";

import type { SpreadsheetHostKind } from "../host/index.js";
import type { CoreToolName } from "./names.js";
import { createUnsupportedHostTool } from "./unsupported-host-tool.js";

export type AnyHostSelectableTool = AgentTool<TSchema, unknown>;

const WPS_SUPPORTED_LOCAL_CORE_TOOL_NAMES = new Set<CoreToolName>([
  "instructions",
  "conventions",
  "skills",
]);

export function isCoreToolUnsupportedOnWps(name: CoreToolName): boolean {
  return !WPS_SUPPORTED_LOCAL_CORE_TOOL_NAMES.has(name);
}

export function selectCoreToolForHost(
  name: CoreToolName,
  tool: AnyHostSelectableTool,
  hostKind: SpreadsheetHostKind,
): AnyHostSelectableTool {
  if (hostKind === "wps" && isCoreToolUnsupportedOnWps(name)) {
    return createUnsupportedHostTool(tool, hostKind);
  }

  return tool;
}
