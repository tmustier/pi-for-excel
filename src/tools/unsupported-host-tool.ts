/** Fail-fast wrappers for tools whose implementation is absent on a host. */

import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { TSchema } from "typebox";

import type { SpreadsheetHostKind } from "../host/index.js";
import { WPS_UNSUPPORTED_PHASE_1_MESSAGE } from "../host/index.js";

function unsupportedMessage(hostKind: SpreadsheetHostKind, toolName: string): string {
  if (hostKind === "wps") {
    return `${toolName} is not yet supported on WPS Spreadsheets. ${WPS_UNSUPPORTED_PHASE_1_MESSAGE}`;
  }

  return `${toolName} is not supported on host '${hostKind}'.`;
}

export function createUnsupportedHostTool<TParameters extends TSchema, TDetails>(
  tool: AgentTool<TParameters, TDetails>,
  hostKind: SpreadsheetHostKind,
): AgentTool<TParameters, TDetails> {
  return {
    ...tool,
    execute: () => {
      throw new Error(unsupportedMessage(hostKind, tool.name));
    },
  };
}
