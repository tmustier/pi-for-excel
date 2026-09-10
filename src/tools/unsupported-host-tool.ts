/** Fail-fast wrappers for tools whose implementation is absent on a host. */

import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { TSchema } from "typebox";

import type { SpreadsheetHostKind } from "../host/index.js";
import { WPS_UNSUPPORTED_PHASE_1_MESSAGE } from "../host/index.js";

const UNSUPPORTED_HOST_TOOL_MARKER = Symbol("pi.unsupportedHostTool");

interface UnsupportedHostToolMarker {
  readonly [UNSUPPORTED_HOST_TOOL_MARKER]: true;
}

function unsupportedMessage(hostKind: SpreadsheetHostKind, toolName: string): string {
  if (hostKind === "wps") {
    return `${toolName} is not yet supported on WPS Spreadsheets. ${WPS_UNSUPPORTED_PHASE_1_MESSAGE}`;
  }

  return `${toolName} is not supported on host '${hostKind}'.`;
}

/**
 * Typed error thrown when a tool has no implementation on the current host.
 *
 * Callers/tests/UI can distinguish "unsupported on this host" from an
 * implementation failure via `instanceof` or the stable `code` field.
 */
export class UnsupportedHostToolError extends Error {
  readonly code = "unsupported_host_tool";
  readonly hostKind: SpreadsheetHostKind;
  readonly toolName: string;

  constructor(hostKind: SpreadsheetHostKind, toolName: string) {
    super(unsupportedMessage(hostKind, toolName));
    this.name = "UnsupportedHostToolError";
    this.hostKind = hostKind;
    this.toolName = toolName;
  }
}

export function isUnsupportedHostTool(tool: AgentTool<TSchema, unknown>): boolean {
  return (tool as AgentTool<TSchema, unknown> & Partial<UnsupportedHostToolMarker>)[UNSUPPORTED_HOST_TOOL_MARKER] === true;
}

export function createUnsupportedHostTool<TParameters extends TSchema, TDetails>(
  tool: AgentTool<TParameters, TDetails>,
  hostKind: SpreadsheetHostKind,
): AgentTool<TParameters, TDetails> {
  const unsupportedTool: AgentTool<TParameters, TDetails> = {
    ...tool,
    execute: () => {
      throw new UnsupportedHostToolError(hostKind, tool.name);
    },
  };

  Object.defineProperty(unsupportedTool, UNSUPPORTED_HOST_TOOL_MARKER, {
    value: true,
    enumerable: true,
  });

  return unsupportedTool;
}
