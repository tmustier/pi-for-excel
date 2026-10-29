import type { Context, Tool } from "@earendil-works/pi-ai";

import { type CoreToolName, CORE_TOOL_NAMES } from "../tools/names.js";
import { type ToolDisclosureBundleId } from "../tools/capabilities.js";

export type ToolBundleId = ToolDisclosureBundleId;

const CORE_TOOL_NAME_SET = new Set<string>(CORE_TOOL_NAMES);

function isCoreToolName(name: string): name is CoreToolName {
  return CORE_TOOL_NAME_SET.has(name);
}

function hasOnlyCoreTools(tools: readonly Tool[]): boolean {
  for (const tool of tools) {
    if (!isCoreToolName(tool.name)) return false;
  }
  return true;
}

export interface ToolDisclosureResult {
  tools: Context["tools"];
  bundleId: ToolBundleId;
}

export function selectToolBundle(context: Context): ToolDisclosureResult {
  if (!context.tools || context.tools.length === 0) {
    return { tools: context.tools, bundleId: "none" };
  }

  if (!hasOnlyCoreTools(context.tools)) {
    return { tools: context.tools, bundleId: "full" };
  }

  return { tools: context.tools, bundleId: "full" };
}
