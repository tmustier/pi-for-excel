import type { Context, Tool } from "@mariozechner/pi-ai";

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

/**
 * Select a deterministic tool bundle for the current call.
 *
 * Cache-first policy:
 * - If tools are present, always expose the full tool list.
 * - We keep core-only and mixed toolsets aligned to the same stable tool schema
 *   so prompt caches are not partitioned by intent-routed bundles.
 */
export function selectToolBundle(context: Context): ToolDisclosureResult {
  if (!context.tools || context.tools.length === 0) {
    return { tools: context.tools, bundleId: "none" };
  }

  // Keep the lightweight core-tool check to make intent explicit and to preserve
  // future room for opt-in routing without reintroducing local tool-name lists.
  if (!hasOnlyCoreTools(context.tools)) {
    return { tools: context.tools, bundleId: "full" };
  }

  return { tools: context.tools, bundleId: "full" };
}
