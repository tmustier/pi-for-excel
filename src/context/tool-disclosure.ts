import type { Context, Tool } from "@mariozechner/pi-ai";

import { CORE_TOOL_NAMES, type CoreToolName } from "../tools/registry.js";

export type ToolBundleId = "none" | "core" | "analysis" | "formatting" | "structure" | "comments" | "full";

type ActiveToolBundleId = Exclude<ToolBundleId, "none">;
type UserMessage = Extract<Context["messages"][number], { role: "user" }>;

const CORE_TOOL_NAME_SET = new Set<string>(CORE_TOOL_NAMES);

const TOOL_BUNDLES = {
  core: [
    "get_workbook_overview",
    "read_range",
    "search_workbook",
    "write_cells",
    "fill_formula",
    "instructions",
    "conventions",
    "workbook_history",
  ],
  analysis: [
    "get_workbook_overview",
    "read_range",
    "search_workbook",
    "write_cells",
    "fill_formula",
    "trace_dependencies",
    "instructions",
    "conventions",
    "workbook_history",
  ],
  formatting: [
    "get_workbook_overview",
    "read_range",
    "search_workbook",
    "write_cells",
    "fill_formula",
    "format_cells",
    "conditional_format",
    "view_settings",
    "instructions",
    "conventions",
    "workbook_history",
  ],
  structure: [
    "get_workbook_overview",
    "read_range",
    "search_workbook",
    "write_cells",
    "fill_formula",
    "modify_structure",
    "view_settings",
    "instructions",
    "conventions",
    "workbook_history",
  ],
  comments: [
    "get_workbook_overview",
    "read_range",
    "search_workbook",
    "write_cells",
    "fill_formula",
    "comments",
    "instructions",
    "conventions",
    "workbook_history",
  ],
  full: CORE_TOOL_NAMES,
} as const satisfies Record<ActiveToolBundleId, readonly CoreToolName[]>;

const FULL_ACCESS_PATTERNS: readonly RegExp[] = [
  /\ball tools?\b/,
  /\bany tools?\b/,
  /\bfull tool(set)?\b/,
  /\bfull access\b/,
  /\buse whatever tools?\b/,
];

const COMMENT_PATTERNS: readonly RegExp[] = [
  /\bcomment(s)?\b/,
  /\breply\b/,
  /\bthread(s)?\b/,
  /\bresolve\b/,
  /\bannotation(s)?\b/,
];

const DEPENDENCY_PATTERNS: readonly RegExp[] = [
  /\btrace\b/,
  /\bprecedent(s)?\b/,
  /\bdependent(s)?\b/,
  /\bdependenc(y|ies)\b/,
  /\blineage\b/,
  /\bformula (audit|debug|explain)\b/,
];

const STRUCTURE_PATTERNS: readonly RegExp[] = [
  /\b(insert|delete|rename|move|shift)\b[^\n]{0,40}\b(row|rows|column|columns|sheet|sheets|tab|tabs)\b/,
  /\b(add|remove)\b[^\n]{0,20}\b(sheet|sheets|tab|tabs)\b/,
  /\bhide\b[^\n]{0,20}\b(sheet|sheets|tab|tabs)\b/,
  /\bunhide\b[^\n]{0,20}\b(sheet|sheets|tab|tabs)\b/,
  /\bfreeze panes?\b/,
  /\bgridlines?\b/,
  /\bheadings?\b/,
  /\btab color\b/,
];

const FORMATTING_PATTERNS: readonly RegExp[] = [
  /\bformat(ting)?\b/,
  /\bstyle(s)?\b/,
  /\bbold\b/,
  /\bborder(s)?\b/,
  /\bfont\b/,
  /\bfill\b/,
  /\bcolor(s)?\b/,
  /\bhighlight\b/,
  /\bconditional format(ting)?\b/,
  /\bnumber format\b/,
  /\bcurrency\b/,
  /\bpercent(age)?\b/,
  /\bdecimal(s)?\b/,
  /\balignment\b/,
  /\bwrap text\b/,
];

function isCoreToolName(name: string): name is CoreToolName {
  return CORE_TOOL_NAME_SET.has(name);
}

function hasOnlyCoreTools(tools: readonly Tool[]): boolean {
  for (const tool of tools) {
    if (!isCoreToolName(tool.name)) return false;
  }
  return true;
}

function matchesAny(text: string, patterns: readonly RegExp[]): boolean {
  return patterns.some((pattern) => pattern.test(text));
}

function extractUserText(content: UserMessage["content"]): string {
  if (typeof content === "string") return content;

  const textParts: string[] = [];
  for (const item of content) {
    if (item.type === "text") {
      textParts.push(item.text);
    }
  }

  return textParts.join(" ");
}

function getLastUserPrompt(messages: Context["messages"]): string | null {
  for (let i = messages.length - 1; i >= 0; i -= 1) {
    const message = messages[i];
    if (message.role !== "user") continue;

    const text = extractUserText(message.content).trim();
    if (text.length === 0) continue;

    if (text.startsWith("[Auto-context]")) continue;
    return text.toLowerCase();
  }

  return null;
}

function chooseBundle(prompt: string): ActiveToolBundleId {
  if (matchesAny(prompt, FULL_ACCESS_PATTERNS)) return "full";

  const matchedBundles: ActiveToolBundleId[] = [];

  if (matchesAny(prompt, COMMENT_PATTERNS)) matchedBundles.push("comments");
  if (matchesAny(prompt, DEPENDENCY_PATTERNS)) matchedBundles.push("analysis");
  if (matchesAny(prompt, STRUCTURE_PATTERNS)) matchedBundles.push("structure");
  if (matchesAny(prompt, FORMATTING_PATTERNS)) matchedBundles.push("formatting");

  // Mixed-intent requests (e.g. "insert a row and highlight it") need tools
  // across categories. Fall back to full for the first call so continuation
  // stripping doesn't block capabilities in the same turn.
  if (matchedBundles.length > 1) return "full";
  if (matchedBundles.length === 1) return matchedBundles[0];
  return "core";
}

function filterToolsByBundle(tools: readonly Tool[], bundleId: ActiveToolBundleId): Tool[] {
  if (bundleId === "full") return [...tools];

  const allowed = new Set<string>(TOOL_BUNDLES[bundleId]);
  const filtered = tools.filter((tool) => allowed.has(tool.name));
  return filtered.length > 0 ? filtered : [...tools];
}

export interface ToolDisclosureResult {
  tools: Context["tools"];
  bundleId: ToolBundleId;
}

/**
 * Select a deterministic tool bundle for the current call.
 *
 * Rules:
 * - Only applies to the core built-in tool set.
 * - If extension/non-core tools are present, keep full tool visibility.
 * - Selection is based on the latest non-auto user prompt.
 */
export function selectToolBundle(context: Context): ToolDisclosureResult {
  if (!context.tools || context.tools.length === 0) {
    return { tools: context.tools, bundleId: "none" };
  }

  if (!hasOnlyCoreTools(context.tools)) {
    return { tools: context.tools, bundleId: "full" };
  }

  const prompt = getLastUserPrompt(context.messages);
  const bundleId = prompt ? chooseBundle(prompt) : "core";
  const tools = filterToolsByBundle(context.tools, bundleId);
  return { tools, bundleId };
}
