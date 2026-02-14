import { CORE_TOOL_NAMES, type CoreToolName } from "./names.js";

export type ToolCapabilityTier =
  | "core"
  | "on_demand_tier_1"
  | "on_demand_tier_2"
  | "on_demand_tier_3"
  | "experimental";

export type CoreToolCapabilityCategory =
  | "read"
  | "write"
  | "navigate"
  | "structure"
  | "format"
  | "inspect"
  | "view"
  | "collaboration"
  | "instructions"
  | "recovery"
  | "skills";

interface CoreToolCapabilityMetadata {
  tier: "core";
  category: CoreToolCapabilityCategory;
  promptDescription: string;
}

export interface CoreToolCapability extends CoreToolCapabilityMetadata {
  name: CoreToolName;
}

const CORE_TOOL_CAPABILITY_METADATA = {
  get_workbook_overview: {
    tier: "core",
    category: "read",
    promptDescription: "structural blueprint (sheets, headers, named ranges, tables); optional sheet-level detail for charts, pivots, shapes",
  },
  read_range: {
    tier: "core",
    category: "read",
    promptDescription: "read cell values/formulas in three formats: compact (markdown), csv (values-only), or detailed (with formatting + comments)",
  },
  write_cells: {
    tier: "core",
    category: "write",
    promptDescription: "write values/formulas with overwrite protection and auto-verification",
  },
  fill_formula: {
    tier: "core",
    category: "write",
    promptDescription: "fill a single formula across a range (AutoFill with relative refs)",
  },
  search_workbook: {
    tier: "core",
    category: "navigate",
    promptDescription: "find text, values, or formula references across all sheets; context_rows for surrounding data",
  },
  modify_structure: {
    tier: "core",
    category: "structure",
    promptDescription: "insert/delete rows/columns, add/rename/delete sheets",
  },
  format_cells: {
    tier: "core",
    category: "format",
    promptDescription: "apply formatting (bold, colors, number format, borders, etc.)",
  },
  conditional_format: {
    tier: "core",
    category: "format",
    promptDescription: "add or clear conditional formatting rules (formula or cell-value)",
  },
  trace_dependencies: {
    tier: "core",
    category: "inspect",
    promptDescription: "trace formula lineage for a cell (mode: `precedents` upstream or `dependents` downstream)",
  },
  explain_formula: {
    tier: "core",
    category: "inspect",
    promptDescription: "explain a single formula cell in plain language with cited direct references",
  },
  view_settings: {
    tier: "core",
    category: "view",
    promptDescription: "control gridlines, headings, freeze panes, tab color, sheet visibility, sheet activation, and standard width",
  },
  comments: {
    tier: "core",
    category: "collaboration",
    promptDescription: "read, add, update, reply, delete, resolve/reopen cell comments",
  },
  instructions: {
    tier: "core",
    category: "instructions",
    promptDescription: "update persistent rules for all files or this file (append or replace)",
  },
  conventions: {
    tier: "core",
    category: "instructions",
    promptDescription: "read/update formatting defaults (currency, negatives, zeros, decimal places)",
  },
  workbook_history: {
    tier: "core",
    category: "recovery",
    promptDescription: "list/restore/delete automatic backups created before Pi edits for supported workbook mutations (`write_cells`, `fill_formula`, `python_transform_range`, `format_cells`, `conditional_format`, `comments`, and supported `modify_structure` actions)",
  },
  skills: {
    tier: "core",
    category: "skills",
    promptDescription: "list/read bundled Agent Skills (SKILL.md) for task-specific workflows",
  },
} satisfies Record<CoreToolName, CoreToolCapabilityMetadata>;

export const CORE_TOOL_CAPABILITIES: readonly CoreToolCapability[] = CORE_TOOL_NAMES.map((name) => ({
  name,
  ...CORE_TOOL_CAPABILITY_METADATA[name],
}));

export function buildCoreToolPromptLines(): string {
  return CORE_TOOL_CAPABILITIES
    .map((capability) => `- **${capability.name}** — ${capability.promptDescription}`)
    .join("\n");
}

export type ToolDisclosureBundleId = "none" | "core" | "analysis" | "formatting" | "structure" | "comments" | "full";

type ActiveToolDisclosureBundleId = Exclude<ToolDisclosureBundleId, "none">;

type TriggeredToolDisclosureBundleId = Exclude<ToolDisclosureBundleId, "none" | "core" | "full">;

const TOOL_DISCLOSURE_CATEGORY_SETS = {
  core: ["read", "write", "navigate", "instructions", "recovery", "skills"],
  analysis: ["read", "write", "navigate", "inspect", "instructions", "recovery", "skills"],
  formatting: ["read", "write", "navigate", "format", "view", "instructions", "recovery", "skills"],
  structure: ["read", "write", "navigate", "structure", "view", "instructions", "recovery", "skills"],
  comments: ["read", "write", "navigate", "collaboration", "instructions", "recovery", "skills"],
} as const satisfies Record<TriggeredToolDisclosureBundleId | "core", readonly CoreToolCapabilityCategory[]>;

function buildCoreDisclosureBundle(categorySet: readonly CoreToolCapabilityCategory[]): readonly CoreToolName[] {
  const allowedCategories = new Set<CoreToolCapabilityCategory>(categorySet);

  return CORE_TOOL_CAPABILITIES
    .filter((capability) => allowedCategories.has(capability.category))
    .map((capability) => capability.name);
}

export const TOOL_DISCLOSURE_BUNDLES = {
  core: buildCoreDisclosureBundle(TOOL_DISCLOSURE_CATEGORY_SETS.core),
  analysis: buildCoreDisclosureBundle(TOOL_DISCLOSURE_CATEGORY_SETS.analysis),
  formatting: buildCoreDisclosureBundle(TOOL_DISCLOSURE_CATEGORY_SETS.formatting),
  structure: buildCoreDisclosureBundle(TOOL_DISCLOSURE_CATEGORY_SETS.structure),
  comments: buildCoreDisclosureBundle(TOOL_DISCLOSURE_CATEGORY_SETS.comments),
  full: CORE_TOOL_NAMES,
} as const satisfies Record<ActiveToolDisclosureBundleId, readonly CoreToolName[]>;

export const TOOL_DISCLOSURE_FULL_ACCESS_PATTERNS: readonly RegExp[] = [
  /\ball tools?\b/,
  /\bany tools?\b/,
  /\bfull tool(set)?\b/,
  /\bfull access\b/,
  /\buse whatever tools?\b/,
];

export const TOOL_DISCLOSURE_TRIGGER_PATTERNS = {
  comments: [
    /\bcomment(s)?\b/,
    /\breply\b/,
    /\bthread(s)?\b/,
    /\bresolve\b/,
    /\bannotation(s)?\b/,
  ],
  analysis: [
    /\btrace\b/,
    /\bprecedent(s)?\b/,
    /\bdependent(s)?\b/,
    /\bdependenc(y|ies)\b/,
    /\blineage\b/,
    /\bformula (audit|debug|explain)\b/,
  ],
  structure: [
    /\b(insert|delete|rename|move|shift)\b[^\n]{0,40}\b(row|rows|column|columns|sheet|sheets|tab|tabs)\b/,
    /\b(add|remove)\b[^\n]{0,20}\b(sheet|sheets|tab|tabs)\b/,
    /\bhide\b[^\n]{0,20}\b(sheet|sheets|tab|tabs)\b/,
    /\bunhide\b[^\n]{0,20}\b(sheet|sheets|tab|tabs)\b/,
    /\bfreeze panes?\b/,
    /\bgridlines?\b/,
    /\bheadings?\b/,
    /\btab color\b/,
  ],
  formatting: [
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
  ],
} as const satisfies Record<TriggeredToolDisclosureBundleId, readonly RegExp[]>;

export const TOOL_DISCLOSURE_TRIGGER_BUNDLE_ORDER = [
  "comments",
  "analysis",
  "structure",
  "formatting",
] as const satisfies readonly TriggeredToolDisclosureBundleId[];

function matchesAny(text: string, patterns: readonly RegExp[]): boolean {
  return patterns.some((pattern) => pattern.test(text));
}

export function chooseToolDisclosureBundle(prompt: string): ActiveToolDisclosureBundleId {
  if (matchesAny(prompt, TOOL_DISCLOSURE_FULL_ACCESS_PATTERNS)) return "full";

  const matchedBundles: TriggeredToolDisclosureBundleId[] = [];

  for (const bundleId of TOOL_DISCLOSURE_TRIGGER_BUNDLE_ORDER) {
    if (matchesAny(prompt, TOOL_DISCLOSURE_TRIGGER_PATTERNS[bundleId])) {
      matchedBundles.push(bundleId);
    }
  }

  // Mixed-intent requests (e.g. "insert a row and highlight it") need tools
  // across categories. Fall back to full for the first call so continuation
  // stripping doesn't block capabilities in the same turn.
  if (matchedBundles.length > 1) return "full";
  if (matchedBundles.length === 1) return matchedBundles[0];
  return "core";
}

export function filterToolsForDisclosureBundle<TTool extends { name: string }>(
  tools: readonly TTool[],
  bundleId: ActiveToolDisclosureBundleId,
): TTool[] {
  if (bundleId === "full") return [...tools];

  const allowed = new Set<string>(TOOL_DISCLOSURE_BUNDLES[bundleId]);
  const filtered = tools.filter((tool) => allowed.has(tool.name));
  return filtered.length > 0 ? filtered : [...tools];
}

export const AUXILIARY_UI_TOOL_NAMES = [
  "web_search",
  "fetch_page",
  "mcp",
  "files",
  "python_transform_range",
  "execute_office_js",
] as const;

export type AuxiliaryUiToolName = (typeof AUXILIARY_UI_TOOL_NAMES)[number];

export type UiToolName = CoreToolName | AuxiliaryUiToolName;

export const UI_TOOL_NAMES: readonly UiToolName[] = [
  ...CORE_TOOL_NAMES,
  ...AUXILIARY_UI_TOOL_NAMES,
];

export interface ToolUiMetadata {
  renderer: boolean;
  humanizer: boolean;
}

export const TOOL_UI_METADATA = {
  get_workbook_overview: { renderer: true, humanizer: true },
  read_range: { renderer: true, humanizer: true },
  write_cells: { renderer: true, humanizer: true },
  fill_formula: { renderer: true, humanizer: true },
  search_workbook: { renderer: true, humanizer: true },
  modify_structure: { renderer: true, humanizer: true },
  format_cells: { renderer: true, humanizer: true },
  conditional_format: { renderer: true, humanizer: true },
  trace_dependencies: { renderer: true, humanizer: true },
  explain_formula: { renderer: true, humanizer: true },
  view_settings: { renderer: true, humanizer: true },
  comments: { renderer: true, humanizer: true },
  instructions: { renderer: true, humanizer: true },
  conventions: { renderer: true, humanizer: true },
  workbook_history: { renderer: true, humanizer: true },
  skills: { renderer: true, humanizer: true },
  web_search: { renderer: true, humanizer: true },
  fetch_page: { renderer: true, humanizer: true },
  mcp: { renderer: true, humanizer: true },
  files: { renderer: true, humanizer: true },
  python_transform_range: { renderer: true, humanizer: true },
  execute_office_js: { renderer: true, humanizer: true },
} as const satisfies Record<UiToolName, ToolUiMetadata>;

export const TOOL_NAMES_WITH_RENDERER: readonly UiToolName[] = UI_TOOL_NAMES
  .filter((name) => TOOL_UI_METADATA[name].renderer);

export const TOOL_NAMES_WITH_HUMANIZER: readonly UiToolName[] = UI_TOOL_NAMES
  .filter((name) => TOOL_UI_METADATA[name].humanizer);
