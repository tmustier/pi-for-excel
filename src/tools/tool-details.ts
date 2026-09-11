/**
 * Structured tool result metadata for the UI.
 *
 * Tools still return human-readable markdown in `content`, but also attach a
 * small stable `details` payload so the UI doesn't need to parse strings.
 *
 * Each payload is a TypeBox schema; the exported types are `Static<>` of the
 * schema so the type and the runtime check cannot drift. `details` reaches the
 * UI as `unknown` (it may come from an older persisted session), so the UI
 * decodes it once with `decodeToolDetails` and switches on `kind` from there.
 */

import { Type, type Static, type TSchema } from "typebox";
import { Value } from "typebox/value";

import { workbookCellChangeSummarySchema } from "../audit/cell-diff.js";
import { connectionToolErrorDetailsSchema } from "../connections/types.js";
import { StringEnum } from "./string-enum.js";

export const recoveryCheckpointDetailsSchema = Type.Object({
  status: StringEnum(["checkpoint_created", "not_available"]),
  snapshotId: Type.Optional(Type.String()),
  reason: Type.Optional(Type.String()),
});

export type RecoveryCheckpointDetails = Static<typeof recoveryCheckpointDetailsSchema>;

const toolOutputTruncationStrategySchema = StringEnum(["head", "tail"]);

export type ToolOutputTruncationStrategy = Static<typeof toolOutputTruncationStrategySchema>;

const toolOutputTruncationReasonSchema = Type.Union([
  Type.Literal("lines"),
  Type.Literal("bytes"),
  Type.Null(),
]);

export type ToolOutputTruncationReason = Static<typeof toolOutputTruncationReasonSchema>;

export const toolOutputTruncationDetailsSchema = Type.Object({
  version: Type.Literal(1),
  strategy: toolOutputTruncationStrategySchema,
  truncated: Type.Boolean(),
  truncatedBy: toolOutputTruncationReasonSchema,
  totalLines: Type.Number(),
  totalBytes: Type.Number(),
  outputLines: Type.Number(),
  outputBytes: Type.Number(),
  maxLines: Type.Number(),
  maxBytes: Type.Number(),
  fullOutputWorkspacePath: Type.Optional(Type.String()),
});

export type ToolOutputTruncationDetails = Static<typeof toolOutputTruncationDetailsSchema>;

/** Truncation is merged into whatever details a tool returned, alongside its own fields. */
const truncatedToolDetailsSchema = Type.Object({
  outputTruncation: toolOutputTruncationDetailsSchema,
});

const mutationDetailsProperties = {
  blocked: Type.Boolean(),
  /** Sheet-qualified range when known, e.g. "Sheet1!A1:C3" */
  address: Type.Optional(Type.String()),
  existingCount: Type.Optional(Type.Number()),
  formulaErrorCount: Type.Optional(Type.Number()),
  changes: Type.Optional(workbookCellChangeSummarySchema),
  recovery: Type.Optional(recoveryCheckpointDetailsSchema),
};

export const writeCellsDetailsSchema = Type.Object({
  kind: Type.Literal("write_cells"),
  ...mutationDetailsProperties,
});

export type WriteCellsDetails = Static<typeof writeCellsDetailsSchema>;

export const fillFormulaDetailsSchema = Type.Object({
  kind: Type.Literal("fill_formula"),
  ...mutationDetailsProperties,
});

export type FillFormulaDetails = Static<typeof fillFormulaDetailsSchema>;

export const formatCellsDetailsSchema = Type.Object({
  kind: Type.Literal("format_cells"),
  /** Sheet-qualified range when known. May be a multi-range string. */
  address: Type.Optional(Type.String()),
  warningsCount: Type.Optional(Type.Number()),
  recovery: Type.Optional(recoveryCheckpointDetailsSchema),
});

export type FormatCellsDetails = Static<typeof formatCellsDetailsSchema>;

export const conditionalFormatDetailsSchema = Type.Object({
  kind: Type.Literal("conditional_format"),
  action: Type.Optional(StringEnum(["add", "clear"])),
  address: Type.Optional(Type.String()),
  recovery: Type.Optional(recoveryCheckpointDetailsSchema),
});

export type ConditionalFormatDetails = Static<typeof conditionalFormatDetailsSchema>;

export const modifyStructureDetailsSchema = Type.Object({
  kind: Type.Literal("modify_structure"),
  action: Type.Optional(Type.String()),
  recovery: Type.Optional(recoveryCheckpointDetailsSchema),
});

export type ModifyStructureDetails = Static<typeof modifyStructureDetailsSchema>;

export const commentsDetailsSchema = Type.Object({
  kind: Type.Literal("comments"),
  action: Type.Optional(Type.String()),
  address: Type.Optional(Type.String()),
  recovery: Type.Optional(recoveryCheckpointDetailsSchema),
});

export type CommentsDetails = Static<typeof commentsDetailsSchema>;

export const viewSettingsDetailsSchema = Type.Object({
  kind: Type.Literal("view_settings"),
  action: Type.Optional(Type.String()),
  address: Type.Optional(Type.String()),
  recovery: Type.Optional(recoveryCheckpointDetailsSchema),
});

export type ViewSettingsDetails = Static<typeof viewSettingsDetailsSchema>;

export const chartPositionDetailsSchema = Type.Object({
  top: Type.Number(),
  left: Type.Number(),
  width: Type.Number(),
  height: Type.Number(),
});

export type ChartPositionDetails = Static<typeof chartPositionDetailsSchema>;

export const chartListItemDetailsSchema = Type.Object({
  name: Type.String(),
  chartType: Type.String(),
  title: Type.String(),
  worksheet: Type.String(),
  position: chartPositionDetailsSchema,
});

export type ChartListItemDetails = Static<typeof chartListItemDetailsSchema>;

export const chartImageDetailsSchema = Type.Object({
  base64: Type.String(),
  mimeType: Type.Literal("image/png"),
  width: Type.Number(),
  height: Type.Number(),
});

export type ChartImageDetails = Static<typeof chartImageDetailsSchema>;

export const chartsDetailsSchema = Type.Object({
  kind: Type.Literal("charts"),
  action: Type.Optional(Type.String()),
  name: Type.Optional(Type.String()),
  address: Type.Optional(Type.String()),
  sourceRange: Type.Optional(Type.String()),
  count: Type.Optional(Type.Number()),
  charts: Type.Optional(Type.Array(chartListItemDetailsSchema)),
  image: Type.Optional(chartImageDetailsSchema),
  recovery: Type.Optional(recoveryCheckpointDetailsSchema),
});

export type ChartsDetails = Static<typeof chartsDetailsSchema>;

const traceDependenciesModeSchema = StringEnum(["precedents", "dependents"]);

export type TraceDependenciesMode = Static<typeof traceDependenciesModeSchema>;

const traceDependencySourceSchema = StringEnum(["api", "formula_scan", "mixed", "none"]);

export type TraceDependencySource = Static<typeof traceDependencySourceSchema>;

/**
 * A cell value as the host returned it. `range.values` is untyped in Office.js
 * and WPS; typing it belongs to the host adapters, after which this narrows.
 */
const hostCellValueSchema = Type.Unknown();

export const depNodeDetailSchema = Type.Cyclic({
  DepNode: Type.Object({
    address: Type.String(),
    value: hostCellValueSchema,
    /** Excel number format string, e.g. "0.00%", "#,##0", "$#,##0.00". */
    numberFormat: Type.Optional(Type.String()),
    formula: Type.Optional(Type.String()),
    /** Child nodes in traversal order (precedents for precedents mode, dependents for dependents mode). */
    precedents: Type.Array(Type.Ref("DepNode")),
  }),
}, "DepNode");

export type DepNodeDetail = Static<typeof depNodeDetailSchema>;

export const traceDependenciesDetailsSchema = Type.Object({
  kind: Type.Literal("trace_dependencies"),
  root: depNodeDetailSchema,
  mode: Type.Optional(traceDependenciesModeSchema),
  maxDepth: Type.Optional(Type.Number()),
  nodeCount: Type.Optional(Type.Number()),
  edgeCount: Type.Optional(Type.Number()),
  source: Type.Optional(traceDependencySourceSchema),
  truncated: Type.Optional(Type.Boolean()),
});

export type TraceDependenciesDetails = Static<typeof traceDependenciesDetailsSchema>;

export const explainFormulaReferenceDetailSchema = Type.Object({
  address: Type.String(),
  valuePreview: Type.Optional(Type.String()),
  formulaPreview: Type.Optional(Type.String()),
});

export type ExplainFormulaReferenceDetail = Static<typeof explainFormulaReferenceDetailSchema>;

export const explainFormulaDetailsSchema = Type.Object({
  kind: Type.Literal("explain_formula"),
  cell: Type.String(),
  hasFormula: Type.Boolean(),
  formula: Type.Optional(Type.String()),
  valuePreview: Type.Optional(Type.String()),
  explanation: Type.String(),
  references: Type.Array(explainFormulaReferenceDetailSchema),
  truncated: Type.Optional(Type.Boolean()),
});

export type ExplainFormulaDetails = Static<typeof explainFormulaDetailsSchema>;

export const readRangeCsvDetailsSchema = Type.Object({
  kind: Type.Literal("read_range_csv"),
  /** 0-indexed starting column (A=0, B=1, …) */
  startCol: Type.Number(),
  /** 1-indexed starting row */
  startRow: Type.Number(),
  /** Raw values grid from the host */
  values: Type.Array(Type.Array(hostCellValueSchema)),
  /** Pre-serialized CSV string for the copy button */
  csv: Type.String(),
});

export type ReadRangeCsvDetails = Static<typeof readRangeCsvDetailsSchema>;

const bridgeGateReasonSchema = StringEnum(["missing_bridge_url", "invalid_bridge_url", "bridge_unreachable"]);

export type BridgeGateReason = Static<typeof bridgeGateReasonSchema>;

const bridgeGateProperties = {
  error: Type.Optional(Type.String()),
  gateReason: Type.Optional(bridgeGateReasonSchema),
  skillHint: Type.Optional(Type.String()),
};

export const tmuxBridgeDetailsSchema = Type.Object({
  kind: Type.Literal("tmux_bridge"),
  ok: Type.Boolean(),
  action: Type.String(),
  bridgeUrl: Type.Optional(Type.String()),
  session: Type.Optional(Type.String()),
  sessionsCount: Type.Optional(Type.Number()),
  outputPreview: Type.Optional(Type.String()),
  ...bridgeGateProperties,
});

export type TmuxBridgeDetails = Static<typeof tmuxBridgeDetailsSchema>;

export const pythonBridgeDetailsSchema = Type.Object({
  kind: Type.Literal("python_bridge"),
  ok: Type.Boolean(),
  action: Type.String(),
  bridgeUrl: Type.Optional(Type.String()),
  exitCode: Type.Optional(Type.Number()),
  stdoutPreview: Type.Optional(Type.String()),
  stderrPreview: Type.Optional(Type.String()),
  resultPreview: Type.Optional(Type.String()),
  truncated: Type.Optional(Type.Boolean()),
  ...bridgeGateProperties,
});

export type PythonBridgeDetails = Static<typeof pythonBridgeDetailsSchema>;

export const libreOfficeBridgeDetailsSchema = Type.Object({
  kind: Type.Literal("libreoffice_bridge"),
  ok: Type.Boolean(),
  action: Type.String(),
  bridgeUrl: Type.Optional(Type.String()),
  inputPath: Type.Optional(Type.String()),
  targetFormat: Type.Optional(Type.String()),
  outputPath: Type.Optional(Type.String()),
  bytes: Type.Optional(Type.Number()),
  converter: Type.Optional(Type.String()),
  ...bridgeGateProperties,
});

export type LibreOfficeBridgeDetails = Static<typeof libreOfficeBridgeDetailsSchema>;

export const pythonTransformRangeDetailsSchema = Type.Object({
  kind: Type.Literal("python_transform_range"),
  blocked: Type.Boolean(),
  inputAddress: Type.Optional(Type.String()),
  outputAddress: Type.Optional(Type.String()),
  bridgeUrl: Type.Optional(Type.String()),
  existingCount: Type.Optional(Type.Number()),
  rowsWritten: Type.Optional(Type.Number()),
  colsWritten: Type.Optional(Type.Number()),
  formulaErrorCount: Type.Optional(Type.Number()),
  changes: Type.Optional(workbookCellChangeSummarySchema),
  recovery: Type.Optional(recoveryCheckpointDetailsSchema),
  ...bridgeGateProperties,
});

export type PythonTransformRangeDetails = Static<typeof pythonTransformRangeDetailsSchema>;

export const workbookHistorySnapshotSummarySchema = Type.Object({
  id: Type.String(),
  at: Type.Number(),
  toolName: Type.String(),
  address: Type.String(),
  changedCount: Type.Number(),
  cellCount: Type.Number(),
  workbookId: Type.Optional(Type.String()),
  workbookLabel: Type.Optional(Type.String()),
});

export type WorkbookHistorySnapshotSummary = Static<typeof workbookHistorySnapshotSummarySchema>;

export const workbookHistoryDetailsSchema = Type.Object({
  kind: Type.Literal("workbook_history"),
  action: StringEnum(["list", "restore", "delete", "clear"]),
  count: Type.Optional(Type.Number()),
  snapshots: Type.Optional(Type.Array(workbookHistorySnapshotSummarySchema)),
  snapshotId: Type.Optional(Type.String()),
  restoredSnapshotId: Type.Optional(Type.String()),
  inverseSnapshotId: Type.Optional(Type.String()),
  address: Type.Optional(Type.String()),
  changedCount: Type.Optional(Type.Number()),
  deletedCount: Type.Optional(Type.Number()),
  error: Type.Optional(Type.String()),
});

export type WorkbookHistoryDetails = Static<typeof workbookHistoryDetailsSchema>;

const skillsSourceKindSchema = StringEnum(["bundled", "external"]);

export type SkillsSourceKind = Static<typeof skillsSourceKindSchema>;

export const skillsListEntryDetailsSchema = Type.Object({
  name: Type.String(),
  sourceKind: skillsSourceKindSchema,
  location: Type.String(),
});

export type SkillsListEntryDetails = Static<typeof skillsListEntryDetailsSchema>;

export const skillsListDetailsSchema = Type.Object({
  kind: Type.Literal("skills_list"),
  count: Type.Number(),
  names: Type.Array(Type.String()),
  entries: Type.Array(skillsListEntryDetailsSchema),
  externalDiscoveryEnabled: Type.Boolean(),
});

export type SkillsListDetails = Static<typeof skillsListDetailsSchema>;

export const skillsReadDetailsSchema = Type.Object({
  kind: Type.Literal("skills_read"),
  skillName: Type.String(),
  sourceKind: skillsSourceKindSchema,
  location: Type.String(),
  cacheHit: Type.Boolean(),
  refreshed: Type.Boolean(),
  sessionScoped: Type.Boolean(),
  readCount: Type.Optional(Type.Number()),
});

export type SkillsReadDetails = Static<typeof skillsReadDetailsSchema>;

export const skillsInstallDetailsSchema = Type.Object({
  kind: Type.Literal("skills_install"),
  skillName: Type.String(),
  location: Type.String(),
});

export type SkillsInstallDetails = Static<typeof skillsInstallDetailsSchema>;

export const skillsUninstallDetailsSchema = Type.Object({
  kind: Type.Literal("skills_uninstall"),
  skillName: Type.String(),
  removed: Type.Boolean(),
});

export type SkillsUninstallDetails = Static<typeof skillsUninstallDetailsSchema>;

export const skillsErrorDetailsSchema = Type.Object({
  kind: Type.Literal("skills_error"),
  action: StringEnum(["read", "install", "uninstall"]),
  message: Type.String(),
  requestedName: Type.Optional(Type.String()),
  availableNames: Type.Optional(Type.Array(Type.String())),
  externalDiscoveryEnabled: Type.Boolean(),
});

export type SkillsErrorDetails = Static<typeof skillsErrorDetailsSchema>;

export type SkillsToolDetails =
  | SkillsListDetails
  | SkillsReadDetails
  | SkillsInstallDetails
  | SkillsUninstallDetails
  | SkillsErrorDetails;

export const webSearchFallbackDetailsSchema = Type.Object({
  fromProvider: Type.String(),
  toProvider: Type.String(),
  reason: Type.String(),
});

export type WebSearchFallbackDetails = Static<typeof webSearchFallbackDetailsSchema>;

export const webSearchDetailsSchema = Type.Object({
  kind: Type.Literal("web_search"),
  ok: Type.Boolean(),
  provider: Type.String(),
  query: Type.String(),
  sentQuery: Type.String(),
  recency: Type.Optional(Type.String()),
  siteFilters: Type.Optional(Type.Array(Type.String())),
  maxResults: Type.Number(),
  resultCount: Type.Optional(Type.Number()),
  proxied: Type.Optional(Type.Boolean()),
  proxyBaseUrl: Type.Optional(Type.String()),
  fallback: Type.Optional(webSearchFallbackDetailsSchema),
  error: Type.Optional(Type.String()),
  /** `true` when the failure is due to the local CORS proxy being unreachable. */
  proxyDown: Type.Optional(Type.Boolean()),
});

export type WebSearchDetails = Static<typeof webSearchDetailsSchema>;

export const fetchPageDetailsSchema = Type.Object({
  kind: Type.Literal("fetch_page"),
  ok: Type.Boolean(),
  url: Type.String(),
  title: Type.Optional(Type.String()),
  chars: Type.Optional(Type.Number()),
  truncated: Type.Optional(Type.Boolean()),
  proxied: Type.Optional(Type.Boolean()),
  proxyBaseUrl: Type.Optional(Type.String()),
  contentType: Type.Optional(Type.String()),
  error: Type.Optional(Type.String()),
  /** `true` when the failure is due to the local CORS proxy being unreachable. */
  proxyDown: Type.Optional(Type.Boolean()),
});

export type FetchPageDetails = Static<typeof fetchPageDetailsSchema>;

export const mcpGatewayDetailsSchema = Type.Object({
  kind: Type.Literal("mcp_gateway"),
  ok: Type.Boolean(),
  operation: Type.String(),
  server: Type.Optional(Type.String()),
  tool: Type.Optional(Type.String()),
  proxied: Type.Optional(Type.Boolean()),
  proxyBaseUrl: Type.Optional(Type.String()),
  resultPreview: Type.Optional(Type.String()),
  error: Type.Optional(Type.String()),
  /** `true` when the failure is due to the local CORS proxy being unreachable. */
  proxyDown: Type.Optional(Type.Boolean()),
});

export type McpGatewayDetails = Static<typeof mcpGatewayDetailsSchema>;

const filesWorkspaceBackendKindSchema = StringEnum(["native-directory", "opfs", "memory"]);

export type FilesWorkspaceBackendKind = Static<typeof filesWorkspaceBackendKindSchema>;

export const filesWorkbookTagDetailsSchema = Type.Object({
  workbookId: Type.String(),
  workbookLabel: Type.String(),
  taggedAt: Type.Number(),
});

export type FilesWorkbookTagDetails = Static<typeof filesWorkbookTagDetailsSchema>;

const filesSourceKindSchema = StringEnum(["workspace", "builtin-doc"]);

export type FilesSourceKind = Static<typeof filesSourceKindSchema>;

const filesFileKindSchema = StringEnum(["text", "binary"]);

export const filesListItemDetailsSchema = Type.Object({
  path: Type.String(),
  size: Type.Number(),
  mimeType: Type.String(),
  fileKind: filesFileKindSchema,
  modifiedAt: Type.Number(),
  sourceKind: Type.Optional(filesSourceKindSchema),
  readOnly: Type.Optional(Type.Boolean()),
  workbookTag: Type.Optional(filesWorkbookTagDetailsSchema),
});

export type FilesListItemDetails = Static<typeof filesListItemDetailsSchema>;

export const filesListDetailsSchema = Type.Object({
  kind: Type.Literal("files_list"),
  backend: filesWorkspaceBackendKindSchema,
  count: Type.Number(),
  files: Type.Array(filesListItemDetailsSchema),
});

export type FilesListDetails = Static<typeof filesListDetailsSchema>;

export const filesReadDetailsSchema = Type.Object({
  kind: Type.Literal("files_read"),
  backend: filesWorkspaceBackendKindSchema,
  path: Type.String(),
  mode: StringEnum(["text", "base64"]),
  size: Type.Number(),
  mimeType: Type.String(),
  fileKind: filesFileKindSchema,
  sourceKind: Type.Optional(filesSourceKindSchema),
  readOnly: Type.Optional(Type.Boolean()),
  truncated: Type.Boolean(),
  workbookTag: Type.Optional(filesWorkbookTagDetailsSchema),
});

export type FilesReadDetails = Static<typeof filesReadDetailsSchema>;

export const filesWriteDetailsSchema = Type.Object({
  kind: Type.Literal("files_write"),
  backend: filesWorkspaceBackendKindSchema,
  path: Type.String(),
  encoding: StringEnum(["text", "base64"]),
  chars: Type.Number(),
  workbookTag: Type.Optional(filesWorkbookTagDetailsSchema),
});

export type FilesWriteDetails = Static<typeof filesWriteDetailsSchema>;

export const filesDeleteDetailsSchema = Type.Object({
  kind: Type.Literal("files_delete"),
  backend: filesWorkspaceBackendKindSchema,
  path: Type.String(),
  workbookTag: Type.Optional(filesWorkbookTagDetailsSchema),
});

export type FilesDeleteDetails = Static<typeof filesDeleteDetailsSchema>;

export type FilesToolDetails =
  | FilesListDetails
  | FilesReadDetails
  | FilesWriteDetails
  | FilesDeleteDetails;

/** Every details payload, keyed by its `kind` discriminator. */
const toolDetailsSchemasByKind = {
  write_cells: writeCellsDetailsSchema,
  fill_formula: fillFormulaDetailsSchema,
  format_cells: formatCellsDetailsSchema,
  conditional_format: conditionalFormatDetailsSchema,
  modify_structure: modifyStructureDetailsSchema,
  comments: commentsDetailsSchema,
  view_settings: viewSettingsDetailsSchema,
  charts: chartsDetailsSchema,
  trace_dependencies: traceDependenciesDetailsSchema,
  explain_formula: explainFormulaDetailsSchema,
  read_range_csv: readRangeCsvDetailsSchema,
  tmux_bridge: tmuxBridgeDetailsSchema,
  python_bridge: pythonBridgeDetailsSchema,
  libreoffice_bridge: libreOfficeBridgeDetailsSchema,
  python_transform_range: pythonTransformRangeDetailsSchema,
  workbook_history: workbookHistoryDetailsSchema,
  skills_list: skillsListDetailsSchema,
  skills_read: skillsReadDetailsSchema,
  skills_install: skillsInstallDetailsSchema,
  skills_uninstall: skillsUninstallDetailsSchema,
  skills_error: skillsErrorDetailsSchema,
  web_search: webSearchDetailsSchema,
  fetch_page: fetchPageDetailsSchema,
  mcp_gateway: mcpGatewayDetailsSchema,
  files_list: filesListDetailsSchema,
  files_read: filesReadDetailsSchema,
  files_write: filesWriteDetailsSchema,
  files_delete: filesDeleteDetailsSchema,
  connection_error: connectionToolErrorDetailsSchema,
} satisfies Record<string, TSchema>;

type ToolDetailsSchemasByKind = typeof toolDetailsSchemasByKind;

export type ToolDetailsKind = keyof ToolDetailsSchemasByKind;

export type ToolDetailsOfKind<K extends ToolDetailsKind> = Static<ToolDetailsSchemasByKind[K]>;

export type ExcelToolDetails = { [K in ToolDetailsKind]: ToolDetailsOfKind<K> }[ToolDetailsKind];

const TOOL_DETAILS_KINDS = new Set<string>(Object.keys(toolDetailsSchemasByKind));

function isToolDetailsKind(kind: string): kind is ToolDetailsKind {
  return TOOL_DETAILS_KINDS.has(kind);
}

const kindedDetailsSchema = Type.Object({ kind: Type.String() });

/**
 * Decode a tool result's `details` at the UI boundary.
 *
 * Returns `undefined` for details with no `kind`, an unknown `kind`, or a
 * payload that does not match its kind's schema (for example a session
 * persisted by an older build); callers then fall back to generic rendering.
 * Extra properties such as `outputTruncation` are allowed and preserved.
 */
export function decodeToolDetails(raw: unknown): ExcelToolDetails | undefined {
  if (!Value.Check(kindedDetailsSchema, raw) || !isToolDetailsKind(raw.kind)) return undefined;
  return decodeToolDetailsOfKind(raw, raw.kind);
}

/** Decode `details` when the caller already knows which kind it expects. */
export function decodeToolDetailsOfKind<K extends ToolDetailsKind>(
  raw: unknown,
  kind: K,
): ToolDetailsOfKind<K> | undefined {
  const schema: ToolDetailsSchemasByKind[K] = toolDetailsSchemasByKind[kind];
  return Value.Check(schema, raw) ? raw : undefined;
}

export function getToolOutputTruncationDetails(raw: unknown): ToolOutputTruncationDetails | undefined {
  return Value.Check(truncatedToolDetailsSchema, raw) ? raw.outputTruncation : undefined;
}

export type BridgeGateErrorDetails =
  | (TmuxBridgeDetails & { ok: false; gateReason: BridgeGateReason; skillHint: string })
  | (PythonBridgeDetails & { ok: false; gateReason: BridgeGateReason; skillHint: string })
  | (LibreOfficeBridgeDetails & { ok: false; gateReason: BridgeGateReason; skillHint: string })
  | (PythonTransformRangeDetails & {
    blocked: false;
    gateReason: BridgeGateReason;
    skillHint: string;
    error: string;
  });

/** A bridge tool result that failed before reaching the bridge, with a setup hint for the user. */
export function isBridgeGateError(details: ExcelToolDetails): details is BridgeGateErrorDetails {
  switch (details.kind) {
    case "tmux_bridge":
    case "python_bridge":
    case "libreoffice_bridge":
      return details.ok === false
        && details.gateReason !== undefined
        && details.skillHint !== undefined;
    case "python_transform_range":
      return details.blocked === false
        && details.error !== undefined
        && details.gateReason !== undefined
        && details.skillHint !== undefined;
    default:
      return false;
  }
}
