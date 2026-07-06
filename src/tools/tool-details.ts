function isToolsToolDetailsPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Structured tool result metadata for the UI.
 *
 * Tools still return human-readable markdown in `content`, but also attach a
 * small stable `details` payload so the UI doesn't need to parse strings.
 */

import type { WorkbookCellChangeSummary } from "../audit/cell-diff.js";
import type { ConnectionToolErrorDetails } from "../connections/types.js";

export interface RecoveryCheckpointDetails {
  status: "checkpoint_created" | "not_available";
  snapshotId?: string;
  reason?: string;
}

export type ToolOutputTruncationStrategy = "head" | "tail";

export type ToolOutputTruncationReason = "lines" | "bytes" | null;

export interface ToolOutputTruncationDetails {
  version: 1;
  strategy: ToolOutputTruncationStrategy;
  truncated: boolean;
  truncatedBy: ToolOutputTruncationReason;
  totalLines: number;
  totalBytes: number;
  outputLines: number;
  outputBytes: number;
  maxLines: number;
  maxBytes: number;
  fullOutputWorkspacePath?: string;
}

export interface WriteCellsDetails {
  kind: "write_cells";
  blocked: boolean;
  /** Sheet-qualified range when known, e.g. "Sheet1!A1:C3" */
  address?: string;
  existingCount?: number;
  formulaErrorCount?: number;
  changes?: WorkbookCellChangeSummary;
  recovery?: RecoveryCheckpointDetails;
}

export interface FillFormulaDetails {
  kind: "fill_formula";
  blocked: boolean;
  /** Sheet-qualified range when known, e.g. "Sheet1!B2:B20" */
  address?: string;
  existingCount?: number;
  formulaErrorCount?: number;
  changes?: WorkbookCellChangeSummary;
  recovery?: RecoveryCheckpointDetails;
}

export interface FormatCellsDetails {
  kind: "format_cells";
  /** Sheet-qualified range when known. May be a multi-range string. */
  address?: string;
  warningsCount?: number;
  recovery?: RecoveryCheckpointDetails;
}

export interface ConditionalFormatDetails {
  kind: "conditional_format";
  action?: "add" | "clear";
  address?: string;
  recovery?: RecoveryCheckpointDetails;
}

export interface ModifyStructureDetails {
  kind: "modify_structure";
  action?: string;
  recovery?: RecoveryCheckpointDetails;
}

export interface CommentsDetails {
  kind: "comments";
  action?: string;
  address?: string;
  recovery?: RecoveryCheckpointDetails;
}

export interface ViewSettingsDetails {
  kind: "view_settings";
  action?: string;
  address?: string;
  recovery?: RecoveryCheckpointDetails;
}

export interface ChartPositionDetails {
  top: number;
  left: number;
  width: number;
  height: number;
}

export interface ChartListItemDetails {
  name: string;
  chartType: string;
  title: string;
  worksheet: string;
  position: ChartPositionDetails;
}

export interface ChartImageDetails {
  base64: string;
  mimeType: "image/png";
  width: number;
  height: number;
}

export interface ChartsDetails {
  kind: "charts";
  action?: string;
  name?: string;
  address?: string;
  sourceRange?: string;
  count?: number;
  charts?: ChartListItemDetails[];
  image?: ChartImageDetails;
  recovery?: RecoveryCheckpointDetails;
}

export type TraceDependenciesMode = "precedents" | "dependents";
export type TraceDependencySource = "api" | "formula_scan" | "mixed" | "none";

export interface DepNodeDetail {
  address: string;
  value: DynamicValue;
  /** Excel number format string, e.g. "0.00%", "#,##0", "$#,##0.00". */
  numberFormat?: string;
  formula?: string;
  /** Child nodes in traversal order (precedents for precedents mode, dependents for dependents mode). */
  precedents: DepNodeDetail[];
}

export interface TraceDependenciesDetails {
  kind: "trace_dependencies";
  root: DepNodeDetail;
  mode?: TraceDependenciesMode;
  maxDepth?: number;
  nodeCount?: number;
  edgeCount?: number;
  source?: TraceDependencySource;
  truncated?: boolean;
}

export interface ExplainFormulaReferenceDetail {
  address: string;
  valuePreview?: string;
  formulaPreview?: string;
}

export interface ExplainFormulaDetails {
  kind: "explain_formula";
  cell: string;
  hasFormula: boolean;
  formula?: string;
  valuePreview?: string;
  explanation: string;
  references: ExplainFormulaReferenceDetail[];
  truncated?: boolean;
}

export interface ReadRangeCsvDetails {
  kind: "read_range_csv";
  /** 0-indexed starting column (A=0, B=1, …) */
  startCol: number;
  /** 1-indexed starting row */
  startRow: number;
  /** Raw values grid from Excel */
  values: DynamicValue[][];
  /** Pre-serialized CSV string for the copy button */
  csv: string;
}

export type BridgeGateReason = "missing_bridge_url" | "invalid_bridge_url" | "bridge_unreachable";

export interface TmuxBridgeDetails {
  kind: "tmux_bridge";
  ok: boolean;
  action: string;
  bridgeUrl?: string;
  session?: string;
  sessionsCount?: number;
  outputPreview?: string;
  error?: string;
  gateReason?: BridgeGateReason;
  skillHint?: string;
}

export interface PythonBridgeDetails {
  kind: "python_bridge";
  ok: boolean;
  action: string;
  bridgeUrl?: string;
  exitCode?: number;
  stdoutPreview?: string;
  stderrPreview?: string;
  resultPreview?: string;
  truncated?: boolean;
  error?: string;
  gateReason?: BridgeGateReason;
  skillHint?: string;
}

export interface LibreOfficeBridgeDetails {
  kind: "libreoffice_bridge";
  ok: boolean;
  action: string;
  bridgeUrl?: string;
  inputPath?: string;
  targetFormat?: string;
  outputPath?: string;
  bytes?: number;
  converter?: string;
  error?: string;
  gateReason?: BridgeGateReason;
  skillHint?: string;
}

export interface PythonTransformRangeDetails {
  kind: "python_transform_range";
  blocked: boolean;
  inputAddress?: string;
  outputAddress?: string;
  bridgeUrl?: string;
  existingCount?: number;
  rowsWritten?: number;
  colsWritten?: number;
  formulaErrorCount?: number;
  changes?: WorkbookCellChangeSummary;
  recovery?: RecoveryCheckpointDetails;
  error?: string;
  gateReason?: BridgeGateReason;
  skillHint?: string;
}

export interface WorkbookHistorySnapshotSummary {
  id: string;
  at: number;
  toolName: string;
  address: string;
  changedCount: number;
  cellCount: number;
}

export interface WorkbookHistoryDetails {
  kind: "workbook_history";
  action: "list" | "restore" | "delete" | "clear";
  count?: number;
  snapshots?: WorkbookHistorySnapshotSummary[];
  snapshotId?: string;
  restoredSnapshotId?: string;
  inverseSnapshotId?: string;
  address?: string;
  changedCount?: number;
  deletedCount?: number;
  error?: string;
}

export type SkillsSourceKind = "bundled" | "external";

export interface SkillsListEntryDetails {
  name: string;
  sourceKind: SkillsSourceKind;
  location: string;
}

export interface SkillsListDetails {
  kind: "skills_list";
  count: number;
  names: string[];
  entries: SkillsListEntryDetails[];
  externalDiscoveryEnabled: boolean;
}

export interface SkillsReadDetails {
  kind: "skills_read";
  skillName: string;
  sourceKind: SkillsSourceKind;
  location: string;
  cacheHit: boolean;
  refreshed: boolean;
  sessionScoped: boolean;
  readCount?: number;
}

export interface SkillsInstallDetails {
  kind: "skills_install";
  skillName: string;
  location: string;
}

export interface SkillsUninstallDetails {
  kind: "skills_uninstall";
  skillName: string;
  removed: boolean;
}

export interface SkillsErrorDetails {
  kind: "skills_error";
  action: "read" | "install" | "uninstall";
  message: string;
  requestedName?: string;
  availableNames?: string[];
  externalDiscoveryEnabled: boolean;
}

export type SkillsToolDetails =
  | SkillsListDetails
  | SkillsReadDetails
  | SkillsInstallDetails
  | SkillsUninstallDetails
  | SkillsErrorDetails;

export interface WebSearchFallbackDetails {
  fromProvider: string;
  toProvider: string;
  reason: string;
}

export interface WebSearchDetails {
  kind: "web_search";
  ok: boolean;
  provider: string;
  query: string;
  sentQuery: string;
  recency?: string;
  siteFilters?: string[];
  maxResults: number;
  resultCount?: number;
  proxied?: boolean;
  proxyBaseUrl?: string;
  fallback?: WebSearchFallbackDetails;
  error?: string;
  /** `true` when the failure is due to the local CORS proxy being unreachable. */
  proxyDown?: boolean;
}

export interface FetchPageDetails {
  kind: "fetch_page";
  ok: boolean;
  url: string;
  title?: string;
  chars?: number;
  truncated?: boolean;
  proxied?: boolean;
  proxyBaseUrl?: string;
  contentType?: string;
  error?: string;
  /** `true` when the failure is due to the local CORS proxy being unreachable. */
  proxyDown?: boolean;
}

export interface McpGatewayDetails {
  kind: "mcp_gateway";
  ok: boolean;
  operation: string;
  server?: string;
  tool?: string;
  proxied?: boolean;
  proxyBaseUrl?: string;
  resultPreview?: string;
  error?: string;
  /** `true` when the failure is due to the local CORS proxy being unreachable. */
  proxyDown?: boolean;
}

export type FilesWorkspaceBackendKind = "native-directory" | "opfs" | "memory";

export interface FilesWorkbookTagDetails {
  workbookId: string;
  workbookLabel: string;
  taggedAt: number;
}

export type FilesSourceKind = "workspace" | "builtin-doc";

export interface FilesListItemDetails {
  path: string;
  size: number;
  mimeType: string;
  fileKind: "text" | "binary";
  modifiedAt: number;
  sourceKind?: FilesSourceKind;
  readOnly?: boolean;
  workbookTag?: FilesWorkbookTagDetails;
}

export interface FilesListDetails {
  kind: "files_list";
  backend: FilesWorkspaceBackendKind;
  count: number;
  files: FilesListItemDetails[];
}

export interface FilesReadDetails {
  kind: "files_read";
  backend: FilesWorkspaceBackendKind;
  path: string;
  mode: "text" | "base64";
  size: number;
  mimeType: string;
  fileKind: "text" | "binary";
  sourceKind?: FilesSourceKind;
  readOnly?: boolean;
  truncated: boolean;
  workbookTag?: FilesWorkbookTagDetails;
}

export interface FilesWriteDetails {
  kind: "files_write";
  backend: FilesWorkspaceBackendKind;
  path: string;
  encoding: "text" | "base64";
  chars: number;
  workbookTag?: FilesWorkbookTagDetails;
}

export interface FilesDeleteDetails {
  kind: "files_delete";
  backend: FilesWorkspaceBackendKind;
  path: string;
  workbookTag?: FilesWorkbookTagDetails;
}

export type FilesToolDetails =
  | FilesListDetails
  | FilesReadDetails
  | FilesWriteDetails
  | FilesDeleteDetails;

export type ExcelToolDetails =
  | WriteCellsDetails
  | FillFormulaDetails
  | FormatCellsDetails
  | ConditionalFormatDetails
  | ModifyStructureDetails
  | CommentsDetails
  | ViewSettingsDetails
  | ChartsDetails
  | TraceDependenciesDetails
  | ExplainFormulaDetails
  | ReadRangeCsvDetails
  | TmuxBridgeDetails
  | PythonBridgeDetails
  | LibreOfficeBridgeDetails
  | PythonTransformRangeDetails
  | WorkbookHistoryDetails
  | SkillsToolDetails
  | WebSearchDetails
  | FetchPageDetails
  | McpGatewayDetails
  | FilesToolDetails
  | ConnectionToolErrorDetails;

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

function isOptionalString(value: DynamicValue): value is string | undefined {
  return value === undefined || typeof value === "string";
}

function isWebSearchFallbackDetails(value: DynamicValue): value is WebSearchFallbackDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  return (
    typeof value.fromProvider === "string" &&
    typeof value.toProvider === "string" &&
    typeof value.reason === "string"
  );
}

function isOptionalWebSearchFallbackDetails(value: DynamicValue): value is WebSearchFallbackDetails | undefined {
  return value === undefined || isWebSearchFallbackDetails(value);
}

function isOptionalNumber(value: DynamicValue): value is number | undefined {
  return value === undefined || typeof value === "number";
}

function isOptionalBoolean(value: DynamicValue): value is boolean | undefined {
  return value === undefined || typeof value === "boolean";
}

function isBridgeGateReason(value: DynamicValue): value is BridgeGateReason {
  return value === "missing_bridge_url"
    || value === "invalid_bridge_url"
    || value === "bridge_unreachable";
}

function isOptionalBridgeGateReason(value: DynamicValue): value is BridgeGateReason | undefined {
  return value === undefined || isBridgeGateReason(value);
}

function isToolOutputTruncationStrategy(value: DynamicValue): value is ToolOutputTruncationStrategy {
  return value === "head" || value === "tail";
}

function isToolOutputTruncationReason(value: DynamicValue): value is ToolOutputTruncationReason {
  return value === "lines" || value === "bytes" || value === null;
}

export function isToolOutputTruncationDetails(value: DynamicValue): value is ToolOutputTruncationDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  return (
    value.version === 1 &&
    isToolOutputTruncationStrategy(value.strategy) &&
    typeof value.truncated === "boolean" &&
    isToolOutputTruncationReason(value.truncatedBy) &&
    typeof value.totalLines === "number" &&
    typeof value.totalBytes === "number" &&
    typeof value.outputLines === "number" &&
    typeof value.outputBytes === "number" &&
    typeof value.maxLines === "number" &&
    typeof value.maxBytes === "number" &&
    isOptionalString(value.fullOutputWorkspacePath)
  );
}

export function getToolOutputTruncationDetails(details: DynamicValue): ToolOutputTruncationDetails | undefined {
  if (!isToolsToolDetailsPayloadShape(details)) return undefined;
  const value = details.outputTruncation;
  return isToolOutputTruncationDetails(value)
    ? value
    : undefined;
}

function isOptionalTraceDependenciesMode(value: DynamicValue): value is TraceDependenciesMode | undefined {
  return value === undefined || value === "precedents" || value === "dependents";
}

function isOptionalTraceDependencySource(value: DynamicValue): value is TraceDependencySource | undefined {
  return value === undefined || value === "api" || value === "formula_scan" || value === "mixed" || value === "none";
}

function isOptionalStringArray(value: DynamicValue): value is string[] | undefined {
  return value === undefined || (Array.isArray(value) && value.every((item) => typeof item === "string"));
}

function isStringArray(value: DynamicValue): value is string[] {
  return Array.isArray(value) && value.every((item) => typeof item === "string");
}

function isSkillsSourceKind(value: DynamicValue): value is SkillsSourceKind {
  return value === "bundled" || value === "external";
}

function isSkillsListEntryDetails(value: DynamicValue): value is SkillsListEntryDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  return (
    typeof value.name === "string" &&
    isSkillsSourceKind(value.sourceKind) &&
    typeof value.location === "string"
  );
}

function isRecoveryCheckpointDetails(value: DynamicValue): value is RecoveryCheckpointDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  const status = value.status;
  if (status !== "checkpoint_created" && status !== "not_available") return false;

  return (
    isOptionalString(value.snapshotId) &&
    isOptionalString(value.reason)
  );
}

function isOptionalRecoveryCheckpointDetails(value: DynamicValue): value is RecoveryCheckpointDetails | undefined {
  return value === undefined || isRecoveryCheckpointDetails(value);
}

function isWorkbookCellChange(value: DynamicValue): value is WorkbookCellChangeSummary["sample"][number] {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  const beforeFormula = value.beforeFormula;
  const afterFormula = value.afterFormula;

  return (
    typeof value.address === "string" &&
    typeof value.beforeValue === "string" &&
    typeof value.afterValue === "string" &&
    (beforeFormula === undefined || typeof beforeFormula === "string") &&
    (afterFormula === undefined || typeof afterFormula === "string")
  );
}

function isWorkbookCellChangeSummary(value: DynamicValue): value is WorkbookCellChangeSummary {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  return (
    typeof value.changedCount === "number" &&
    typeof value.truncated === "boolean" &&
    Array.isArray(value.sample) &&
    value.sample.every((item) => isWorkbookCellChange(item))
  );
}

function isOptionalWorkbookCellChangeSummary(value: DynamicValue): value is WorkbookCellChangeSummary | undefined {
  return value === undefined || isWorkbookCellChangeSummary(value);
}

function isWorkbookHistorySnapshotSummary(value: DynamicValue): value is WorkbookHistorySnapshotSummary {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  return (
    typeof value.id === "string" &&
    typeof value.at === "number" &&
    typeof value.toolName === "string" &&
    typeof value.address === "string" &&
    typeof value.changedCount === "number" &&
    typeof value.cellCount === "number"
  );
}

function isOptionalWorkbookHistorySnapshotSummaryArray(
  value: DynamicValue,
): value is WorkbookHistorySnapshotSummary[] | undefined {
  return value === undefined || (Array.isArray(value) && value.every((item) => isWorkbookHistorySnapshotSummary(item)));
}

function isFilesWorkspaceBackendKind(value: DynamicValue): value is FilesWorkspaceBackendKind {
  return value === "native-directory" || value === "opfs" || value === "memory";
}

function isFilesSourceKind(value: DynamicValue): value is FilesSourceKind {
  return value === "workspace" || value === "builtin-doc";
}

function isOptionalFilesSourceKind(value: DynamicValue): value is FilesSourceKind | undefined {
  return value === undefined || isFilesSourceKind(value);
}

function isFilesWorkbookTagDetails(value: DynamicValue): value is FilesWorkbookTagDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  return (
    typeof value.workbookId === "string" &&
    typeof value.workbookLabel === "string" &&
    typeof value.taggedAt === "number"
  );
}

function isOptionalFilesWorkbookTagDetails(value: DynamicValue): value is FilesWorkbookTagDetails | undefined {
  return value === undefined || isFilesWorkbookTagDetails(value);
}

function isFilesListItemDetails(value: DynamicValue): value is FilesListItemDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  return (
    typeof value.path === "string" &&
    typeof value.size === "number" &&
    typeof value.mimeType === "string" &&
    (value.fileKind === "text" || value.fileKind === "binary") &&
    typeof value.modifiedAt === "number" &&
    isOptionalFilesSourceKind(value.sourceKind) &&
    isOptionalBoolean(value.readOnly) &&
    isOptionalFilesWorkbookTagDetails(value.workbookTag)
  );
}

export function isWriteCellsDetails(value: DynamicValue): value is WriteCellsDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  if (value.kind !== "write_cells") return false;
  if (typeof value.blocked !== "boolean") return false;

  return (
    isOptionalString(value.address) &&
    isOptionalNumber(value.existingCount) &&
    isOptionalNumber(value.formulaErrorCount) &&
    isOptionalWorkbookCellChangeSummary(value.changes) &&
    isOptionalRecoveryCheckpointDetails(value.recovery)
  );
}

export function isFillFormulaDetails(value: DynamicValue): value is FillFormulaDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  if (value.kind !== "fill_formula") return false;
  if (typeof value.blocked !== "boolean") return false;

  return (
    isOptionalString(value.address) &&
    isOptionalNumber(value.existingCount) &&
    isOptionalNumber(value.formulaErrorCount) &&
    isOptionalWorkbookCellChangeSummary(value.changes) &&
    isOptionalRecoveryCheckpointDetails(value.recovery)
  );
}

export function isFormatCellsDetails(value: DynamicValue): value is FormatCellsDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  if (value.kind !== "format_cells") return false;

  return (
    isOptionalString(value.address) &&
    isOptionalNumber(value.warningsCount) &&
    isOptionalRecoveryCheckpointDetails(value.recovery)
  );
}

export function isConditionalFormatDetails(value: DynamicValue): value is ConditionalFormatDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "conditional_format") return false;

  const action = value.action;
  const validAction = action === undefined || action === "add" || action === "clear";
  if (!validAction) return false;

  return (
    isOptionalString(value.address) &&
    isOptionalRecoveryCheckpointDetails(value.recovery)
  );
}

export function isModifyStructureDetails(value: DynamicValue): value is ModifyStructureDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "modify_structure") return false;

  return (
    isOptionalString(value.action) &&
    isOptionalRecoveryCheckpointDetails(value.recovery)
  );
}

export function isCommentsDetails(value: DynamicValue): value is CommentsDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "comments") return false;

  return (
    isOptionalString(value.action) &&
    isOptionalString(value.address) &&
    isOptionalRecoveryCheckpointDetails(value.recovery)
  );
}

export function isViewSettingsDetails(value: DynamicValue): value is ViewSettingsDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "view_settings") return false;

  return (
    isOptionalString(value.action) &&
    isOptionalString(value.address) &&
    isOptionalRecoveryCheckpointDetails(value.recovery)
  );
}

function isChartPositionDetails(value: DynamicValue): value is ChartPositionDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  return (
    typeof value.top === "number" &&
    typeof value.left === "number" &&
    typeof value.width === "number" &&
    typeof value.height === "number"
  );
}

function isChartListItemDetails(value: DynamicValue): value is ChartListItemDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  return (
    typeof value.name === "string" &&
    typeof value.chartType === "string" &&
    typeof value.title === "string" &&
    typeof value.worksheet === "string" &&
    isChartPositionDetails(value.position)
  );
}

function isOptionalChartListItemArray(value: DynamicValue): value is ChartListItemDetails[] | undefined {
  return value === undefined || (Array.isArray(value) && value.every((item) => isChartListItemDetails(item)));
}

function isChartImageDetails(value: DynamicValue): value is ChartImageDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  return (
    typeof value.base64 === "string" &&
    value.mimeType === "image/png" &&
    typeof value.width === "number" &&
    typeof value.height === "number"
  );
}

function isOptionalChartImageDetails(value: DynamicValue): value is ChartImageDetails | undefined {
  return value === undefined || isChartImageDetails(value);
}

export function isChartsDetails(value: DynamicValue): value is ChartsDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "charts") return false;

  return (
    isOptionalString(value.action) &&
    isOptionalString(value.name) &&
    isOptionalString(value.address) &&
    isOptionalString(value.sourceRange) &&
    isOptionalNumber(value.count) &&
    isOptionalChartListItemArray(value.charts) &&
    isOptionalChartImageDetails(value.image) &&
    isOptionalRecoveryCheckpointDetails(value.recovery)
  );
}

export function isReadRangeCsvDetails(value: DynamicValue): value is ReadRangeCsvDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "read_range_csv") return false;
  return (
    typeof value.startCol === "number" &&
    typeof value.startRow === "number" &&
    Array.isArray(value.values) &&
    typeof value.csv === "string"
  );
}

export function isTraceDependenciesDetails(value: DynamicValue): value is TraceDependenciesDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "trace_dependencies") return false;
  if (!isToolsToolDetailsPayloadShape(value.root)) return false;

  const root = value.root;
  if (!(typeof root.address === "string" && Array.isArray(root.precedents))) return false;

  return (
    isOptionalTraceDependenciesMode(value.mode) &&
    isOptionalNumber(value.maxDepth) &&
    isOptionalNumber(value.nodeCount) &&
    isOptionalNumber(value.edgeCount) &&
    isOptionalTraceDependencySource(value.source) &&
    isOptionalBoolean(value.truncated)
  );
}

function isExplainFormulaReferenceDetail(value: DynamicValue): value is ExplainFormulaReferenceDetail {
  if (!isToolsToolDetailsPayloadShape(value)) return false;

  return (
    typeof value.address === "string" &&
    isOptionalString(value.valuePreview) &&
    isOptionalString(value.formulaPreview)
  );
}

export function isExplainFormulaDetails(value: DynamicValue): value is ExplainFormulaDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "explain_formula") return false;

  return (
    typeof value.cell === "string" &&
    typeof value.hasFormula === "boolean" &&
    isOptionalString(value.formula) &&
    isOptionalString(value.valuePreview) &&
    typeof value.explanation === "string" &&
    Array.isArray(value.references) &&
    value.references.every((reference) => isExplainFormulaReferenceDetail(reference)) &&
    isOptionalBoolean(value.truncated)
  );
}

export function isTmuxBridgeDetails(value: DynamicValue): value is TmuxBridgeDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "tmux_bridge") return false;

  return (
    typeof value.ok === "boolean" &&
    typeof value.action === "string" &&
    isOptionalString(value.bridgeUrl) &&
    isOptionalString(value.session) &&
    isOptionalNumber(value.sessionsCount) &&
    isOptionalString(value.outputPreview) &&
    isOptionalString(value.error) &&
    isOptionalBridgeGateReason(value.gateReason) &&
    isOptionalString(value.skillHint)
  );
}

export function isPythonBridgeDetails(value: DynamicValue): value is PythonBridgeDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "python_bridge") return false;

  const truncated = value.truncated;

  return (
    typeof value.ok === "boolean" &&
    typeof value.action === "string" &&
    isOptionalString(value.bridgeUrl) &&
    isOptionalNumber(value.exitCode) &&
    isOptionalString(value.stdoutPreview) &&
    isOptionalString(value.stderrPreview) &&
    isOptionalString(value.resultPreview) &&
    (truncated === undefined || typeof truncated === "boolean") &&
    isOptionalString(value.error) &&
    isOptionalBridgeGateReason(value.gateReason) &&
    isOptionalString(value.skillHint)
  );
}

export function isLibreOfficeBridgeDetails(value: DynamicValue): value is LibreOfficeBridgeDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "libreoffice_bridge") return false;

  return (
    typeof value.ok === "boolean" &&
    typeof value.action === "string" &&
    isOptionalString(value.bridgeUrl) &&
    isOptionalString(value.inputPath) &&
    isOptionalString(value.targetFormat) &&
    isOptionalString(value.outputPath) &&
    isOptionalNumber(value.bytes) &&
    isOptionalString(value.converter) &&
    isOptionalString(value.error) &&
    isOptionalBridgeGateReason(value.gateReason) &&
    isOptionalString(value.skillHint)
  );
}

export function isBridgeGateError(value: DynamicValue): value is BridgeGateErrorDetails {
  if (isTmuxBridgeDetails(value)) {
    return value.ok === false
      && isBridgeGateReason(value.gateReason)
      && typeof value.skillHint === "string";
  }

  if (isPythonBridgeDetails(value)) {
    return value.ok === false
      && isBridgeGateReason(value.gateReason)
      && typeof value.skillHint === "string";
  }

  if (isLibreOfficeBridgeDetails(value)) {
    return value.ok === false
      && isBridgeGateReason(value.gateReason)
      && typeof value.skillHint === "string";
  }

  if (isPythonTransformRangeDetails(value)) {
    return value.blocked === false
      && typeof value.error === "string"
      && isBridgeGateReason(value.gateReason)
      && typeof value.skillHint === "string";
  }

  return false;
}

export function isPythonTransformRangeDetails(value: DynamicValue): value is PythonTransformRangeDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "python_transform_range") return false;

  return (
    typeof value.blocked === "boolean" &&
    isOptionalString(value.inputAddress) &&
    isOptionalString(value.outputAddress) &&
    isOptionalString(value.bridgeUrl) &&
    isOptionalNumber(value.existingCount) &&
    isOptionalNumber(value.rowsWritten) &&
    isOptionalNumber(value.colsWritten) &&
    isOptionalNumber(value.formulaErrorCount) &&
    isOptionalWorkbookCellChangeSummary(value.changes) &&
    isOptionalRecoveryCheckpointDetails(value.recovery) &&
    isOptionalString(value.error) &&
    isOptionalBridgeGateReason(value.gateReason) &&
    isOptionalString(value.skillHint)
  );
}

export function isWorkbookHistoryDetails(value: DynamicValue): value is WorkbookHistoryDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "workbook_history") return false;

  const action = value.action;
  const validAction = action === "list" || action === "restore" || action === "delete" || action === "clear";
  if (!validAction) return false;

  return (
    isOptionalNumber(value.count) &&
    isOptionalWorkbookHistorySnapshotSummaryArray(value.snapshots) &&
    isOptionalString(value.snapshotId) &&
    isOptionalString(value.restoredSnapshotId) &&
    isOptionalString(value.inverseSnapshotId) &&
    isOptionalString(value.address) &&
    isOptionalNumber(value.changedCount) &&
    isOptionalNumber(value.deletedCount) &&
    isOptionalString(value.error)
  );
}

export function isSkillsListDetails(value: DynamicValue): value is SkillsListDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "skills_list") return false;

  return (
    typeof value.count === "number" &&
    isStringArray(value.names) &&
    Array.isArray(value.entries) &&
    value.entries.every((entry) => isSkillsListEntryDetails(entry)) &&
    typeof value.externalDiscoveryEnabled === "boolean"
  );
}

export function isSkillsReadDetails(value: DynamicValue): value is SkillsReadDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "skills_read") return false;

  return (
    typeof value.skillName === "string" &&
    isSkillsSourceKind(value.sourceKind) &&
    typeof value.location === "string" &&
    typeof value.cacheHit === "boolean" &&
    typeof value.refreshed === "boolean" &&
    typeof value.sessionScoped === "boolean" &&
    isOptionalNumber(value.readCount)
  );
}

export function isSkillsInstallDetails(value: DynamicValue): value is SkillsInstallDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "skills_install") return false;

  return (
    typeof value.skillName === "string" &&
    typeof value.location === "string"
  );
}

export function isSkillsUninstallDetails(value: DynamicValue): value is SkillsUninstallDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "skills_uninstall") return false;

  return (
    typeof value.skillName === "string" &&
    typeof value.removed === "boolean"
  );
}

export function isSkillsErrorDetails(value: DynamicValue): value is SkillsErrorDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "skills_error") return false;

  const action = value.action;
  if (action !== "read" && action !== "install" && action !== "uninstall") return false;

  return (
    typeof value.message === "string" &&
    isOptionalString(value.requestedName) &&
    isOptionalStringArray(value.availableNames) &&
    typeof value.externalDiscoveryEnabled === "boolean"
  );
}

export function isWebSearchDetails(value: DynamicValue): value is WebSearchDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "web_search") return false;

  return (
    typeof value.ok === "boolean" &&
    typeof value.provider === "string" &&
    typeof value.query === "string" &&
    typeof value.sentQuery === "string" &&
    isOptionalString(value.recency) &&
    isOptionalStringArray(value.siteFilters) &&
    typeof value.maxResults === "number" &&
    isOptionalNumber(value.resultCount) &&
    isOptionalBoolean(value.proxied) &&
    isOptionalString(value.proxyBaseUrl) &&
    isOptionalWebSearchFallbackDetails(value.fallback) &&
    isOptionalString(value.error) &&
    isOptionalBoolean(value.proxyDown)
  );
}

export function isFetchPageDetails(value: DynamicValue): value is FetchPageDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "fetch_page") return false;

  return (
    typeof value.ok === "boolean" &&
    typeof value.url === "string" &&
    isOptionalString(value.title) &&
    isOptionalNumber(value.chars) &&
    isOptionalBoolean(value.truncated) &&
    isOptionalBoolean(value.proxied) &&
    isOptionalString(value.proxyBaseUrl) &&
    isOptionalString(value.contentType) &&
    isOptionalString(value.error) &&
    isOptionalBoolean(value.proxyDown)
  );
}

export function isMcpGatewayDetails(value: DynamicValue): value is McpGatewayDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "mcp_gateway") return false;

  return (
    typeof value.ok === "boolean" &&
    typeof value.operation === "string" &&
    isOptionalString(value.server) &&
    isOptionalString(value.tool) &&
    isOptionalBoolean(value.proxied) &&
    isOptionalString(value.proxyBaseUrl) &&
    isOptionalString(value.resultPreview) &&
    isOptionalString(value.error) &&
    isOptionalBoolean(value.proxyDown)
  );
}

function isConnectionToolErrorCode(value: DynamicValue): value is ConnectionToolErrorDetails["errorCode"] {
  return value === "missing_connection"
    || value === "invalid_connection"
    || value === "connection_auth_failed";
}

export function isConnectionToolErrorDetails(value: DynamicValue): value is ConnectionToolErrorDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "connection_error") return false;

  const reason = value.reason;

  return (
    value.ok === false &&
    isConnectionToolErrorCode(value.errorCode) &&
    typeof value.connectionId === "string" &&
    typeof value.connectionTitle === "string" &&
    (value.status === "connected" || value.status === "missing" || value.status === "invalid" || value.status === "error") &&
    typeof value.setupHint === "string" &&
    (reason === undefined || typeof reason === "string")
  );
}

export function isFilesListDetails(value: DynamicValue): value is FilesListDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "files_list") return false;

  return (
    isFilesWorkspaceBackendKind(value.backend) &&
    typeof value.count === "number" &&
    Array.isArray(value.files) &&
    value.files.every((item) => isFilesListItemDetails(item))
  );
}

export function isFilesReadDetails(value: DynamicValue): value is FilesReadDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "files_read") return false;

  return (
    isFilesWorkspaceBackendKind(value.backend) &&
    typeof value.path === "string" &&
    (value.mode === "text" || value.mode === "base64") &&
    typeof value.size === "number" &&
    typeof value.mimeType === "string" &&
    (value.fileKind === "text" || value.fileKind === "binary") &&
    isOptionalFilesSourceKind(value.sourceKind) &&
    isOptionalBoolean(value.readOnly) &&
    typeof value.truncated === "boolean" &&
    isOptionalFilesWorkbookTagDetails(value.workbookTag)
  );
}

export function isFilesWriteDetails(value: DynamicValue): value is FilesWriteDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "files_write") return false;

  return (
    isFilesWorkspaceBackendKind(value.backend) &&
    typeof value.path === "string" &&
    (value.encoding === "text" || value.encoding === "base64") &&
    typeof value.chars === "number" &&
    isOptionalFilesWorkbookTagDetails(value.workbookTag)
  );
}

export function isFilesDeleteDetails(value: DynamicValue): value is FilesDeleteDetails {
  if (!isToolsToolDetailsPayloadShape(value)) return false;
  if (value.kind !== "files_delete") return false;

  return (
    isFilesWorkspaceBackendKind(value.backend) &&
    typeof value.path === "string" &&
    isOptionalFilesWorkbookTagDetails(value.workbookTag)
  );
}
