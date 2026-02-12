/**
 * Structured tool result metadata for the UI.
 *
 * Tools still return human-readable markdown in `content`, but also attach a
 * small stable `details` payload so the UI doesn't need to parse strings.
 */

import type { WorkbookCellChangeSummary } from "../audit/cell-diff.js";
import { isRecord } from "../utils/type-guards.js";

export interface WriteCellsDetails {
  kind: "write_cells";
  blocked: boolean;
  /** Sheet-qualified range when known, e.g. "Sheet1!A1:C3" */
  address?: string;
  existingCount?: number;
  formulaErrorCount?: number;
  changes?: WorkbookCellChangeSummary;
}

export interface FillFormulaDetails {
  kind: "fill_formula";
  blocked: boolean;
  /** Sheet-qualified range when known, e.g. "Sheet1!B2:B20" */
  address?: string;
  existingCount?: number;
  formulaErrorCount?: number;
  changes?: WorkbookCellChangeSummary;
}

export interface FormatCellsDetails {
  kind: "format_cells";
  /** Sheet-qualified range when known. May be a multi-range string. */
  address?: string;
  warningsCount?: number;
}

export interface DepNodeDetail {
  address: string;
  value: unknown;
  /** Excel number format string, e.g. "0.00%", "#,##0", "$#,##0.00". */
  numberFormat?: string;
  formula?: string;
  precedents: DepNodeDetail[];
}

export interface TraceDependenciesDetails {
  kind: "trace_dependencies";
  root: DepNodeDetail;
}

export interface ReadRangeCsvDetails {
  kind: "read_range_csv";
  /** 0-indexed starting column (A=0, B=1, …) */
  startCol: number;
  /** 1-indexed starting row */
  startRow: number;
  /** Raw values grid from Excel */
  values: unknown[][];
  /** Pre-serialized CSV string for the copy button */
  csv: string;
}

export interface TmuxBridgeDetails {
  kind: "tmux_bridge";
  ok: boolean;
  action: string;
  bridgeUrl?: string;
  session?: string;
  sessionsCount?: number;
  outputPreview?: string;
  error?: string;
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
  error?: string;
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
  error?: string;
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
}

export type FilesWorkspaceBackendKind = "native-directory" | "opfs" | "memory";

export interface FilesWorkbookTagDetails {
  workbookId: string;
  workbookLabel: string;
  taggedAt: number;
}

export interface FilesListItemDetails {
  path: string;
  size: number;
  mimeType: string;
  fileKind: "text" | "binary";
  modifiedAt: number;
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
  | TraceDependenciesDetails
  | ReadRangeCsvDetails
  | TmuxBridgeDetails
  | PythonBridgeDetails
  | LibreOfficeBridgeDetails
  | PythonTransformRangeDetails
  | WebSearchDetails
  | McpGatewayDetails
  | FilesToolDetails;

function isOptionalString(value: unknown): value is string | undefined {
  return value === undefined || typeof value === "string";
}

function isOptionalNumber(value: unknown): value is number | undefined {
  return value === undefined || typeof value === "number";
}

function isOptionalBoolean(value: unknown): value is boolean | undefined {
  return value === undefined || typeof value === "boolean";
}

function isOptionalStringArray(value: unknown): value is string[] | undefined {
  return value === undefined || (Array.isArray(value) && value.every((item) => typeof item === "string"));
}

function isWorkbookCellChange(value: unknown): value is WorkbookCellChangeSummary["sample"][number] {
  if (!isRecord(value)) return false;

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

function isWorkbookCellChangeSummary(value: unknown): value is WorkbookCellChangeSummary {
  if (!isRecord(value)) return false;

  return (
    typeof value.changedCount === "number" &&
    typeof value.truncated === "boolean" &&
    Array.isArray(value.sample) &&
    value.sample.every((item) => isWorkbookCellChange(item))
  );
}

function isOptionalWorkbookCellChangeSummary(value: unknown): value is WorkbookCellChangeSummary | undefined {
  return value === undefined || isWorkbookCellChangeSummary(value);
}

function isFilesWorkspaceBackendKind(value: unknown): value is FilesWorkspaceBackendKind {
  return value === "native-directory" || value === "opfs" || value === "memory";
}

function isFilesWorkbookTagDetails(value: unknown): value is FilesWorkbookTagDetails {
  if (!isRecord(value)) return false;

  return (
    typeof value.workbookId === "string" &&
    typeof value.workbookLabel === "string" &&
    typeof value.taggedAt === "number"
  );
}

function isOptionalFilesWorkbookTagDetails(value: unknown): value is FilesWorkbookTagDetails | undefined {
  return value === undefined || isFilesWorkbookTagDetails(value);
}

function isFilesListItemDetails(value: unknown): value is FilesListItemDetails {
  if (!isRecord(value)) return false;

  return (
    typeof value.path === "string" &&
    typeof value.size === "number" &&
    typeof value.mimeType === "string" &&
    (value.fileKind === "text" || value.fileKind === "binary") &&
    typeof value.modifiedAt === "number" &&
    isOptionalFilesWorkbookTagDetails(value.workbookTag)
  );
}

export function isWriteCellsDetails(value: unknown): value is WriteCellsDetails {
  if (!isRecord(value)) return false;

  if (value.kind !== "write_cells") return false;
  if (typeof value.blocked !== "boolean") return false;

  return (
    isOptionalString(value.address) &&
    isOptionalNumber(value.existingCount) &&
    isOptionalNumber(value.formulaErrorCount) &&
    isOptionalWorkbookCellChangeSummary(value.changes)
  );
}

export function isFillFormulaDetails(value: unknown): value is FillFormulaDetails {
  if (!isRecord(value)) return false;

  if (value.kind !== "fill_formula") return false;
  if (typeof value.blocked !== "boolean") return false;

  return (
    isOptionalString(value.address) &&
    isOptionalNumber(value.existingCount) &&
    isOptionalNumber(value.formulaErrorCount) &&
    isOptionalWorkbookCellChangeSummary(value.changes)
  );
}

export function isFormatCellsDetails(value: unknown): value is FormatCellsDetails {
  if (!isRecord(value)) return false;

  if (value.kind !== "format_cells") return false;

  return (
    isOptionalString(value.address) &&
    isOptionalNumber(value.warningsCount)
  );
}

export function isReadRangeCsvDetails(value: unknown): value is ReadRangeCsvDetails {
  if (!isRecord(value)) return false;
  if (value.kind !== "read_range_csv") return false;
  return (
    typeof value.startCol === "number" &&
    typeof value.startRow === "number" &&
    Array.isArray(value.values) &&
    typeof value.csv === "string"
  );
}

export function isTraceDependenciesDetails(value: unknown): value is TraceDependenciesDetails {
  if (!isRecord(value)) return false;
  if (value.kind !== "trace_dependencies") return false;
  if (!isRecord(value.root)) return false;
  const root = value.root;
  return typeof root.address === "string" && Array.isArray(root.precedents);
}

export function isTmuxBridgeDetails(value: unknown): value is TmuxBridgeDetails {
  if (!isRecord(value)) return false;
  if (value.kind !== "tmux_bridge") return false;

  return (
    typeof value.ok === "boolean" &&
    typeof value.action === "string" &&
    isOptionalString(value.bridgeUrl) &&
    isOptionalString(value.session) &&
    isOptionalNumber(value.sessionsCount) &&
    isOptionalString(value.outputPreview) &&
    isOptionalString(value.error)
  );
}

export function isPythonBridgeDetails(value: unknown): value is PythonBridgeDetails {
  if (!isRecord(value)) return false;
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
    isOptionalString(value.error)
  );
}

export function isLibreOfficeBridgeDetails(value: unknown): value is LibreOfficeBridgeDetails {
  if (!isRecord(value)) return false;
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
    isOptionalString(value.error)
  );
}

export function isPythonTransformRangeDetails(value: unknown): value is PythonTransformRangeDetails {
  if (!isRecord(value)) return false;
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
    isOptionalString(value.error)
  );
}

export function isWebSearchDetails(value: unknown): value is WebSearchDetails {
  if (!isRecord(value)) return false;
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
    isOptionalString(value.error)
  );
}

export function isMcpGatewayDetails(value: unknown): value is McpGatewayDetails {
  if (!isRecord(value)) return false;
  if (value.kind !== "mcp_gateway") return false;

  return (
    typeof value.ok === "boolean" &&
    typeof value.operation === "string" &&
    isOptionalString(value.server) &&
    isOptionalString(value.tool) &&
    isOptionalBoolean(value.proxied) &&
    isOptionalString(value.proxyBaseUrl) &&
    isOptionalString(value.resultPreview) &&
    isOptionalString(value.error)
  );
}

export function isFilesListDetails(value: unknown): value is FilesListDetails {
  if (!isRecord(value)) return false;
  if (value.kind !== "files_list") return false;

  return (
    isFilesWorkspaceBackendKind(value.backend) &&
    typeof value.count === "number" &&
    Array.isArray(value.files) &&
    value.files.every((item) => isFilesListItemDetails(item))
  );
}

export function isFilesReadDetails(value: unknown): value is FilesReadDetails {
  if (!isRecord(value)) return false;
  if (value.kind !== "files_read") return false;

  return (
    isFilesWorkspaceBackendKind(value.backend) &&
    typeof value.path === "string" &&
    (value.mode === "text" || value.mode === "base64") &&
    typeof value.size === "number" &&
    typeof value.mimeType === "string" &&
    (value.fileKind === "text" || value.fileKind === "binary") &&
    typeof value.truncated === "boolean" &&
    isOptionalFilesWorkbookTagDetails(value.workbookTag)
  );
}

export function isFilesWriteDetails(value: unknown): value is FilesWriteDetails {
  if (!isRecord(value)) return false;
  if (value.kind !== "files_write") return false;

  return (
    isFilesWorkspaceBackendKind(value.backend) &&
    typeof value.path === "string" &&
    (value.encoding === "text" || value.encoding === "base64") &&
    typeof value.chars === "number" &&
    isOptionalFilesWorkbookTagDetails(value.workbookTag)
  );
}

export function isFilesDeleteDetails(value: unknown): value is FilesDeleteDetails {
  if (!isRecord(value)) return false;
  if (value.kind !== "files_delete") return false;

  return (
    isFilesWorkspaceBackendKind(value.backend) &&
    typeof value.path === "string" &&
    isOptionalFilesWorkbookTagDetails(value.workbookTag)
  );
}
