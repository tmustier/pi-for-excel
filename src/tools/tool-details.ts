/**
 * Structured tool result metadata for the UI.
 *
 * Tools still return human-readable markdown in `content`, but also attach a
 * small stable `details` payload so the UI doesn't need to parse strings.
 */

import { isRecord } from "../utils/type-guards.js";

export interface WriteCellsDetails {
  kind: "write_cells";
  blocked: boolean;
  /** Sheet-qualified range when known, e.g. "Sheet1!A1:C3" */
  address?: string;
  existingCount?: number;
  formulaErrorCount?: number;
}

export interface FillFormulaDetails {
  kind: "fill_formula";
  blocked: boolean;
  /** Sheet-qualified range when known, e.g. "Sheet1!B2:B20" */
  address?: string;
  existingCount?: number;
  formulaErrorCount?: number;
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
  error?: string;
}

export type ExcelToolDetails =
  | WriteCellsDetails
  | FillFormulaDetails
  | FormatCellsDetails
  | TraceDependenciesDetails
  | ReadRangeCsvDetails
  | TmuxBridgeDetails
  | PythonBridgeDetails
  | LibreOfficeBridgeDetails
  | PythonTransformRangeDetails;

function isOptionalString(value: unknown): value is string | undefined {
  return value === undefined || typeof value === "string";
}

function isOptionalNumber(value: unknown): value is number | undefined {
  return value === undefined || typeof value === "number";
}

export function isWriteCellsDetails(value: unknown): value is WriteCellsDetails {
  if (!isRecord(value)) return false;

  if (value.kind !== "write_cells") return false;
  if (typeof value.blocked !== "boolean") return false;

  return (
    isOptionalString(value.address) &&
    isOptionalNumber(value.existingCount) &&
    isOptionalNumber(value.formulaErrorCount)
  );
}

export function isFillFormulaDetails(value: unknown): value is FillFormulaDetails {
  if (!isRecord(value)) return false;

  if (value.kind !== "fill_formula") return false;
  if (typeof value.blocked !== "boolean") return false;

  return (
    isOptionalString(value.address) &&
    isOptionalNumber(value.existingCount) &&
    isOptionalNumber(value.formulaErrorCount)
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
    isOptionalString(value.error)
  );
}
