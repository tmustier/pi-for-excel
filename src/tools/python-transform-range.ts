/**
 * python_transform_range — Read a range, run Python transform, write results back.
 *
 * High-level helper for common offline workflows:
 * 1) read tabular values from Excel
 * 2) send them to the local Python bridge as JSON
 * 3) parse returned result JSON into a 2D grid
 * 4) write transformed data back into Excel
 */

import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static, type TSchema } from "@sinclair/typebox";

import {
  computeRangeAddress,
  excelRun,
  getRange,
  padValues,
  parseRangeRef,
  qualifiedAddress,
} from "../excel/helpers.js";
import { buildWorkbookCellChangeSummary } from "../audit/cell-diff.js";
import { getWorkbookChangeAuditLog } from "../audit/workbook-change-audit.js";
import { getErrorMessage } from "../utils/errors.js";
import { findErrors } from "../utils/format.js";
import { isRecord } from "../utils/type-guards.js";
import {
  callDefaultPythonBridge,
  getDefaultPythonBridgeConfig,
  type PythonBridgeConfig,
  type PythonBridgeRequest,
  type PythonBridgeResponse,
} from "./python-run.js";
import { countOccupiedCells } from "./write-cells.js";
import type { PythonTransformRangeDetails } from "./tool-details.js";

const MIN_TIMEOUT_MS = 100;
const MAX_TIMEOUT_MS = 120_000;

const schema = Type.Object({
  range: Type.String({
    description:
      "Source range to read, e.g. \"A1:C20\" or \"Sheet2!B3:F200\".",
  }),
  code: Type.String({
    minLength: 1,
    maxLength: 40_000,
    description:
      "Python code to execute. Use `input_data` (range + values) and assign JSON-serializable output to `result`.",
  }),
  output_start_cell: Type.Optional(Type.String({
    description:
      "Top-left destination cell for transformed output. Defaults to source range top-left on the same sheet.",
  })),
  allow_overwrite: Type.Optional(Type.Boolean({
    description:
      "If false (default), write is blocked when destination contains existing values/formulas.",
  })),
  timeout_ms: Type.Optional(Type.Integer({
    minimum: MIN_TIMEOUT_MS,
    maximum: MAX_TIMEOUT_MS,
    description: "Optional Python bridge timeout in ms.",
  })),
});

type Params = Static<typeof schema>;

interface InputRangeSnapshot {
  sheetName: string;
  /** Sheet-local address, e.g. A1:C10 */
  address: string;
  values: unknown[][];
}

interface WriteOutputRequest {
  outputStartCell: string;
  values: unknown[][];
  allowOverwrite: boolean;
}

type WriteOutputResult =
  | {
    blocked: true;
    outputAddress: string;
    existingCount: number;
  }
  | {
    blocked: false;
    outputAddress: string;
    rowsWritten: number;
    colsWritten: number;
    formulaErrorCount: number;
    beforeValues?: unknown[][];
    beforeFormulas?: unknown[][];
    readBackValues?: unknown[][];
    readBackFormulas?: unknown[][];
    outputStartCell?: string;
    outputSheetName?: string;
  };

export interface PythonTransformRangeToolDependencies {
  getBridgeConfig?: () => Promise<PythonBridgeConfig | null>;
  callBridge?: (
    request: PythonBridgeRequest,
    config: PythonBridgeConfig,
    signal: AbortSignal | undefined,
  ) => Promise<PythonBridgeResponse>;
  readInputRange?: (rangeRef: string) => Promise<InputRangeSnapshot>;
  writeOutputValues?: (request: WriteOutputRequest) => Promise<WriteOutputResult>;
}

function cleanOptionalString(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined;

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function toOptionalInteger(value: unknown): number | undefined {
  if (typeof value !== "number") return undefined;
  if (!Number.isFinite(value)) return undefined;
  if (!Number.isInteger(value)) return undefined;
  return value;
}

function parseParams(raw: unknown): Params {
  if (!isRecord(raw) || Array.isArray(raw)) {
    throw new Error("Invalid python_transform_range params: expected an object.");
  }

  if (typeof raw.range !== "string" || raw.range.trim().length === 0) {
    throw new Error("range is required.");
  }

  if (typeof raw.code !== "string" || raw.code.trim().length === 0) {
    throw new Error("code is required.");
  }

  const params: Params = {
    range: raw.range,
    code: raw.code,
  };

  if (typeof raw.output_start_cell === "string") {
    params.output_start_cell = raw.output_start_cell;
  }

  if (typeof raw.allow_overwrite === "boolean") {
    params.allow_overwrite = raw.allow_overwrite;
  }

  const timeoutMs = toOptionalInteger(raw.timeout_ms);
  if (timeoutMs !== undefined) {
    params.timeout_ms = timeoutMs;
  }

  return params;
}

function validateParams(params: Params): void {
  if (params.timeout_ms !== undefined && (params.timeout_ms < MIN_TIMEOUT_MS || params.timeout_ms > MAX_TIMEOUT_MS)) {
    throw new Error(`timeout_ms must be between ${MIN_TIMEOUT_MS} and ${MAX_TIMEOUT_MS}`);
  }
}

function stripSheetPrefix(address: string): string {
  return address.includes("!") ? (address.split("!")[1] ?? address) : address;
}

function topLeftCellFromAddress(address: string): string {
  const local = stripSheetPrefix(address).trim();
  const first = local.split(":")[0];
  return first ?? local;
}

function toSheetQualifiedCell(sheetName: string, cell: string): string {
  if (cell.includes("!")) return cell;
  return qualifiedAddress(sheetName, cell);
}

function defaultOutputStartCell(input: InputRangeSnapshot): string {
  const startCell = topLeftCellFromAddress(input.address);
  return qualifiedAddress(input.sheetName, startCell);
}

async function defaultReadInputRange(rangeRef: string): Promise<InputRangeSnapshot> {
  return excelRun<InputRangeSnapshot>(async (context) => {
    const { sheet, range } = getRange(context, rangeRef);
    sheet.load("name");
    range.load("address,values");
    await context.sync();

    return {
      sheetName: sheet.name,
      address: stripSheetPrefix(range.address),
      values: range.values,
    };
  });
}

async function defaultWriteOutputValues(request: WriteOutputRequest): Promise<WriteOutputResult> {
  const { padded, rows, cols } = padValues(request.values);

  if (rows === 0 || cols === 0) {
    throw new Error("Transformed result is empty.");
  }

  return excelRun<WriteOutputResult>(async (context) => {
    const { sheet, range } = getRange(context, request.outputStartCell);
    sheet.load("name");
    range.load("address");
    await context.sync();

    const startCell = topLeftCellFromAddress(range.address);
    const outputAddressLocal = computeRangeAddress(startCell, rows, cols);
    const outputAddress = qualifiedAddress(sheet.name, outputAddressLocal);

    const targetRange = sheet.getRange(outputAddressLocal);

    let beforeValues: unknown[][] | undefined;
    let beforeFormulas: unknown[][] | undefined;

    if (!request.allowOverwrite) {
      // Skip pre-write snapshot reads when overwrite is explicitly allowed
      // to avoid an extra full-range workbook round-trip on large outputs.
      targetRange.load("values,formulas");
      await context.sync();

      beforeValues = targetRange.values;
      beforeFormulas = targetRange.formulas;

      const existingCount = countOccupiedCells(beforeValues, beforeFormulas);
      if (existingCount > 0) {
        return {
          blocked: true,
          outputAddress,
          existingCount,
        };
      }
    }

    targetRange.values = padded;
    await context.sync();

    const verify = sheet.getRange(outputAddressLocal);
    verify.load("address,values,formulas");
    await context.sync();

    const formulaErrors = findErrors(verify.values, topLeftCellFromAddress(verify.address));

    return {
      blocked: false,
      outputAddress: qualifiedAddress(sheet.name, stripSheetPrefix(verify.address)),
      rowsWritten: rows,
      colsWritten: cols,
      formulaErrorCount: formulaErrors.length,
      beforeValues,
      beforeFormulas,
      readBackValues: verify.values,
      readBackFormulas: verify.formulas,
      outputStartCell: startCell,
      outputSheetName: sheet.name,
    };
  });
}

function isRecordButNotArray(value: unknown): value is Record<string, unknown> {
  return isRecord(value) && !Array.isArray(value);
}

function normalizeTo2dValues(value: unknown): unknown[][] | null {
  if (Array.isArray(value)) {
    if (value.length === 0) return [];

    const rows: unknown[][] = [];
    let allRows = true;

    for (const item of value) {
      if (!Array.isArray(item)) {
        allRows = false;
        break;
      }

      const row: unknown[] = [];
      for (const cell of item) {
        row.push(cell);
      }
      rows.push(row);
    }

    if (allRows) {
      return rows;
    }

    return [value];
  }

  if (isRecordButNotArray(value)) {
    const valuesCandidate = value.values;
    if (valuesCandidate !== undefined) {
      return normalizeTo2dValues(valuesCandidate);
    }

    const rowsCandidate = value.rows;
    if (rowsCandidate !== undefined) {
      return normalizeTo2dValues(rowsCandidate);
    }

    return [[value]];
  }

  return [[value]];
}

function parseBridgeResultJson(resultJson: string | undefined): unknown[][] {
  const trimmed = resultJson?.trim();
  if (!trimmed) {
    throw new Error(
      "Python bridge returned no result_json. Set `result = ...` in your Python code and return a 2D array (or object with values/rows).",
    );
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(trimmed);
  } catch {
    throw new Error("Python bridge returned invalid result_json payload.");
  }

  const normalized = normalizeTo2dValues(parsed);
  if (!normalized) {
    throw new Error("Could not normalize Python result to a 2D array.");
  }

  return normalized;
}

function formatBlockedMessage(inputAddress: string, outputAddress: string, existingCount: number): string {
  return [
    `⛔ **Transform blocked** — destination **${outputAddress}** contains ${existingCount} non-empty cell(s).`,
    "",
    `Source: **${inputAddress}**`,
    "",
    "To overwrite, confirm with the user and retry with `allow_overwrite: true`.",
  ].join("\n");
}

function formatSuccessMessage(inputAddress: string, output: Exclude<WriteOutputResult, { blocked: true }>): string {
  const lines: string[] = [];

  lines.push(`✅ Transformed **${inputAddress}** and wrote results to **${output.outputAddress}**.`);
  lines.push(`Rows × Cols written: ${output.rowsWritten} × ${output.colsWritten}.`);

  if (output.formulaErrorCount > 0) {
    lines.push(`⚠️ ${output.formulaErrorCount} formula error(s) detected in written output.`);
  }

  lines.push("Python input payload shape: `{ range, values }` (available as `input_data`).");

  return lines.join("\n\n");
}

export function createPythonTransformRangeTool(
  dependencies: PythonTransformRangeToolDependencies = {},
): AgentTool<TSchema, PythonTransformRangeDetails> {
  const getBridgeConfig = dependencies.getBridgeConfig ?? getDefaultPythonBridgeConfig;
  const callBridge = dependencies.callBridge ?? callDefaultPythonBridge;
  const readInputRange = dependencies.readInputRange ?? defaultReadInputRange;
  const writeOutputValues = dependencies.writeOutputValues ?? defaultWriteOutputValues;

  return {
    name: "python_transform_range",
    label: "Python Transform Range",
    description:
      "Read an Excel range, run Python transformation on `{ range, values }`, " +
      "and write the result grid back to Excel.",
    parameters: schema,
    execute: async (
      toolCallId: string,
      rawParams: unknown,
      signal: AbortSignal | undefined,
    ): Promise<AgentToolResult<PythonTransformRangeDetails>> => {
      try {
        const params = parseParams(rawParams);
        validateParams(params);

        const input = await readInputRange(params.range);
        const sourceAddress = qualifiedAddress(input.sheetName, input.address);

        const outputStartCell = cleanOptionalString(params.output_start_cell)
          ? toSheetQualifiedCell(input.sheetName, cleanOptionalString(params.output_start_cell) ?? "")
          : defaultOutputStartCell(input);

        const bridgeConfig = await getBridgeConfig();
        if (!bridgeConfig) {
          return {
            content: [{
              type: "text",
              text:
                "Python bridge URL is not configured. " +
                "Run /experimental python-bridge-url https://localhost:3340 and enable /experimental on python-bridge.",
            }],
            details: {
              kind: "python_transform_range",
              blocked: false,
              inputAddress: sourceAddress,
              error: "missing_bridge_url",
            },
          };
        }

        const bridgeRequest: PythonBridgeRequest = {
          code: params.code,
          input_json: JSON.stringify({
            range: sourceAddress,
            values: input.values,
          }),
          timeout_ms: params.timeout_ms,
        };

        const bridgeResponse = await callBridge(bridgeRequest, bridgeConfig, signal);
        if (!bridgeResponse.ok) {
          throw new Error(bridgeResponse.error ?? "Python bridge rejected the request.");
        }

        const transformedValues = parseBridgeResultJson(bridgeResponse.result_json);

        const writeResult = await writeOutputValues({
          outputStartCell,
          values: transformedValues,
          allowOverwrite: params.allow_overwrite === true,
        });

        if (writeResult.blocked) {
          const blockedResult: AgentToolResult<PythonTransformRangeDetails> = {
            content: [{
              type: "text",
              text: formatBlockedMessage(sourceAddress, writeResult.outputAddress, writeResult.existingCount),
            }],
            details: {
              kind: "python_transform_range",
              blocked: true,
              inputAddress: sourceAddress,
              outputAddress: writeResult.outputAddress,
              bridgeUrl: bridgeConfig.url,
              existingCount: writeResult.existingCount,
            },
          };

          await getWorkbookChangeAuditLog().append({
            toolName: "python_transform_range",
            toolCallId,
            blocked: true,
            inputAddress: sourceAddress,
            outputAddress: writeResult.outputAddress,
            changedCount: 0,
            changes: [],
          });

          return blockedResult;
        }

        const parsedOutput = parseRangeRef(writeResult.outputAddress);
        const outputDiffStartCell = writeResult.outputStartCell ?? topLeftCellFromAddress(parsedOutput.address);
        const outputSheetName = writeResult.outputSheetName ?? parsedOutput.sheet ?? input.sheetName;

        const changes =
          writeResult.beforeValues &&
          writeResult.beforeFormulas &&
          writeResult.readBackValues &&
          writeResult.readBackFormulas
            ? buildWorkbookCellChangeSummary({
              sheetName: outputSheetName,
              startCell: outputDiffStartCell,
              beforeValues: writeResult.beforeValues,
              beforeFormulas: writeResult.beforeFormulas,
              afterValues: writeResult.readBackValues,
              afterFormulas: writeResult.readBackFormulas,
            })
            : undefined;

        const successText =
          changes && changes.changedCount > 0
            ? `${formatSuccessMessage(sourceAddress, writeResult)}\n\nChanged cell(s): ${changes.changedCount}.`
            : formatSuccessMessage(sourceAddress, writeResult);

        const successResult: AgentToolResult<PythonTransformRangeDetails> = {
          content: [{
            type: "text",
            text: successText,
          }],
          details: {
            kind: "python_transform_range",
            blocked: false,
            inputAddress: sourceAddress,
            outputAddress: writeResult.outputAddress,
            bridgeUrl: bridgeConfig.url,
            rowsWritten: writeResult.rowsWritten,
            colsWritten: writeResult.colsWritten,
            formulaErrorCount: writeResult.formulaErrorCount,
            changes,
          },
        };

        await getWorkbookChangeAuditLog().append({
          toolName: "python_transform_range",
          toolCallId,
          blocked: false,
          inputAddress: sourceAddress,
          outputAddress: writeResult.outputAddress,
          changedCount: changes?.changedCount ?? writeResult.rowsWritten * writeResult.colsWritten,
          changes: changes?.sample ?? [],
        });

        return successResult;
      } catch (error: unknown) {
        const message = getErrorMessage(error);
        return {
          content: [{ type: "text", text: `Error: ${message}` }],
          details: {
            kind: "python_transform_range",
            blocked: false,
            error: message,
          },
        };
      }
    },
  };
}
