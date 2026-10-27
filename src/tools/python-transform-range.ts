function isToolsPythonTransformRangePayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * python_transform_range — Read a range, run Python transform, write results back.
 *
 * High-level helper for common workflows:
 * 1) read tabular values from Excel
 * 2) execute Python (native bridge preferred, Pyodide fallback when available);
 *    input too large for the bridge is refused before anything runs, with the
 *    number of rows that fit
 * 3) parse returned result JSON into a 2D grid (short rows are padded and reported)
 * 4) write it through the same checked path as write_cells (`mutation/range-write.ts`):
 *    merged cells, spills and protection block the write; the read-back reports
 *    cells that did not take their value and text Excel converted
 * The result includes the code's stdout and stderr.
 */

import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { Type, type Static, type TSchema } from "typebox";

import {
  excelRun,
  getRange,
  parseCell,
  parseRangeRef,
  qualifiedAddress,
} from "../excel/helpers.js";
import { buildWorkbookCellChangeSummary, type WorkbookCellChangeSummary } from "../audit/cell-diff.js";
import { getWorkbookChangeAuditLog } from "../audit/workbook-change-audit.js";
import { getWorkbookRecoveryLog, MAX_RECOVERY_CELLS } from "../workbook/recovery-log.js";
import {
  callDefaultPythonBridge,
  formatPythonFailure,
  getDefaultPythonBridgeConfig,
  PYODIDE_BACKEND_LABEL,
  PYTHON_BRIDGE_LIMITS,
  shouldFallbackToPyodideAfterBridgeError,
  type PythonBridgeConfig,
  type PythonBridgeRequest,
  type PythonBridgeResponse,
} from "./python-run.js";
import { probeBridgeHealth } from "./bridge-service-utils.js";
import { getErrorMessage } from "../utils/errors.js";
import { parseJson, type JsonValue } from "../utils/json.js";
import { CHECKPOINT_SKIPPED_REASON } from "./recovery-metadata.js";
import type { PythonTransformRangeDetails } from "./tool-details.js";
import { describeExcelFailure, parseExcelError } from "./mutation/excel-failure.js";
import { finalizeMutationOperation } from "./mutation/finalize.js";
import type { MutationFinalizeDependencies } from "./mutation/types.js";
import { appendMutationResultNote } from "./mutation/result-note.js";
import {
  backupUnavailable,
  blockAddress,
  chunkRowsForWrite,
  classifySentValue,
  describeWriteObstacle,
  formatVerificationNotes,
  formatInterruptedWrite,
  formatReplacedFormulas,
  formatStoppedAfterWrite,
  formatVerificationProblems,
  inspectWriteTarget,
  replacedFormulas,
  sampledChanges,
  toCellGrid,
  verifyWrittenRange,
  type CellScalar,
  type SentGrid,
  type WriteBlock,
  type WriteInterruption,
  type WriteTargetInspection,
  type WriteVerification,
} from "./mutation/range-write.js";
import { runMutation } from "./mutation/run-mutation.js";
import { STOPPED_AFTER_CHANGE_NOTE, type StopReporting } from "./with-workbook-coordinator.js";

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
      "Python code to execute. Use `input_data` (range, raw values, displayed text) and assign JSON-serializable output to `result`. Dates remain Excel serials in values; use text for their displayed representation.",
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
  values: CellScalar[][];
  text: string[][];
}

interface WriteOutputRequest {
  outputStartCell: string;
  values: CellScalar[][];
  allowOverwrite: boolean;
  /** Stop is honored between chunks, as in write_cells. */
  signal?: AbortSignal;
}

type WriteOutputResult =
  | {
    blocked: true;
    outputAddress: string;
    existingCount: number;
    /** Why the write cannot land (merge, spill, protection, too large to back up); absent when only occupied. */
    reason?: string;
  }
  | {
    blocked: false;
    outputAddress: string;
    rowsWritten: number;
    colsWritten: number;
    formulaErrorCount: number;
    /** Verification problems and notes, as write_cells reports them. */
    verificationLines?: string[];
    verification?: WriteVerification["mode"];
    notWrittenCount?: number;
    convertedCount?: number;
    changes?: WorkbookCellChangeSummary;
    beforeValues?: CellScalar[][];
    beforeFormulas?: CellScalar[][];
    backupUnavailable?: { unavailableReason: string; unavailableNote?: string };
    /** Cells that held a formula and hold a constant now. */
    formulasReplacedCount?: number;
    /**
     * Set when the write ended early (Stop, or Excel failed partway). Its text
     * comes from write_cells' wording; `rowsWritten` counts rows that landed or
     * may have landed (0 when nothing changed).
     */
    interrupted?: { text: string; stopped: boolean };
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
  /** Override Pyodide availability check (for testing). */
  isPyodideAvailable?: () => boolean;
  /** Override Pyodide runtime call (for testing). */
  callPyodide?: (
    request: PythonBridgeRequest,
    signal: AbortSignal | undefined,
  ) => Promise<PythonBridgeResponse>;
  /** Override the bridge health probe used when input is too large for the bridge (for testing). */
  probeBridge?: (url: string) => Promise<boolean>;
}

const MAX_OUTPUT_PREVIEW_CHARS = 2_000;

const mutationFinalizeDependencies: MutationFinalizeDependencies = {
  appendAuditEntry: (entry) => getWorkbookChangeAuditLog().append(entry),
};

function parseParams(raw: unknown): Params {
  if (!isToolsPythonTransformRangePayloadShape(raw)) {
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

  const timeoutMs = raw.timeout_ms;
  if (typeof timeoutMs === "number" && Number.isInteger(timeoutMs)) {
    if (timeoutMs < MIN_TIMEOUT_MS || timeoutMs > MAX_TIMEOUT_MS) {
      throw new Error(`timeout_ms must be between ${MIN_TIMEOUT_MS} and ${MAX_TIMEOUT_MS}`);
    }
    params.timeout_ms = timeoutMs;
  }

  return params;
}

function stripSheetPrefix(address: string): string {
  return address.includes("!") ? (address.split("!")[1] ?? address) : address;
}

function topLeftCellFromAddress(address: string): string {
  const local = stripSheetPrefix(address).trim();
  const first = local.split(":")[0];
  return first ?? local;
}

async function defaultReadInputRange(rangeRef: string): Promise<InputRangeSnapshot> {
  return excelRun<InputRangeSnapshot>(async (context) => {
    const { sheet, range } = getRange(context, rangeRef);
    sheet.load("name");
    range.load("address,values,text");
    await context.sync();

    return {
      sheetName: sheet.name,
      address: stripSheetPrefix(range.address),
      values: toCellGrid(range.values),
      text: range.text,
    };
  });
}

/** Write through the write_cells path: inspect the target, write in chunks, verify the read-back. */
async function defaultWriteOutputValues(request: WriteOutputRequest): Promise<WriteOutputResult> {
  const { padded, cols } = padRows(request.values);
  if (padded.length === 0 || cols === 0) {
    throw new Error("Transformed result is empty.");
  }

  const outputRef = parseRangeRef(request.outputStartCell);
  const start = parseCell(topLeftCellFromAddress(outputRef.address));
  const block: WriteBlock = { startRow: start.row, startCol: start.col, rowCount: padded.length, columnCount: cols };
  const sent: SentGrid = (r, c) => classifySentValue(padded[r]?.[c]);
  const sentFormulas = padded.some((row) => row.some((value) => typeof value === "string" && value.startsWith("=")));
  const sentEmptyCount = padded.reduce((total, row) => total + row.filter((value) => value === "").length, 0);
  // Filled in as the Excel batch progresses, so a failure can say what already happened.
  const progress: {
    outputAddress: string;
    inspection: WriteTargetInspection | null;
    writtenRows: number;
    inFlightRows: number;
  } = {
    outputAddress: outputRef.sheet ? qualifiedAddress(outputRef.sheet, blockAddress(block)) : request.outputStartCell,
    inspection: null, writtenRows: 0, inFlightRows: 0 };
  const signal = request.signal;

  const interrupted = (
    cause: WriteInterruption,
    uncertainRows: number,
    writtenRows = progress.writtenRows,
  ): Extract<WriteOutputResult, { blocked: false }> => {
    const inspection = progress.inspection;
    const before = inspection?.before ?? null;
    return {
      blocked: false,
      outputAddress: progress.outputAddress,
      rowsWritten: writtenRows + uncertainRows,
      colsWritten: block.columnCount,
      formulaErrorCount: 0,
      interrupted: {
        text: formatInterruptedWrite({
          cause,
          block,
          sheetName: inspection?.sheetName ?? null,
          fullAddr: progress.outputAddress,
          writtenRows,
          uncertainRows,
        }),
        stopped: cause.kind === "stopped",
      },
      ...(before ? { beforeValues: before.values, beforeFormulas: before.formulas } : {}),
      backupUnavailable: backupUnavailable(before !== null, block, progress.outputAddress),
    };
  };

  try {
    // runMutation: once this call fails (e.g. excelRun's timeout), the abandoned run's
    // later syncs reject, so it sends no further chunks.
    return await runMutation<WriteOutputResult>(async (context) => {
      const { sheet } = getRange(context, request.outputStartCell);
      const inspection = await inspectWriteTarget(context, sheet, block, sent);
      progress.inspection = inspection;
      const qualify = (address: string): string => qualifiedAddress(inspection.sheetName, address);
      const outputAddress = qualify(blockAddress(block));
      progress.outputAddress = outputAddress;

      const obstacle = describeWriteObstacle(inspection, qualify);
      if (obstacle) {
        return { blocked: true, outputAddress, existingCount: inspection.occupiedCount, reason: obstacle };
      }
      if (inspection.occupiedCount > 0 && !request.allowOverwrite) {
        return { blocked: true, outputAddress, existingCount: inspection.occupiedCount };
      }
      if (inspection.occupiedCount > 0 && !inspection.before) {
        const cells = block.rowCount * block.columnCount;
        return {
          blocked: true,
          outputAddress,
          existingCount: inspection.occupiedCount,
          reason: `${outputAddress} already holds data, and this write covers ${cells.toLocaleString("en-US")} cells, ` +
            `more than the ${MAX_RECOVERY_CELLS.toLocaleString("en-US")} a backup can hold, so the overwrite could not be undone. ` +
            "Transform it in smaller blocks, or ask the user to clear the old data first.",
        };
      }

      // Stop is honored between chunks, never inside one, so the result can say which rows landed.
      for (const chunk of chunkRowsForWrite(padded)) {
        if (signal?.aborted) return interrupted({ kind: "stopped" }, 0);
        sheet.getRange(blockAddress(block, chunk.rowOffset, chunk.rows.length)).values = chunk.rows;
        progress.inFlightRows = chunk.rows.length;
        await context.sync();
        progress.inFlightRows = 0;
        progress.writtenRows += chunk.rows.length;
      }

      const before = inspection.before;
      const backup = backupUnavailable(before !== null, block, outputAddress);
      if (signal?.aborted) {
        return {
          ...interrupted({ kind: "stopped" }, 0),
          interrupted: { text: formatStoppedAfterWrite(outputAddress, block), stopped: true },
        };
      }
      let verification: WriteVerification;
      try {
        verification = await verifyWrittenRange(context, sheet, block, sent, {
          calculationMode: inspection.calculationMode,
          sentFormulas,
          sentEmptyCount,
          tablesBefore: inspection.tableNames,
          sentChars: JSON.stringify(padded).length,
        });
      } catch (verifyError) {
        const failure = await describeExcelFailure(parseExcelError(verifyError), { range: outputAddress });
        return {
          blocked: false,
          outputAddress,
          rowsWritten: block.rowCount,
          colsWritten: block.columnCount,
          formulaErrorCount: 0,
          verification: "failed",
          verificationLines: [
            `⚠️ The result could not be read back to verify it: ${failure.reason} Read the range before relying on it.`,
          ],
          ...(before ? { beforeValues: before.values, beforeFormulas: before.formulas } : {}),
          backupUnavailable: backup,
        };
      }

      const changes = verification.values && verification.formulas && before
        ? buildWorkbookCellChangeSummary({
          sheetName: inspection.sheetName,
          startCell: topLeftCellFromAddress(blockAddress(block)),
          beforeValues: before.values,
          beforeFormulas: before.formulas,
          afterValues: verification.values,
          afterFormulas: verification.formulas,
        })
        : sampledChanges(inspection.sheetName, block, verification, padded.reduce((total, row) =>
          total + row.filter((value) => value !== "" && value !== null && value !== undefined).length, 0));

      const lostFormulas = before && verification.formulas
        ? replacedFormulas(block, before.formulas, verification.formulas)
        : [];
      const lostFormulasLine = formatReplacedFormulas(inspection.sheetName, lostFormulas);

      return {
        blocked: false,
        outputAddress,
        rowsWritten: block.rowCount,
        colsWritten: block.columnCount,
        formulaErrorCount: verification.formulaErrorCount,
        verification: verification.mode,
        verificationLines: [
          ...formatVerificationProblems(verification, qualify),
          ...(lostFormulasLine ? [lostFormulasLine] : []),
          ...formatVerificationNotes(verification),
        ],
        ...(lostFormulas.length > 0 ? { formulasReplacedCount: lostFormulas.length } : {}),
        ...(verification.notWrittenCount > 0 ? { notWrittenCount: verification.notWrittenCount } : {}),
        ...(verification.converted.length > 0 ? { convertedCount: verification.converted.length } : {}),
        changes,
        ...(before ? { beforeValues: before.values, beforeFormulas: before.formulas } : {}),
        backupUnavailable: backup,
      };
    });
  } catch (error) {
    // Snapshot before any await: the chunk in flight when the call failed may still land.
    const writtenRows = progress.writtenRows;
    const inFlightRows = progress.inFlightRows;
    const parsed = parseExcelError(error);
    // A batch Excel refused did not apply. One it never answered (timeout, Stop,
    // or a failure that is not an Excel answer) may have.
    const uncertainRows = parsed.outcomeUnknown || parsed.code === null ? inFlightRows : 0;
    if (writtenRows + uncertainRows === 0) {
      if (!parsed.code) throw error;
      // Nothing landed. With no chunk in flight nothing was even sent, so say so rather than "unknown".
      const failure = await describeExcelFailure(parsed, { range: progress.outputAddress }, inFlightRows === 0);
      throw new Error(`Nothing was written to ${progress.outputAddress}. ${failure.reason}`);
    }
    const failure = await describeExcelFailure(parsed, { range: progress.outputAddress });
    return interrupted({ kind: "failed", reason: failure.reason }, uncertainRows, writtenRows);
  }
}

/** Pad short rows with "" (which clears those cells) to the widest row. */
function padRows(values: CellScalar[][]): { padded: CellScalar[][]; cols: number; shortRows: number[] } {
  const cols = values.reduce((max, row) => Math.max(max, row.length), 0);
  const shortRows: number[] = [];
  const padded = values.map((row, index) => {
    if (row.length < cols) shortRows.push(index);
    return Array.from({ length: cols }, (_, col) => col < row.length ? row[col] : "");
  });
  return { padded, cols, shortRows };
}

function utf8Bytes(text: string): number {
  return new TextEncoder().encode(text).length;
}

function buildBridgeRequest(
  code: string,
  sourceAddress: string,
  values: CellScalar[][],
  text: string[][],
  timeoutMs: number | undefined,
): PythonBridgeRequest {
  return {
    code,
    input_json: JSON.stringify({ range: sourceAddress, values, text }),
    ...(timeoutMs !== undefined ? { timeout_ms: timeoutMs } : {}),
  };
}

function fitsBridgeLimits(request: PythonBridgeRequest): boolean {
  return (request.input_json?.length ?? 0) <= PYTHON_BRIDGE_LIMITS.inputJsonChars
    && utf8Bytes(JSON.stringify(request)) <= PYTHON_BRIDGE_LIMITS.requestBodyBytes;
}

/** Largest number of leading source rows whose request still fits the bridge's limits. */
function rowsThatFitBridge(code: string, input: InputRangeSnapshot, sourceAddress: string, timeoutMs: number | undefined): number {
  let low = 0;
  let high = input.values.length;
  while (low < high) {
    const mid = Math.ceil((low + high) / 2);
    if (fitsBridgeLimits(buildBridgeRequest(code, sourceAddress, input.values.slice(0, mid), input.text.slice(0, mid), timeoutMs))) low = mid;
    else high = mid - 1;
  }
  return low;
}

function formatInputTooLarge(
  request: PythonBridgeRequest,
  input: InputRangeSnapshot,
  sourceAddress: string,
  fitRows: number,
): string {
  const range = parseRangeRef(input.address).address;
  const [first = range, last = first] = range.split(":");
  const startCell = parseCell(first);
  const endCell = parseCell(last);
  const columns = `${first.replace(/\d+$/u, "")}`;
  const lastColumn = `${last.replace(/\d+$/u, "")}`;
  const lines = [
    `⛔ **Input too large for the Python bridge** — nothing ran and nothing was written.`,
    "",
    `${sourceAddress} (${input.values.length.toLocaleString("en-US")} rows) serializes to ` +
      `${(request.input_json?.length ?? 0).toLocaleString("en-US")} characters of input_json ` +
      `(${utf8Bytes(JSON.stringify(request)).toLocaleString("en-US")} bytes per request). The bridge accepts at most ` +
      `${PYTHON_BRIDGE_LIMITS.inputJsonChars.toLocaleString("en-US")} characters of input_json and ` +
      `${PYTHON_BRIDGE_LIMITS.requestBodyBytes.toLocaleString("en-US")} bytes per request.`,
  ];
  if (fitRows > 0) {
    const blockEnd = Math.min(startCell.row + fitRows - 1, endCell.row);
    lines.push(
      "",
      `About ${fitRows.toLocaleString("en-US")} rows fit per call. Transform it in blocks of at most ${fitRows.toLocaleString("en-US")} rows ` +
        `(e.g. ${qualifiedAddress(input.sheetName, `${columns}${startCell.row}:${lastColumn}${blockEnd}`)}, then the next block, ` +
        "with output_start_cell moved down to match), or read fewer columns.",
    );
  } else {
    lines.push("", "Not even one row fits; shorten the code or read fewer columns.");
  }
  return lines.join("\n");
}

function withOutputLimitGuidance(message: string, input: InputRangeSnapshot): string {
  if (!/output exceeded/iu.test(message)) return message;
  const rows = Math.max(1, input.values.length);
  const avgRowBytes = Math.max(1, Math.ceil(utf8Bytes(JSON.stringify(input.values)) / rows));
  const estimate = Math.floor(PYTHON_BRIDGE_LIMITS.outputBytes / avgRowBytes);
  return `${message}\n\nThe result JSON plus anything printed must stay under ${PYTHON_BRIDGE_LIMITS.outputBytes.toLocaleString("en-US")} bytes. ` +
    `Rows like this input's (~${avgRowBytes.toLocaleString("en-US")} bytes each) fit about ${estimate.toLocaleString("en-US")} per call: ` +
    "transform the source in blocks and write each with its own output_start_cell, and avoid printing large output.";
}

function formatOutputSections(response: PythonBridgeResponse): string[] {
  const sections: string[] = [];
  for (const [label, output] of [["Stdout", response.stdout], ["Stderr", response.stderr]] as const) {
    const trimmed = output?.trim() ?? "";
    if (trimmed.length === 0) continue;
    const shown = trimmed.length > MAX_OUTPUT_PREVIEW_CHARS
      ? `${trimmed.slice(0, MAX_OUTPUT_PREVIEW_CHARS)}\n… (${(trimmed.length - MAX_OUTPUT_PREVIEW_CHARS).toLocaleString("en-US")} more characters)`
      : trimmed;
    sections.push(`${label}:\n\n\`\`\`\n${shown}\n\`\`\``);
  }
  return sections;
}

function formatPaddingWarning(shortRows: number[], totalRows: number, cols: number): string | null {
  if (shortRows.length === 0) return null;
  const listed = shortRows.slice(0, 8).map((index) => index + 1).join(", ");
  const more = shortRows.length > 8 ? ` and ${shortRows.length - 8} more` : "";
  return `⚠️ Python returned rows of different lengths: ${shortRows.length} of ${totalRows} row(s) had fewer than ${cols} values ` +
    `(result rows ${listed}${more}). They were padded with blank cells, which clears those cells. Return a rectangular grid to avoid this.`;
}

function normalizeTo2dValues(value: JsonValue): JsonValue[][] {
  if (Array.isArray(value)) {
    return value.every((item): item is JsonValue[] => Array.isArray(item)) ? value : [value];
  }

  if (value !== null && typeof value === "object") {
    if (value.values !== undefined) return normalizeTo2dValues(value.values);
    if (value.rows !== undefined) return normalizeTo2dValues(value.rows);
  }

  return [[value]];
}

function parseBridgeResultJson(resultJson: string | undefined): CellScalar[][] {
  const trimmed = resultJson?.trim();
  if (!trimmed) {
    throw new Error(
      "Python bridge returned no result_json. Set `result = ...` in your Python code and return a 2D array (or object with values/rows).",
    );
  }

  let parsed: JsonValue;
  try {
    parsed = parseJson(trimmed);
  } catch {
    throw new Error("Python bridge returned invalid result_json payload.");
  }

  return toCellGrid(normalizeTo2dValues(parsed));
}

function formatBlockedMessage(
  inputAddress: string,
  output: Extract<WriteOutputResult, { blocked: true }>,
): string {
  if (output.reason) {
    return [
      `⛔ **Transform blocked** — nothing was written to **${output.outputAddress}**.`,
      "",
      `Source: **${inputAddress}**`,
      "",
      output.reason,
    ].join("\n");
  }
  return [
    `⛔ **Transform blocked** — destination **${output.outputAddress}** contains ${output.existingCount} non-empty cell(s).`,
    "",
    `Source: **${inputAddress}**`,
    "",
    "If the user asked to change these cells, retry with `allow_overwrite: true`. " +
      "Otherwise show the user the existing data and ask before overwriting.",
  ].join("\n");
}

function formatSuccessMessage(
  inputAddress: string,
  output: Extract<WriteOutputResult, { blocked: false }>,
  paddingWarning: string | null,
): string {
  const lines: string[] = [];

  lines.push(output.notWrittenCount
    ? `⚠️ **Partly written** — transformed **${inputAddress}**, but ${output.notWrittenCount} of ` +
      `${output.rowsWritten * output.colsWritten} cell(s) in **${output.outputAddress}** do not hold the value sent.`
    : `✅ Transformed **${inputAddress}** and wrote results to **${output.outputAddress}**.`);
  lines.push(`Rows × Cols written: ${output.rowsWritten} × ${output.colsWritten}.`);

  if (paddingWarning) lines.push(paddingWarning);

  if (output.verificationLines && output.verificationLines.length > 0) {
    lines.push(output.verificationLines.join("\n"));
  } else if (output.formulaErrorCount > 0) {
    lines.push(`⚠️ ${output.formulaErrorCount} formula error(s) detected in written output.`);
  }

  lines.push("Python input payload shape: `{ range, values, text }` (available as `input_data`): values are raw (dates are Excel serials); text is displayed cell text.");

  return lines.join("\n\n");
}

function withSkillHintLine(message: string, skillName: string): string {
  return `${message}\nSkill: ${skillName}`;
}

function shouldAttachPythonBridgeSkillHint(message: string): boolean {
  const normalized = message.toLowerCase();

  return normalized.includes("python bridge")
    || normalized.includes("python-bridge-url")
    || normalized.includes("bridge url")
    || normalized.includes("no_python_runtime")
    || normalized.includes("webassembly workers")
    || normalized.includes("bridge unavailable")
    || normalized.includes("bridge request")
    || normalized.includes("failed to fetch")
    || normalized.includes("fetch failed")
    || normalized.includes("network request failed")
    || normalized.includes("econnrefused");
}

/** Result, audit and backup for a write that ended early, worded like write_cells. */
async function finishInterrupted(
  writeResult: Extract<WriteOutputResult, { blocked: false }>,
  interrupted: { text: string; stopped: boolean },
  context: { toolCallId: string; sourceAddress: string; bridgeUrl: string | undefined; withOutput: (text: string) => string },
): Promise<AgentToolResult<PythonTransformRangeDetails>> {
  const { toolCallId, sourceAddress } = context;
  const changedCount = writeResult.rowsWritten * writeResult.colsWritten;
  const result: AgentToolResult<PythonTransformRangeDetails> = {
    content: [{ type: "text", text: context.withOutput(interrupted.text) }],
    details: {
      kind: "python_transform_range",
      blocked: false,
      inputAddress: sourceAddress,
      outputAddress: writeResult.outputAddress,
      ...(context.bridgeUrl !== undefined ? { bridgeUrl: context.bridgeUrl } : {}),
      rowsWritten: writeResult.rowsWritten,
      colsWritten: writeResult.colsWritten,
      verification: "failed",
      ...(interrupted.stopped ? { stopped: true } : {}),
    },
    isError: true,
  };

  const summary = interrupted.stopped
    ? (changedCount === 0 ? "stopped before writing" : `stopped after ${writeResult.rowsWritten} row(s)`)
    : `interrupted after ${writeResult.rowsWritten} row(s)`;

  if (changedCount === 0) {
    await finalizeMutationOperation(mutationFinalizeDependencies, {
      auditEntry: {
        toolName: "python_transform_range", toolCallId, blocked: true, inputAddress: sourceAddress,
        outputAddress: writeResult.outputAddress, changedCount: 0, changes: [], summary,
      },
    });
    return result;
  }

  const { beforeValues, beforeFormulas } = writeResult;
  await finalizeMutationOperation(mutationFinalizeDependencies, {
    auditEntry: {
      toolName: "python_transform_range", toolCallId, blocked: false, inputAddress: sourceAddress,
      outputAddress: writeResult.outputAddress, changedCount, changes: [], summary,
    },
    recovery: {
      result,
      appendRecoverySnapshot: () => beforeValues && beforeFormulas
        ? getWorkbookRecoveryLog().append({
          toolName: "python_transform_range",
          toolCallId,
          address: writeResult.outputAddress,
          changedCount,
          beforeValues,
          beforeFormulas,
        })
        : Promise.resolve(null),
      appendResultNote: appendMutationResultNote,
      ...(writeResult.backupUnavailable ?? { unavailableReason: CHECKPOINT_SKIPPED_REASON }),
    },
  });
  return result;
}

export function createPythonTransformRangeTool(
  dependencies: PythonTransformRangeToolDependencies = {},
): AgentTool<TSchema, PythonTransformRangeDetails> & StopReporting {
  const getBridgeConfig = dependencies.getBridgeConfig ?? getDefaultPythonBridgeConfig;
  const callBridge = dependencies.callBridge ?? callDefaultPythonBridge;
  const readInputRange = dependencies.readInputRange ?? defaultReadInputRange;
  const writeOutputValues = dependencies.writeOutputValues ?? defaultWriteOutputValues;
  const checkPyodide = dependencies.isPyodideAvailable;
  const pyodideCall = dependencies.callPyodide;
  const probeBridge = dependencies.probeBridge ?? probeBridgeHealth;

  return {
    name: "python_transform_range",
    reportsStop: true,
    label: "Python Transform Range",
    description:
      "Read an Excel range, run Python transformation on `{ range, values, text }` (available as `input_data`; values are raw, text is displayed, dates remain serials in values), " +
      "and write the result grid back. " +
      "Runs in-browser via Pyodide by default; uses native Python bridge when configured. " +
      "The result shows the code's stdout/stderr and, like write_cells, cells that did not take their value " +
      "(merged cells, spills) and text Excel converted.",
    parameters: schema,
    execute: async (
      toolCallId: string,
      rawParams: unknown,
      signal: AbortSignal | undefined,
    ): Promise<AgentToolResult<PythonTransformRangeDetails>> => {
      let writeStarted = false;
      let stoppedTarget: string | undefined;
      try {
        const params = parseParams(rawParams);

        const input = await readInputRange(params.range);
        const sourceAddress = qualifiedAddress(input.sheetName, input.address);

        const requestedOutputStartCell = params.output_start_cell?.trim() || topLeftCellFromAddress(input.address);
        const outputStartCell = requestedOutputStartCell.includes("!")
          ? requestedOutputStartCell
          : qualifiedAddress(input.sheetName, requestedOutputStartCell);
        stoppedTarget = outputStartCell;

        const bridgeConfig = await getBridgeConfig();
        const bridgeRequest = buildBridgeRequest(params.code, sourceAddress, input.values, input.text, params.timeout_ms);

        let bridgeResponse: PythonBridgeResponse | null = null;
        let bridgeUrlUsed: string | undefined;
        let useBridge = bridgeConfig !== null;

        // Check the bridge's request limits before sending. The default URL is always
        // configured, so only refuse when a bridge is actually answering there;
        // otherwise Pyodide (no such limit) runs the code as before.
        if (bridgeConfig && !fitsBridgeLimits(bridgeRequest)) {
          const reachable = bridgeConfig.source !== "default" || await probeBridge(bridgeConfig.url);
          if (reachable) {
            const fitRows = rowsThatFitBridge(params.code, input, sourceAddress, params.timeout_ms);
            const message = formatInputTooLarge(bridgeRequest, input, sourceAddress, fitRows);
            return {
              content: [{ type: "text", text: message }],
              details: {
                kind: "python_transform_range",
                blocked: false,
                inputAddress: sourceAddress,
                bridgeUrl: bridgeConfig.url,
                error: message,
              },
              isError: true,
            };
          }
          useBridge = false;
        }

        if (bridgeConfig && useBridge) {
          try {
            bridgeResponse = await callBridge(bridgeRequest, bridgeConfig, signal);
            bridgeUrlUsed = bridgeConfig.url;
          } catch (error) {
            if (!shouldFallbackToPyodideAfterBridgeError(error, bridgeConfig)) {
              throw new Error(withOutputLimitGuidance(getErrorMessage(error), input));
            }
          }
        }

        if (!bridgeResponse) {
          // Fall back to Pyodide
          let pyodideAvailable: boolean;
          if (checkPyodide) {
            pyodideAvailable = checkPyodide();
          } else {
            const mod = await import("../python/pyodide-runtime.js");
            pyodideAvailable = mod.isPyodideAvailable();
          }

          if (!pyodideAvailable) {
            return {
              content: [{
                type: "text",
                text: withSkillHintLine(
                  "Python is unavailable in this environment. " +
                    "The current browser/WebView does not support WebAssembly Workers (needed for in-browser Pyodide). " +
                    "Power users can configure a native Python bridge in Settings → Experimental.",
                  "python-bridge",
                ),
              }],
              details: {
                kind: "python_transform_range",
                blocked: false,
                inputAddress: sourceAddress,
                error: "no_python_runtime",
                skillHint: "python-bridge",
              },
            };
          }

          if (pyodideCall) {
            bridgeResponse = await pyodideCall(bridgeRequest, signal);
          } else {
            const mod = await import("../python/pyodide-runtime.js");
            bridgeResponse = await mod.callPyodideRuntime(bridgeRequest, signal ?? undefined);
          }
        }

        if (!bridgeResponse.ok) {
          throw new Error(formatPythonFailure(
            bridgeResponse,
            bridgeUrlUsed !== undefined ? `local Python (bridge at ${bridgeUrlUsed})` : PYODIDE_BACKEND_LABEL,
          ));
        }

        const outputSections = formatOutputSections(bridgeResponse);
        const withOutput = (text: string): string =>
          [text, ...outputSections].join("\n\n");

        const transformedValues = parseBridgeResultJson(bridgeResponse.result_json);
        const { shortRows, cols } = padRows(transformedValues);
        const paddingWarning = formatPaddingWarning(shortRows, transformedValues.length, cols);

        const writeResult = await writeOutputValues({
          outputStartCell,
          values: transformedValues,
          allowOverwrite: params.allow_overwrite === true,
          ...(signal ? { signal } : {}),
        });
        writeStarted = true;

        if (writeResult.blocked) {
          const blockedResult: AgentToolResult<PythonTransformRangeDetails> = {
            content: [{
              type: "text",
              text: withOutput(formatBlockedMessage(sourceAddress, writeResult)),
            }],
            details: {
              kind: "python_transform_range",
              blocked: true,
              inputAddress: sourceAddress,
              outputAddress: writeResult.outputAddress,
              ...(bridgeUrlUsed !== undefined ? { bridgeUrl: bridgeUrlUsed } : {}),
              existingCount: writeResult.existingCount,
            },
            isError: true,
          };

          await finalizeMutationOperation(mutationFinalizeDependencies, {
            auditEntry: {
              toolName: "python_transform_range",
              toolCallId,
              blocked: true,
              inputAddress: sourceAddress,
              outputAddress: writeResult.outputAddress,
              changedCount: 0,
              changes: [],
              ...(writeResult.reason ? { summary: "blocked: target cannot hold the values" } : {}),
            },
          });

          return blockedResult;
        }

        if (writeResult.interrupted) {
          return finishInterrupted(writeResult, writeResult.interrupted, {
            toolCallId,
            sourceAddress,
            bridgeUrl: bridgeUrlUsed,
            withOutput,
          });
        }

        const changes = writeResult.changes;
        const successText = changes && changes.changedCount > 0
          ? `${formatSuccessMessage(sourceAddress, writeResult, paddingWarning)}\n\nChanged cell(s): ${changes.changedCount}.`
          : formatSuccessMessage(sourceAddress, writeResult, paddingWarning);

        const successResult: AgentToolResult<PythonTransformRangeDetails> = {
          content: [{
            type: "text",
            text: withOutput(successText),
          }],
          details: {
            kind: "python_transform_range",
            blocked: false,
            inputAddress: sourceAddress,
            outputAddress: writeResult.outputAddress,
            ...(bridgeUrlUsed !== undefined ? { bridgeUrl: bridgeUrlUsed } : {}),
            rowsWritten: writeResult.rowsWritten,
            colsWritten: writeResult.colsWritten,
            formulaErrorCount: writeResult.formulaErrorCount,
            ...(changes !== undefined ? { changes } : {}),
            ...(writeResult.verification !== undefined ? { verification: writeResult.verification } : {}),
            ...(writeResult.notWrittenCount !== undefined ? { notWrittenCount: writeResult.notWrittenCount } : {}),
            ...(writeResult.convertedCount !== undefined ? { convertedCount: writeResult.convertedCount } : {}),
            ...(writeResult.formulasReplacedCount !== undefined ? { formulasReplacedCount: writeResult.formulasReplacedCount } : {}),
            ...(shortRows.length > 0 ? { paddedRowCount: shortRows.length } : {}),
          },
          ...(writeResult.notWrittenCount ? { isError: true } : {}),
        };
        // Stop pressed after verification finished: the result is complete, but say the change stands.
        if (signal?.aborted) appendMutationResultNote(successResult, STOPPED_AFTER_CHANGE_NOTE);

        const changedCount = changes?.changedCount ?? writeResult.rowsWritten * writeResult.colsWritten;
        const beforeValues = writeResult.beforeValues;
        const beforeFormulas = writeResult.beforeFormulas;

        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "python_transform_range",
            toolCallId,
            blocked: false,
            inputAddress: sourceAddress,
            outputAddress: writeResult.outputAddress,
            changedCount,
            changes: changes?.sample ?? [],
          },
          recovery: {
            result: successResult,
            appendRecoverySnapshot: () => {
              if (!beforeValues || !beforeFormulas) {
                return Promise.resolve(null);
              }

              return getWorkbookRecoveryLog().append({
                toolName: "python_transform_range",
                toolCallId,
                address: writeResult.outputAddress,
                changedCount,
                beforeValues,
                beforeFormulas,
              });
            },
            appendResultNote: appendMutationResultNote,
            ...(writeResult.backupUnavailable ?? { unavailableReason: CHECKPOINT_SKIPPED_REASON }),
          },
        });

        return successResult;
      } catch (error) {
        // Stop while Python ran (or before the write began): nothing changed in the workbook.
        if (signal?.aborted && !writeStarted) {
          const target = stoppedTarget ?? "the destination";
          return {
            content: [{
              type: "text",
              text: formatInterruptedWrite({
                cause: { kind: "stopped" },
                block: { startRow: 1, startCol: 0, rowCount: 1, columnCount: 1 },
                sheetName: null,
                fullAddr: target,
                writtenRows: 0,
                uncertainRows: 0,
              }),
            }],
            details: { kind: "python_transform_range", blocked: false, stopped: true, ...(stoppedTarget ? { outputAddress: stoppedTarget } : {}) },
            isError: true,
          };
        }

        const parsed = parseExcelError(error);
        // Excel failures get a specific reason; bridge/Python failures keep their message.
        const message = parsed.code
          ? (await describeExcelFailure(parsed)).reason
          : parsed.message;
        const skillHint = shouldAttachPythonBridgeSkillHint(message)
          ? "python-bridge"
          : undefined;

        return {
          content: [{
            type: "text",
            text: skillHint
              ? `Error: ${withSkillHintLine(message, skillHint)}`
              : `Error: ${message}`,
          }],
          details: {
            kind: "python_transform_range",
            blocked: false,
            error: message,
            ...(skillHint !== undefined ? { skillHint } : {}),
          },
          isError: true,
        };
      }
    },
  };
}
