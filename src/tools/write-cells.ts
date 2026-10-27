/**
 * write_cells — Write values and formulas to Excel cells.
 *
 * The result is the truth about what is now in the workbook:
 * - Before writing: blocks on existing data (unless allowed), sheet protection,
 *   merged cells and spilled ranges that would silently drop or break values.
 * - Writes in chunks that stay under Excel's request payload limit.
 * - After writing: compares every cell (or, for large writes, both ends plus a
 *   whole-range scan) with what was sent, and reports cells that did not take
 *   their value, text Excel converted, formula errors, data-validation
 *   violations and manual calculation.
 * - Failures are tool errors (`isError: true`) with a specific reason.
 */

import { runMutation } from "./mutation/run-mutation.js";
import { Type, type Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import type { WriteCellsDetails } from "./tool-details.js";
import { getRange, qualifiedAddress, parseCell, colToLetter, parseRangeRef } from "../excel/helpers.js";
import { buildWorkbookCellChangeSummary } from "../audit/cell-diff.js";
import { getWorkbookChangeAuditLog } from "../audit/workbook-change-audit.js";
import {
  getWorkbookRecoveryLog,
  MAX_RECOVERY_CELLS,
  type AppendWorkbookRecoverySnapshotArgs,
  type WorkbookRecoverySnapshot,
} from "../workbook/recovery-log.js";
import { formatAsMarkdownTable } from "../utils/format.js";
import { STOPPED_AFTER_CHANGE_NOTE, type StopReporting } from "./with-workbook-coordinator.js";
import { finalizeMutationOperation } from "./mutation/finalize.js";
import type { MutationFinalizeDependencies } from "./mutation/types.js";
import { appendMutationResultNote } from "./mutation/result-note.js";
import type { ApprovalPrecheck, OverwrittenCells } from "./with-workbook-coordinator.js";

const schema = Type.Object({
  start_cell: Type.String({
    description:
      'Top-left cell to write from, e.g. "A1", "Sheet2!B3". ' +
      "If no sheet is specified, uses the active sheet.",
  }),
  values: Type.Array(Type.Array(Type.Any()), {
    description:
      "2D array of values. Each inner array is a row. " +
      'Strings starting with "=" are formulas. ' +
      'Excel converts text that looks like a number or date ("00123", "1/2", "3-4"); ' +
      "prefix it with an apostrophe (\"'00123\") to store it as text. " +
      'Example: [["Name", "Total"], ["Alice", "=SUM(B2:B10)"]]',
  }),
  allow_overwrite: Type.Optional(
    Type.Boolean({
      description:
        "Set to true to replace existing values or formulas. Default: false. " +
        "Set it when the user asked to change those cells. " +
        "If false and the target range contains values or formulas, the write is blocked " +
        "and the existing data is returned so you can ask the user.",
    }),
  ),
});

type Params = Static<typeof schema>;

interface InvalidFormula {
  address: string;
  formula: string;
  reason: string;
}

type WriteOutcome =
  | { kind: "obstacle"; inspection: WriteTargetInspection; reason: string }
  | { kind: "occupied"; inspection: WriteTargetInspection }
  | { kind: "too_large_to_back_up"; inspection: WriteTargetInspection }
  | { kind: "stopped"; inspection: WriteTargetInspection }
  | ({
    kind: "written";
    inspection: WriteTargetInspection;
  } & WriteReadBack);

interface WriteCellsToolDependencies {
  appendAuditEntry: MutationFinalizeDependencies["appendAuditEntry"];
  appendRecoverySnapshot: (
    args: AppendWorkbookRecoverySnapshotArgs,
  ) => Promise<WorkbookRecoverySnapshot | null>;
}

const defaultDependencies: WriteCellsToolDependencies = {
  appendAuditEntry: (entry) => getWorkbookChangeAuditLog().append(entry),
  appendRecoverySnapshot: (args) => getWorkbookRecoveryLog().append(args),
};

type CheckedWrite =
  | { result: AgentToolResult<WriteCellsDetails> }
  | { padded: unknown[][]; rows: number; cols: number; startCellRef: string };

/** Input checks that refuse a write before it touches the workbook. */
function checkWriteInput(params: Params): CheckedWrite {
  if (!params.values || params.values.length === 0) {
    return {
      result: {
        content: [{ type: "text", text: "Error: values array is empty." }],
        details: { kind: "write_cells", blocked: false },
      },
    };
  }

  const { padded, rows, cols } = padValues(params.values);

  const startCellRef = parseRangeRef(params.start_cell).address;

  if (startCellRef.includes(":")) {
    return {
      result: {
        content: [{ type: "text", text: "Error: start_cell must be a single cell (e.g. \"A1\")." }],
        details: { kind: "write_cells", blocked: false },
      },
    };
  }

  let invalidFormulas: InvalidFormula[] = [];
  try {
    invalidFormulas = findInvalidFormulas(padded, startCellRef);
  } catch {
    return {
      result: {
        content: [{ type: "text", text: `Error: invalid start_cell "${params.start_cell}".` }],
        details: { kind: "write_cells", blocked: false },
      },
    };
  }

  if (invalidFormulas.length > 0) {
    const lines: string[] = [];
    lines.push("⛔ **Write blocked** — invalid formula syntax detected:");
    for (const invalid of invalidFormulas) {
      lines.push(`- ${invalid.address}: ${invalid.formula} (${invalid.reason})`);
    }
    lines.push("");
    lines.push("Fix the formulas and retry.");
    return {
      result: {
        content: [{ type: "text", text: lines.join("\n") }],
        details: { kind: "write_cells", blocked: true },
      },
    };
  }

  return { padded, rows, cols, startCellRef };
}

export function createWriteCellsTool(
  dependencies: Partial<WriteCellsToolDependencies> = {},
): AgentTool<typeof schema, WriteCellsDetails> & ApprovalPrecheck<typeof schema, WriteCellsDetails> {
  const resolvedDependencies: WriteCellsToolDependencies = {
    appendAuditEntry: dependencies.appendAuditEntry ?? defaultDependencies.appendAuditEntry,
    appendRecoverySnapshot:
      dependencies.appendRecoverySnapshot ?? defaultDependencies.appendRecoverySnapshot,
  };
  const mutationFinalizeDependencies: MutationFinalizeDependencies = {
    appendAuditEntry: resolvedDependencies.appendAuditEntry,
  };

  const reportBlocked = async (
    toolCallId: string,
    blocked: BlockedWriteCellsResult,
  ): Promise<AgentToolResult<WriteCellsDetails>> => {
    const blockedResult = formatBlocked(blocked);

    await finalizeMutationOperation(mutationFinalizeDependencies, {
      auditEntry: {
        toolName: "write_cells",
        toolCallId,
        blocked: true,
        ...(blockedResult.details.address !== undefined ? { outputAddress: blockedResult.details.address } : {}),
        changedCount: 0,
        changes: [],
      },
    });

    return blockedResult;
  };

  return {
    name: "write_cells",
    label: "Write Cells",
    description:
      "Write values and formulas to Excel cells. Provide a start cell and a 2D array. " +
      'Strings starting with "=" are treated as formulas. ' +
      "By default, blocks if the target range already contains data — " +
      "set allow_overwrite=true when the user asked to change those cells, otherwise confirm with the user first. " +
      "After writing, verifies what Excel stored and reports cells that did not take their value " +
      "(merged cells, spilled ranges), converted text, formula errors and data-validation violations.",
    parameters: schema,
    // In Confirm mode, a write that would be refused (bad input, or overwrite
    // protection) is refused before the user is asked to approve it.
    checkBeforeApproval: async (toolCallId, params) => {
      try {
        const checked = checkWriteInput(params);
        if ("result" in checked) return { refused: checked.result };

        const target = await excelRun(async (context) => {
          const { sheet } = getRange(context, params.start_cell);
          sheet.load("name");
          const rangeAddr = computeRangeAddress(checked.startCellRef, checked.rows, checked.cols);
          const targetRange = sheet.getRange(rangeAddr);
          targetRange.load("values,formulas");
          await context.sync();
          return { sheetName: sheet.name, rangeAddr, values: targetRange.values, formulas: targetRange.formulas };
        });

        const occupied = describeOccupiedCells({
          sheetName: target.sheetName,
          topLeft: checked.startCellRef,
          values: target.values,
          formulas: target.formulas,
        });
        if (!occupied) return null;
        if (params.allow_overwrite) return { overwrites: occupied };
        return {
          refused: await reportBlocked(toolCallId, {
            blocked: true,
            sheetName: target.sheetName,
            address: target.rangeAddr,
            existingCount: occupied.count,
            existingValues: target.values,
          }),
        };
      } catch {
        // Let the call proceed to approval; it reports its own error when it runs.
        return null;
      }
    },
    execute: async (
      toolCallId: string,
      params: Params,
      signal?: AbortSignal,
    ): Promise<AgentToolResult<WriteCellsDetails>> => {
      if (!params.values || params.values.length === 0) {
        return errorResult("Error: values array is empty.");
      }

      const sentGrid = toCellGrid(params.values);
      const cols = Math.max(...sentGrid.map((row) => row.length));
      const rows = sentGrid.length;
      // Short rows are padded with "" (which clears those cells), as before.
      const padded: CellScalar[][] = sentGrid.map((row) =>
        Array.from({ length: cols }, (_, index) => index < row.length ? row[index] : ""));
      const startCellRef = parseRangeRef(params.start_cell).address;

      if (startCellRef.includes(":")) {
        return errorResult("Error: start_cell must be a single cell (e.g. \"A1\").");
      }

      let start: { col: number; row: number };
      try {
        const checked = checkWriteInput(params);
        if ("result" in checked) return checked.result;
        const { padded, rows, cols, startCellRef } = checked;

        const result = await excelRun<WriteCellsResult>(async (context) => {
          const { sheet } = getRange(context, params.start_cell);
          const inspected = await inspectWriteTarget(context, sheet, block, sent);
          progress.inspection = inspected;

          const obstacle = describeWriteObstacle(inspected, (address) => qualifiedAddress(inspected.sheetName, address));
          if (obstacle) return { kind: "obstacle", inspection: inspected, reason: obstacle };
          if (inspected.occupiedCount > 0 && !params.allow_overwrite) return { kind: "occupied", inspection: inspected };
          if (inspected.occupiedCount > 0 && !inspected.before) return { kind: "too_large_to_back_up", inspection: inspected };

          // Stop is honored between chunks, never inside one, so the result can
          // say exactly which rows landed.
          for (const chunk of chunkRowsForWrite(padded)) {
            if (signal?.aborted) return { kind: "stopped", inspection: inspected };
            sheet.getRange(blockAddress(block, chunk.rowOffset, chunk.rows.length)).values = chunk.rows;
            progress.inFlightRows = chunk.rows.length;
            await context.sync();
            progress.inFlightRows = 0;
            progress.writtenRows += chunk.rows.length;
          }

          if (signal?.aborted) {
            return { kind: "written", inspection: inspected, verification: null, verifyError: null, stopped: true };
          }

          try {
            const verification = await verifyWrittenRange(context, sheet, block, sent, {
              calculationMode: inspected.calculationMode,
              sentFormulas,
              sentEmptyCount,
              tablesBefore: inspected.tableNames,
              sentChars: JSON.stringify(padded).length,
            });
            return { kind: "written", inspection: inspected, verification, verifyError: null, stopped: false };
          } catch (verifyError) {
            const failure = await describeExcelFailure(parseExcelError(verifyError), { range: params.start_cell });
            return { kind: "written", inspection: inspected, verification: null, verifyError: failure.reason, stopped: false };
          }
        });
      } catch (error) {
        // Snapshot before any await: an abandoned run (runMutation) may still finish its in-flight chunk.
        const writtenRows = progress.writtenRows;
        const inFlightRows = progress.inFlightRows;
        const parsed = parseExcelError(error);
        // A batch Excel refused did not apply. One it never answered (timeout,
        // Stop, or a failure that is not an Excel answer) may have.
        const inFlightUnknown = parsed.outcomeUnknown || parsed.code === null;
        return interruptedWrite({
          cause: {
            kind: "failed",
            failure: await describeExcelFailure(
              parsed,
              { range: params.start_cell },
              writtenRows === 0 && inFlightRows === 0,
            ),
          },
          params,
          toolCallId,
          block,
          inspection: progress.inspection,
          writtenRows,
          uncertainRows: inFlightUnknown ? inFlightRows : 0,
          dependencies: resolvedDependencies,
          finalizeDependencies: mutationFinalizeDependencies,
        });
      }

        if (result.blocked) {
          return reportBlocked(toolCallId, result);
        }

        const successResult = formatSuccess(result, rows, cols);

        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "write_cells",
            toolCallId,
            blocked: false,
            ...(successResult.details.address !== undefined ? { outputAddress: successResult.details.address } : {}),
            changedCount: successResult.details.changes?.changedCount ?? 0,
            changes: successResult.details.changes?.sample ?? [],
          },
          recovery: {
            result: successResult,
            appendRecoverySnapshot: () => resolvedDependencies.appendRecoverySnapshot({
              toolName: "write_cells",
              toolCallId,
              address: successResult.details.address ?? qualifiedAddress(result.sheetName, result.address),
              changedCount: successResult.details.changes?.changedCount ?? 0,
              beforeValues: result.beforeValues,
              beforeFormulas: result.beforeFormulas,
              afterFormulas: result.readBackFormulas,
            }),
            appendResultNote: appendMutationResultNote,
            unavailableReason: CHECKPOINT_SKIPPED_REASON,
            unavailableNote: CHECKPOINT_SKIPPED_NOTE,
          },
        });
        return blockedResult;
      }

      const successResult = formatWritten(outcome, block, padded, fullAddr);
      // Stop pressed after verification finished: the result is complete, but say the change stands.
      if (signal?.aborted && !outcome.stopped) appendMutationResultNote(successResult, STOPPED_AFTER_CHANGE_NOTE);
      const backupBefore = outcome.inspection.before;

      await finalizeMutationOperation(mutationFinalizeDependencies, {
        auditEntry: {
          toolName: "write_cells",
          toolCallId,
          blocked: false,
          outputAddress: fullAddr,
          changedCount: successResult.details.changes?.changedCount ?? 0,
          changes: successResult.details.changes?.sample ?? [],
        },
        recovery: {
          result: successResult,
          appendRecoverySnapshot: () => backupBefore
            ? resolvedDependencies.appendRecoverySnapshot({
              toolName: "write_cells",
              toolCallId,
              address: fullAddr,
              changedCount: successResult.details.changes?.changedCount ?? 0,
              beforeValues: backupBefore.values,
              beforeFormulas: backupBefore.formulas,
            })
            : Promise.resolve(null),
          appendResultNote: appendMutationResultNote,
          ...backupUnavailable(backupBefore !== null, block, fullAddr),
        },
      });

      return successResult;
    },
  };
}

type InterruptCause =
  | { kind: "stopped" }
  | { kind: "failed"; failure: Awaited<ReturnType<typeof describeExcelFailure>> };

interface InterruptedWriteArgs {
  cause: InterruptCause;
  params: Params;
  toolCallId: string;
  block: WriteBlock;
  inspection: WriteTargetInspection | null;
  /** Rows whose chunk Excel confirmed. */
  writtenRows: number;
  /** Rows of the chunk in flight when Excel stopped answering: may or may not have landed. */
  uncertainRows: number;
  dependencies: WriteCellsToolDependencies;
  finalizeDependencies: MutationFinalizeDependencies;
}

/**
 * The write stopped partway (Stop, or an Excel failure): say exactly what is
 * in the workbook. Rows that landed stay, and are backed up and audited.
 */
async function interruptedWrite(args: InterruptedWriteArgs): Promise<AgentToolResult<WriteCellsDetails>> {
  const { block, inspection, writtenRows, uncertainRows, cause } = args;
  const fullAddr = inspection
    ? qualifiedAddress(inspection.sheetName, blockAddress(block))
    : args.params.start_cell;
  const reason = cause.kind === "failed" ? cause.failure.reason : "";
  const errorCode = cause.kind === "failed" && cause.failure.code ? { errorCode: cause.failure.code } : {};
  const stopped = cause.kind === "stopped" ? { stopped: true } : {};
  const text = formatInterruptedWrite({
    cause: cause.kind === "stopped" ? cause : { kind: "failed", reason },
    block,
    sheetName: inspection?.sheetName ?? null,
    fullAddr,
    writtenRows,
    uncertainRows,
  });

  if (!inspection || writtenRows + uncertainRows === 0) {
    await finalizeMutationOperation(args.finalizeDependencies, {
      auditEntry: {
        toolName: "write_cells",
        toolCallId: args.toolCallId,
        blocked: true,
        outputAddress: fullAddr,
        changedCount: 0,
        changes: [],
        summary: cause.kind === "stopped" ? "stopped before writing" : `error: ${reason}`,
      },
    });
    return errorResult(text, {
      kind: "write_cells", blocked: false, address: fullAddr, ...errorCode, ...stopped,
      ...(cause.kind === "stopped" ? { writtenCount: 0 } : {}),
    });
  }

  const touchedRows = writtenRows + uncertainRows;
  const touchedAddr = qualifiedAddress(inspection.sheetName, blockAddress(block, 0, touchedRows));

  const result = errorResult(text, {
    kind: "write_cells",
    blocked: false,
    address: fullAddr,
    ...errorCode,
    ...stopped,
    ...(cause.kind === "stopped" ? { writtenCount: writtenRows * block.columnCount } : {}),
  });
  const before = inspection.before;
  const changedCount = touchedRows * block.columnCount;
  await finalizeMutationOperation(args.finalizeDependencies, {
    auditEntry: {
      toolName: "write_cells",
      toolCallId: args.toolCallId,
      blocked: false,
      outputAddress: touchedAddr,
      changedCount,
      changes: [],
      summary: `${cause.kind === "stopped" ? "stopped" : "interrupted"} after ${writtenRows} of ${block.rowCount} rows` +
        (uncertainRows > 0 ? ` (${uncertainRows} more unconfirmed)` : "") + (reason ? `: ${reason}` : ""),
    },
    recovery: {
      result,
      appendRecoverySnapshot: () => before
        ? args.dependencies.appendRecoverySnapshot({
          toolName: "write_cells",
          toolCallId: args.toolCallId,
          address: fullAddr,
          changedCount,
          beforeValues: before.values,
          beforeFormulas: before.formulas,
        })
        : Promise.resolve(null),
      appendResultNote: appendMutationResultNote,
      ...backupUnavailable(before !== null, block, fullAddr),
    },
  });
  return result;
}

function findInvalidFormulas(values: CellScalar[][], start: { col: number; row: number }): InvalidFormula[] {
  const invalid: InvalidFormula[] = [];

  for (let r = 0; r < values.length; r++) {
    const valueRow = values[r] ?? [];
    for (let c = 0; c < valueRow.length; c++) {
      const value = valueRow[c];
      if (typeof value === "string" && value.startsWith("=")) {
        const reason = validateFormula(value);
        if (reason) {
          invalid.push({
            address: `${colToLetter(start.col + c)}${start.row + r}`,
            formula: value,
            reason,
          });
        }
      }
    }
  }

  return invalid;
}

export function validateFormula(formula: string): string | null {
  if (!formula.startsWith("=")) return null;
  const body = formula.slice(1);

  if (body.trim().length === 0) return "Empty formula";

  const quoteCount = (body.match(/"/g) || []).length;
  if (quoteCount % 2 !== 0) return "Unbalanced quotes";

  let depth = 0;
  let inString = false;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') {
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (ch === "(") depth += 1;
    if (ch === ")") {
      depth -= 1;
      if (depth < 0) return "Unbalanced parentheses";
    }
  }
  if (depth !== 0) return "Unbalanced parentheses";

  const trimmed = body.trim();
  if (/[+\-*/^&,]$/.test(trimmed)) return "Formula ends with an operator";

  return null;
}

export function countOccupiedCells(values: unknown[][], formulas: unknown[][]): number {
  let count = 0;
  for (let r = 0; r < values.length; r++) {
    const valueRow = values[r] ?? [];
    for (let c = 0; c < valueRow.length; c++) {
      const value = valueRow[c];
      const formula = formulas[r]?.[c];
      const hasValue = value !== null && value !== undefined && value !== "";
      const hasFormula = typeof formula === "string" && formula.startsWith("=");
      if (hasValue || hasFormula) count += 1;
    }
  }
  return count;
}

const OVERWRITE_SAMPLE_SIZE = 3;

/**
 * The non-empty cells of a grid starting at `topLeft` (e.g. "B2"), for an
 * approval that must say what a change replaces. Formulas are shown as
 * formulas, other cells by value.
 */
export function describeOccupiedCells(args: {
  sheetName: string;
  topLeft: string;
  values: unknown[][];
  formulas: unknown[][];
}): OverwrittenCells | null {
  const start = parseCell(args.topLeft);
  const sample: OverwrittenCells["sample"] = [];
  let count = 0;
  args.values.forEach((row, r) => {
    row.forEach((value, c) => {
      const formula = args.formulas[r]?.[c];
      const isFormula = typeof formula === "string" && formula.startsWith("=");
      if (!isFormula && (value === null || value === undefined || value === "")) return;
      count += 1;
      if (sample.length >= OVERWRITE_SAMPLE_SIZE) return;
      sample.push({
        address: qualifiedAddress(args.sheetName, `${colToLetter(start.col + c)}${start.row + r}`),
        content: isFormula ? formula : String(value),
      });
    });
  });
  return count > 0 ? { count, sample } : null;
}

const VERIFIED_VALUES_PREVIEW_ROWS = 8;
const VERIFIED_VALUES_PREVIEW_COLS = 6;

function previewLines(values: CellScalar[][], label: string, totalRows: number, totalCols: number): string[] {
  const shownRows = Math.min(values.length, VERIFIED_VALUES_PREVIEW_ROWS);
  const shownCols = Math.min(totalCols, VERIFIED_VALUES_PREVIEW_COLS);
  const preview = values.slice(0, shownRows).map((row) => row.slice(0, shownCols));
  const truncated = shownRows < totalRows || shownCols < totalCols;

  const lines: string[] = [];
  lines.push(truncated
    ? `**${label} (preview ${shownRows}×${shownCols} of ${totalRows}×${totalCols}):**`
    : `**${label}:**`);
  lines.push(formatAsMarkdownTable(preview));
  if (truncated) {
    const omissions: string[] = [];
    const omittedRows = totalRows - shownRows;
    const omittedCols = totalCols - shownCols;
    if (omittedRows > 0) omissions.push(`${omittedRows} more row${omittedRows === 1 ? "" : "s"}`);
    if (omittedCols > 0) omissions.push(`${omittedCols} more column${omittedCols === 1 ? "" : "s"}`);
    lines.push("");
    lines.push(`_Showing preview only (${omissions.join(" and ")})._`);
  }
  return lines;
}

function formatOccupied(inspection: WriteTargetInspection, fullAddr: string): AgentToolResult<WriteCellsDetails> {
  const lines: string[] = [];
  lines.push(`⛔ **Write blocked** — ${fullAddr} contains ${inspection.occupiedCount} non-empty cell(s); nothing was written.`);
  lines.push("");
  if (inspection.before) {
    lines.push("**Existing data:**");
    lines.push(formatAsMarkdownTable(inspection.before.values));
  } else if (inspection.occupiedAddress) {
    lines.push(`**Existing data in:** ${qualifiedAddress(inspection.sheetName, inspection.occupiedAddress)} (too large to show; read it with read_range).`);
  }
  lines.push("");
  lines.push(
    "If the user asked to change these cells, retry with `allow_overwrite: true`. " +
    "Otherwise show the user the existing data and ask before overwriting.",
  );

  return errorResult(lines.join("\n"), {
    kind: "write_cells",
    blocked: true,
    address: fullAddr,
    existingCount: inspection.occupiedCount,
  });
}

function formatTooLargeToBackUp(
  inspection: WriteTargetInspection,
  fullAddr: string,
  block: WriteBlock,
): AgentToolResult<WriteCellsDetails> {
  const cells = block.rowCount * block.columnCount;
  const occupied = inspection.occupiedAddress
    ? qualifiedAddress(inspection.sheetName, inspection.occupiedAddress)
    : fullAddr;
  return errorResult(
    `⛔ **Write refused** — nothing was written. ${occupied} already holds data, and this write covers ` +
    `${cells.toLocaleString("en-US")} cells, more than the ${MAX_RECOVERY_CELLS.toLocaleString("en-US")} a backup can hold, ` +
    "so the overwrite could not be undone.\n\n" +
    `Write it in blocks of at most ${MAX_RECOVERY_CELLS.toLocaleString("en-US")} cells (each gets its own backup), ` +
    "or ask the user to clear the old data first.",
    { kind: "write_cells", blocked: true, address: fullAddr, existingCount: inspection.occupiedCount },
  );
}

function formatWritten(
  outcome: Extract<WriteOutcome, { kind: "written" }>,
  block: WriteBlock,
  padded: CellScalar[][],
  fullAddr: string,
): AgentToolResult<WriteCellsDetails> {
  const { inspection } = outcome;
  const qualify = (address: string): string => qualifiedAddress(inspection.sheetName, address);
  const startCell = `${colToLetter(block.startCol)}${block.startRow}`;
  const size = `(${block.rowCount}×${block.columnCount})`;

  if (outcome.stopped) {
    return {
      content: [{
        type: "text",
        text: formatStoppedAfterWrite(fullAddr, block),
      }],
      details: {
        kind: "write_cells", blocked: false, address: fullAddr, verification: "failed", stopped: true,
        writtenCount: block.rowCount * block.columnCount,
      },
    };
  }

  if (!outcome.verification) {
    return {
      content: [{
        type: "text",
        text: `Written to **${fullAddr}** ${size}, but the result could not be read back to verify it: ${outcome.verifyError}\n\n` +
          "Read the range before relying on it.",
      }],
      details: { kind: "write_cells", blocked: false, address: fullAddr, verification: "failed" },
    };
  }

  const verification = outcome.verification;
  const lines: string[] = [];
  const notWritten = verification.notWrittenCount;
  const totalCells = block.rowCount * block.columnCount;
  if (notWritten > 0) {
    lines.push(`⚠️ **Partly written** to **${fullAddr}** ${size}: ${notWritten} of ${totalCells} cell(s) do not hold the value sent.`);
  } else {
    lines.push(`Written to **${fullAddr}** ${size}`);
  }

  const problems = formatVerificationProblems(verification, qualify);
  if (problems.length > 0) {
    lines.push("");
    lines.push(...problems);
  }

  const lostFormulas = inspection.before && verification.formulas
    ? replacedFormulas(block, inspection.before.formulas, verification.formulas)
    : [];
  const lostFormulasLine = formatReplacedFormulas(inspection.sheetName, lostFormulas);
  if (lostFormulasLine) {
    lines.push("");
    lines.push(lostFormulasLine);
  }

  const readBack = verification.values ?? verification.sampleRows[0]?.values ?? null;
  if (readBack && verification.formulaErrorCount === 0) {
    lines.push("");
    lines.push(...previewLines(
      readBack,
      hasVerificationProblems(verification) ? "Values now in the cells" : "Verified values",
      block.rowCount,
      block.columnCount,
    ));
  }

  const notes = formatVerificationNotes(verification);
  if (notes.length > 0) {
    lines.push("");
    lines.push(...notes);
  }

  const changes = verification.values && verification.formulas && inspection.before
    ? buildWorkbookCellChangeSummary({
      sheetName: inspection.sheetName,
      startCell,
      beforeValues: inspection.before.values,
      beforeFormulas: inspection.before.formulas,
      afterValues: verification.values,
      afterFormulas: verification.formulas,
    })
    : sampledChanges(inspection.sheetName, block, verification, padded.reduce((total, row) =>
      total + row.filter((value) => value !== "" && value !== null && value !== undefined).length, 0));

  if (changes.changedCount > 0) {
    lines.push("");
    lines.push(`Changed cell(s): ${changes.changedCount}.`);
  }

  const invalidCount = verification.invalidCells.length;
  const details: WriteCellsDetails = {
    kind: "write_cells",
    blocked: false,
    address: fullAddr,
    formulaErrorCount: verification.formulaErrorCount,
    changes,
    verification: verification.mode,
    ...(notWritten > 0 ? { notWrittenCount: notWritten } : {}),
    ...(verification.converted.length > 0 ? { convertedCount: verification.converted.length } : {}),
    ...(invalidCount > 0 ? { invalidCount } : {}),
    ...(lostFormulas.length > 0 ? { formulasReplacedCount: lostFormulas.length } : {}),
    ...(verification.manualCalculation ? { calculationMode: verification.manualCalculation } : {}),
  };

  return {
    content: [{ type: "text", text: lines.join("\n") }],
    details,
    ...(notWritten > 0 ? { isError: true } : {}),
  };
}
