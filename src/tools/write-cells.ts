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

import { Type, type Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import type { WriteCellsDetails } from "./tool-details.js";
import {
  excelRun, getRange, qualifiedAddress, parseCell,
  colToLetter, parseRangeRef,
} from "../excel/helpers.js";
import { buildWorkbookCellChangeSummary } from "../audit/cell-diff.js";
import { getWorkbookChangeAuditLog } from "../audit/workbook-change-audit.js";
import {
  getWorkbookRecoveryLog,
  MAX_RECOVERY_CELLS,
  type AppendWorkbookRecoverySnapshotArgs,
  type WorkbookRecoverySnapshot,
} from "../workbook/recovery-log.js";
import { formatAsMarkdownTable } from "../utils/format.js";
import { finalizeMutationOperation } from "./mutation/finalize.js";
import type { MutationFinalizeDependencies } from "./mutation/types.js";
import { appendMutationResultNote } from "./mutation/result-note.js";
import { describeExcelFailure, parseExcelError } from "./mutation/excel-failure.js";
import {
  blockAddress,
  chunkRowsForWrite,
  classifySentValue,
  describeWriteObstacle,
  formatVerificationNotes,
  formatVerificationProblems,
  hasVerificationProblems,
  inspectWriteTarget,
  verifyWrittenRange,
  backupUnavailable,
  sampledChanges,
  toCellGrid,
  type CellScalar,
  type SentGrid,
  type WriteBlock,
  type WriteTargetInspection,
  type WriteVerification,
} from "./mutation/range-write.js";

export { countOccupiedCells } from "./mutation/range-write.js";

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
  | { kind: "written"; inspection: WriteTargetInspection; verification: WriteVerification | null; verifyError: string | null };

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

function errorResult(text: string, details: WriteCellsDetails = { kind: "write_cells", blocked: false }): AgentToolResult<WriteCellsDetails> {
  return { content: [{ type: "text", text }], details, isError: true };
}

export function createWriteCellsTool(
  dependencies: Partial<WriteCellsToolDependencies> = {},
): AgentTool<typeof schema, WriteCellsDetails> {
  const resolvedDependencies: WriteCellsToolDependencies = {
    appendAuditEntry: dependencies.appendAuditEntry ?? defaultDependencies.appendAuditEntry,
    appendRecoverySnapshot:
      dependencies.appendRecoverySnapshot ?? defaultDependencies.appendRecoverySnapshot,
  };
  const mutationFinalizeDependencies: MutationFinalizeDependencies = {
    appendAuditEntry: resolvedDependencies.appendAuditEntry,
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
    execute: async (
      toolCallId: string,
      params: Params,
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
        start = parseCell(startCellRef);
      } catch {
        return errorResult(`Error: invalid start_cell "${params.start_cell}".`);
      }

      const invalidFormulas = findInvalidFormulas(padded, start);
      if (invalidFormulas.length > 0) {
        const lines: string[] = [];
        lines.push("⛔ **Write blocked** — invalid formula syntax detected (nothing was written):");
        for (const invalid of invalidFormulas) {
          lines.push(`- ${invalid.address}: ${invalid.formula} (${invalid.reason})`);
        }
        lines.push("");
        lines.push("Fix the formulas and retry.");
        return errorResult(lines.join("\n"), { kind: "write_cells", blocked: true });
      }

      const block: WriteBlock = { startRow: start.row, startCol: start.col, rowCount: rows, columnCount: cols };
      const sent: SentGrid = (r, c) => classifySentValue(padded[r]?.[c]);
      const sentFormulas = padded.some((row) => row.some((value) => typeof value === "string" && value.startsWith("=")));
      const sentEmptyCount = padded.reduce((total, row) => total + row.filter((value) => value === "").length, 0);

      // Filled in as the Excel batch progresses, so a failure can say what already happened.
      const progress: { inspection: WriteTargetInspection | null; writtenRows: number } = {
        inspection: null,
        writtenRows: 0,
      };

      let outcome: WriteOutcome;
      try {
        outcome = await excelRun<WriteOutcome>(async (context) => {
          const { sheet } = getRange(context, params.start_cell);
          const inspected = await inspectWriteTarget(context, sheet, block, sent);
          progress.inspection = inspected;

          const obstacle = describeWriteObstacle(inspected, (address) => qualifiedAddress(inspected.sheetName, address));
          if (obstacle) return { kind: "obstacle", inspection: inspected, reason: obstacle };
          if (inspected.occupiedCount > 0 && !params.allow_overwrite) return { kind: "occupied", inspection: inspected };
          if (inspected.occupiedCount > 0 && !inspected.before) return { kind: "too_large_to_back_up", inspection: inspected };

          for (const chunk of chunkRowsForWrite(padded)) {
            sheet.getRange(blockAddress(block, chunk.rowOffset, chunk.rows.length)).values = chunk.rows;
            await context.sync();
            progress.writtenRows += chunk.rows.length;
          }

          try {
            const verification = await verifyWrittenRange(context, sheet, block, sent, {
              calculationMode: inspected.calculationMode,
              sentFormulas,
              sentEmptyCount,
            });
            return { kind: "written", inspection: inspected, verification, verifyError: null };
          } catch (verifyError) {
            const failure = await describeExcelFailure(parseExcelError(verifyError), { range: params.start_cell });
            return { kind: "written", inspection: inspected, verification: null, verifyError: failure.reason };
          }
        });
      } catch (error) {
        return failedWrite({
          error: parseExcelError(error),
          params,
          toolCallId,
          block,
          inspection: progress.inspection,
          writtenRows: progress.writtenRows,
          dependencies: resolvedDependencies,
          finalizeDependencies: mutationFinalizeDependencies,
        });
      }

      const fullAddr = qualifiedAddress(outcome.inspection.sheetName, blockAddress(block));

      if (outcome.kind !== "written") {
        const blockedResult = outcome.kind === "obstacle"
          ? errorResult(`⛔ **Write blocked** — nothing was written to ${fullAddr}.\n\n${outcome.reason}`, {
            kind: "write_cells", blocked: true, address: fullAddr,
          })
          : outcome.kind === "occupied"
            ? formatOccupied(outcome.inspection, fullAddr)
            : formatTooLargeToBackUp(outcome.inspection, fullAddr, block);

        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "write_cells",
            toolCallId,
            blocked: true,
            outputAddress: fullAddr,
            changedCount: 0,
            changes: [],
            summary: outcome.kind === "obstacle" ? "blocked: target cannot hold the values" : `blocked: ${outcome.kind.replaceAll("_", " ")}`,
          },
        });
        return blockedResult;
      }

      const successResult = formatWritten(outcome.inspection, outcome.verification, outcome.verifyError, block, padded, fullAddr);
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

function countSentCells(padded: CellScalar[][]): number {
  return padded.reduce((total, row) =>
    total + row.filter((value) => value !== "" && value !== null && value !== undefined).length, 0);
}

interface FailedWriteArgs {
  error: ReturnType<typeof parseExcelError>;
  params: Params;
  toolCallId: string;
  block: WriteBlock;
  inspection: WriteTargetInspection | null;
  writtenRows: number;
  dependencies: WriteCellsToolDependencies;
  finalizeDependencies: MutationFinalizeDependencies;
}

async function failedWrite(args: FailedWriteArgs): Promise<AgentToolResult<WriteCellsDetails>> {
  const { block, inspection, writtenRows } = args;
  const failure = await describeExcelFailure(args.error, { range: args.params.start_cell });
  const fullAddr = inspection
    ? qualifiedAddress(inspection.sheetName, blockAddress(block))
    : args.params.start_cell;

  if (writtenRows === 0 || !inspection) {
    await finalizeMutationOperation(args.finalizeDependencies, {
      auditEntry: {
        toolName: "write_cells",
        toolCallId: args.toolCallId,
        blocked: true,
        outputAddress: fullAddr,
        changedCount: 0,
        changes: [],
        summary: `error: ${failure.reason}`,
      },
    });
    return errorResult(`Error writing cells — nothing was written to ${fullAddr}. ${failure.reason}`, {
      kind: "write_cells",
      blocked: false,
      address: fullAddr,
      ...(failure.code ? { errorCode: failure.code } : {}),
    });
  }

  // Some chunks landed: report exactly which rows, back them up, and audit them.
  const writtenAddr = qualifiedAddress(inspection.sheetName, blockAddress(block, 0, writtenRows));
  const result = errorResult(
    `⚠️ **Partly written** — rows ${block.startRow}–${block.startRow + writtenRows - 1} of ${fullAddr} ` +
    `(${writtenAddr}) were written; the remaining ${block.rowCount - writtenRows} row(s) failed and were not written. ` +
    `${failure.reason}\n\nNothing was rolled back. Read ${writtenAddr} before retrying the rest.`,
    {
      kind: "write_cells",
      blocked: false,
      address: fullAddr,
      ...(failure.code ? { errorCode: failure.code } : {}),
    },
  );
  const before = inspection.before;
  await finalizeMutationOperation(args.finalizeDependencies, {
    auditEntry: {
      toolName: "write_cells",
      toolCallId: args.toolCallId,
      blocked: false,
      outputAddress: writtenAddr,
      changedCount: writtenRows * block.columnCount,
      changes: [],
      summary: `partly written (${writtenRows} of ${block.rowCount} rows): ${failure.reason}`,
    },
    recovery: {
      result,
      appendRecoverySnapshot: () => before
        ? args.dependencies.appendRecoverySnapshot({
          toolName: "write_cells",
          toolCallId: args.toolCallId,
          address: fullAddr,
          changedCount: writtenRows * block.columnCount,
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
  inspection: WriteTargetInspection,
  verification: WriteVerification | null,
  verifyError: string | null,
  block: WriteBlock,
  padded: CellScalar[][],
  fullAddr: string,
): AgentToolResult<WriteCellsDetails> {
  const qualify = (address: string): string => qualifiedAddress(inspection.sheetName, address);
  const startCell = `${colToLetter(block.startCol)}${block.startRow}`;
  const size = `(${block.rowCount}×${block.columnCount})`;

  if (!verification) {
    return {
      content: [{
        type: "text",
        text: `Written to **${fullAddr}** ${size}, but the result could not be read back to verify it: ${verifyError ?? "unknown error"}\n\n` +
          "Read the range before relying on it.",
      }],
      details: { kind: "write_cells", blocked: false, address: fullAddr, verification: "failed" },
    };
  }

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
    : sampledChanges(inspection.sheetName, block, verification, countSentCells(padded));

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
    ...(verification.manualCalculation ? { calculationMode: verification.manualCalculation } : {}),
  };

  return {
    content: [{ type: "text", text: lines.join("\n") }],
    details,
    ...(notWritten > 0 ? { isError: true } : {}),
  };
}
