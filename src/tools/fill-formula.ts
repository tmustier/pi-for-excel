/**
 * fill_formula — Fill a single formula across a range with Excel's AutoFill.
 *
 * This avoids constructing large 2D formula arrays. References adjust automatically.
 * The fill is synced on its own and verified with bounded reads (see
 * `mutation/range-write.ts`), so a 200k-row fill reports what landed instead of
 * failing on the read-back after the change was applied.
 */

import { runMutation } from "./mutation/run-mutation.js";
import { Type, type Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import type { FillFormulaDetails } from "./tool-details.js";
import { getRange, qualifiedAddress, colToLetter } from "../excel/helpers.js";
import { buildWorkbookCellChangeSummary } from "../audit/cell-diff.js";
import { getWorkbookChangeAuditLog } from "../audit/workbook-change-audit.js";
import { getWorkbookRecoveryLog, MAX_RECOVERY_CELLS } from "../workbook/recovery-log.js";
import { validateFormula } from "./write-cells.js";
import { STOPPED_AFTER_CHANGE_NOTE, type StopReporting } from "./with-workbook-coordinator.js";
import { finalizeMutationOperation } from "./mutation/finalize.js";
import type { MutationFinalizeDependencies } from "./mutation/types.js";
import { appendMutationResultNote } from "./mutation/result-note.js";
import { describeExcelFailure, parseExcelError } from "./mutation/excel-failure.js";
import {
  backupUnavailable,
  blockAddress,
  describeWriteObstacle,
  formatVerificationNotes,
  formatVerificationProblems,
  inspectWriteTarget,
  sampledChanges,
  verifyWrittenRange,
  writeBlockFromAddress,
  type SentGrid,
  type WriteBlock,
  type WriteTargetInspection,
  type WriteReadBack,
  type WriteVerification,
} from "./mutation/range-write.js";

const schema = Type.Object({
  range: Type.String({
    description: 'Target range to fill, e.g. "B2:B20" or "Sheet1!C3:F20". Single contiguous range only.',
  }),
  formula: Type.String({
    description: 'Formula to fill, starting with "=", e.g. "=SUM(B2:B10)". Relative refs will adjust.',
  }),
  allow_overwrite: Type.Optional(
    Type.Boolean({
      description:
        "Set to true to overwrite existing data. Default: false. " +
        "If false and the target range contains data, the fill is blocked.",
    }),
  ),
});

type Params = Static<typeof schema>;

type FillOutcome =
  | { kind: "obstacle"; inspection: WriteTargetInspection; block: WriteBlock; reason: string }
  | { kind: "occupied"; inspection: WriteTargetInspection; block: WriteBlock }
  | { kind: "too_large_to_back_up"; inspection: WriteTargetInspection; block: WriteBlock }
  | ({
    kind: "filled";
    inspection: WriteTargetInspection;
    block: WriteBlock;
  } & WriteReadBack)
  | { kind: "stopped"; inspection: WriteTargetInspection; block: WriteBlock };

const mutationFinalizeDependencies: MutationFinalizeDependencies = {
  appendAuditEntry: (entry) => getWorkbookChangeAuditLog().append(entry),
};

function errorResult(text: string, details: FillFormulaDetails = { kind: "fill_formula", blocked: false }): AgentToolResult<FillFormulaDetails> {
  return { content: [{ type: "text", text }], details, isError: true };
}

export function createFillFormulaTool(): AgentTool<typeof schema, FillFormulaDetails> & StopReporting {
  return {
    name: "fill_formula",
    label: "Fill Formula",
    description:
      "Fill a single formula across a range using Excel's AutoFill. " +
      "Relative references adjust automatically. Use this instead of building large formula arrays.",
    parameters: schema,
    reportsStop: true,
    execute: async (toolCallId: string, params: Params, signal?: AbortSignal): Promise<AgentToolResult<FillFormulaDetails>> => {
      if (!params.formula.startsWith("=")) {
        return errorResult("Error: formula must start with '='.");
      }

      const invalid = validateFormula(params.formula);
      if (invalid) {
        return errorResult(`Error: invalid formula (${invalid}).`);
      }

      if (/[;,]/.test(params.range)) {
        return errorResult("Error: fill_formula only supports a single contiguous range.");
      }

      const sent: SentGrid = () => ({ kind: "formula", formula: params.formula });
      const progress: { inspection: WriteTargetInspection | null; filled: boolean; fillInFlight: boolean } =
        { inspection: null, filled: false, fillInFlight: false };

      let outcome: FillOutcome;
      try {
        outcome = await runMutation<FillOutcome>(async (context) => {
          const { sheet, range } = getRange(context, params.range);
          range.load("address");
          await context.sync();

          const block = writeBlockFromAddress(range.address);
          const inspection = await inspectWriteTarget(context, sheet, block, sent);
          progress.inspection = inspection;

          const obstacle = describeWriteObstacle(inspection, (address) => qualifiedAddress(inspection.sheetName, address));
          if (obstacle) return { kind: "obstacle", inspection, block, reason: obstacle };
          if (inspection.occupiedCount > 0 && !params.allow_overwrite) return { kind: "occupied", inspection, block };
          if (inspection.occupiedCount > 0 && !inspection.before) return { kind: "too_large_to_back_up", inspection, block };
          if (signal?.aborted) return { kind: "stopped", inspection, block };

          // Fill by writing the formula's R1C1 form to every cell: the same
          // relative references autofill produces, but only formulas change.
          // Autofill also copied the top-left cell's formats across the
          // range, which the backup (values and formulas) cannot undo.
          const topLeft = range.getCell(0, 0);
          topLeft.formulas = [[params.formula]];
          topLeft.load("formulasR1C1");
          await context.sync();
          const storedR1C1: unknown = topLeft.formulasR1C1[0]?.[0];
          const r1c1: string = typeof storedR1C1 === "string" ? storedR1C1 : params.formula;
          range.formulasR1C1 = Array.from({ length: range.rowCount }, () =>
            Array.from({ length: range.columnCount }, () => r1c1));

          range.load("values,formulas,address,rowCount,columnCount");
          await context.sync();
          progress.fillInFlight = false;
          progress.filled = true;
          if (signal?.aborted) return { kind: "filled", inspection, block, verification: null, verifyError: null, stopped: true };

          try {
            const verification = await verifyWrittenRange(context, sheet, block, sent, {
              calculationMode: inspection.calculationMode,
              sentFormulas: true,
              sentEmptyCount: 0,
              tablesBefore: inspection.tableNames,
              sentChars: params.formula.length * block.rowCount * block.columnCount,
            });
            return { kind: "filled", inspection, block, verification, verifyError: null, stopped: false };
          } catch (verifyError) {
            const failure = await describeExcelFailure(parseExcelError(verifyError), { range: params.range });
            return { kind: "filled", inspection, block, verification: null, verifyError: failure.reason, stopped: false };
          }
        });
      } catch (error) {
        // Snapshot before any await: an abandoned run (runMutation) may still finish its in-flight fill.
        const filled = progress.filled;
        const fillInFlight = progress.fillInFlight;
        const parsed = parseExcelError(error);
        const failure = await describeExcelFailure(
          parsed,
          { range: params.range },
          !filled && !fillInFlight,
        );
        // Excel never answered the fill batch (timeout, Stop, not an Excel error): it may have landed.
        const fillUnknown = fillInFlight && (parsed.outcomeUnknown || parsed.code === null);
        const inspection = progress.inspection;
        const address = inspection ? qualifiedAddress(inspection.sheetName, params.range) : params.range;
        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "fill_formula",
            toolCallId,
            blocked: !filled && !fillUnknown,
            outputAddress: address,
            changedCount: 0,
            changes: [],
            summary: `error: ${failure.reason}`,
          },
        });
        return errorResult(
          filled
            ? `The fill of ${address} was applied, but Pi could not finish afterwards: ${failure.reason}\n\nRead the range before relying on it.`
            : fillUnknown
              ? `⚠️ **Outcome unknown** — the fill of ${address} may or may not have been applied: Excel did not confirm it. ` +
                `${failure.reason}\n\nRead ${address} before retrying.`
              : `Error filling formula — nothing was written to ${address}. ${failure.reason}`,
          { kind: "fill_formula", blocked: false, address, ...(failure.code ? { errorCode: failure.code } : {}) },
        );
      }

      const fullAddr = qualifiedAddress(outcome.inspection.sheetName, blockAddress(outcome.block));

      if (outcome.kind === "stopped") {
        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "fill_formula",
            toolCallId,
            blocked: true,
            outputAddress: fullAddr,
            changedCount: 0,
            changes: [],
            summary: "stopped before filling",
          },
        });
        return errorResult(`⏹ **Stopped before filling** — nothing changed in ${fullAddr}.`, {
          kind: "fill_formula", blocked: false, address: fullAddr, stopped: true, writtenCount: 0,
        });
      }

      if (outcome.kind !== "filled") {
        const text = outcome.kind === "obstacle"
          ? `⛔ **Fill blocked** — nothing was written to ${fullAddr}.\n\n${outcome.reason}`
          : outcome.kind === "occupied"
            ? `⛔ **Fill blocked** — ${fullAddr} contains ${outcome.inspection.occupiedCount} non-empty cell(s); nothing was written.\n\n` +
              "If the user asked to replace them, retry with `allow_overwrite: true`; otherwise ask the user first."
            : `⛔ **Fill refused** — nothing was written. ${fullAddr} already holds data, and the fill covers ` +
              `${(outcome.block.rowCount * outcome.block.columnCount).toLocaleString("en-US")} cells, more than the ` +
              `${MAX_RECOVERY_CELLS.toLocaleString("en-US")} a backup can hold, so the overwrite could not be undone.\n\n` +
              `Fill it in blocks of at most ${MAX_RECOVERY_CELLS.toLocaleString("en-US")} cells (each gets its own backup), ` +
              "or ask the user to clear the old data first.";

        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "fill_formula",
            toolCallId,
            blocked: true,
            outputAddress: fullAddr,
            changedCount: 0,
            changes: [],
            summary: `blocked: ${outcome.kind.replaceAll("_", " ")}`,
          },
        });

        return errorResult(text, {
          kind: "fill_formula",
          blocked: true,
          address: fullAddr,
          ...(outcome.kind === "obstacle" ? {} : { existingCount: outcome.inspection.occupiedCount }),
        });
      }

      const { inspection, block, verification } = outcome;
      const toolResult: AgentToolResult<FillFormulaDetails> = outcome.stopped
        ? {
          content: [{
            type: "text",
            text: `⏹ **Stopped after filling ${fullAddr}** — the change stands. ` +
              "It was not verified because Stop was pressed; read the range before relying on it.",
          }],
          details: {
            kind: "fill_formula", blocked: false, address: fullAddr, verification: "failed", stopped: true,
            writtenCount: block.rowCount * block.columnCount,
          },
        }
        : verification
        ? formatFilled(inspection, block, verification, params.formula, fullAddr)
        : {
          content: [{
            type: "text",
            text: `Filled formula across **${fullAddr}**, but the result could not be read back to verify it: ${outcome.verifyError}\n\n` +
              "Read the range before relying on it.",
          }],
          details: { kind: "fill_formula", blocked: false, address: fullAddr, verification: "failed" },
        };

      if (signal?.aborted && !outcome.stopped) appendMutationResultNote(toolResult, STOPPED_AFTER_CHANGE_NOTE);

      const before = inspection.before;
      await finalizeMutationOperation(mutationFinalizeDependencies, {
        auditEntry: {
          toolName: "fill_formula",
          toolCallId,
          blocked: false,
          outputAddress: fullAddr,
          changedCount: toolResult.details.changes?.changedCount ?? block.rowCount * block.columnCount,
          changes: toolResult.details.changes?.sample ?? [],
        },
        recovery: {
          result: toolResult,
          appendRecoverySnapshot: () => before
            ? getWorkbookRecoveryLog().append({
              toolName: "fill_formula",
              toolCallId,
              address: fullAddr,
              changedCount: toolResult.details.changes?.changedCount ?? 0,
              beforeValues: before.values,
              beforeFormulas: before.formulas,
            })
            : Promise.resolve(null),
          appendResultNote: appendMutationResultNote,
          ...backupUnavailable(before !== null, block, fullAddr),
        },
      });

      return toolResult;
    },
  };
}

function formulaText(raw: string | number | boolean | null | undefined): string {
  return typeof raw === "string" ? raw : raw === null || raw === undefined ? "" : String(raw);
}

function formatFilled(
  inspection: WriteTargetInspection,
  block: WriteBlock,
  verification: WriteVerification,
  formula: string,
  fullAddr: string,
): AgentToolResult<FillFormulaDetails> {
  const qualify = (address: string): string => qualifiedAddress(inspection.sheetName, address);
  const lines: string[] = [];
  const notFilled = verification.notWrittenCount;
  const totalCells = block.rowCount * block.columnCount;

  lines.push(notFilled > 0
    ? `⚠️ **Partly filled** **${fullAddr}** (${block.rowCount}×${block.columnCount}): ${notFilled} of ${totalCells} cell(s) did not get the formula.`
    : `Filled formula across **${fullAddr}** (${block.rowCount}×${block.columnCount})`);
  lines.push(`**Formula pattern:** \`${formula}\``);

  const formulas = verification.formulas;
  const head = verification.sampleRows[0];
  const tail = verification.sampleRows.at(-1);
  const topLeft = formulaText(formulas ? formulas[0]?.[0] : head?.formulas[0]?.[0]);
  const bottomRight = formulaText(formulas
    ? formulas[block.rowCount - 1]?.[block.columnCount - 1]
    : tail?.formulas.at(-1)?.[block.columnCount - 1]);
  if (topLeft && bottomRight && totalCells > 1) {
    lines.push(`**Example formulas:** top-left \`${topLeft}\`, bottom-right \`${bottomRight}\``);
  }

  const problems = formatVerificationProblems(verification, qualify);
  if (problems.length > 0) {
    lines.push("");
    lines.push(...problems);
  }

  const notes = formatVerificationNotes(verification);
  if (notes.length > 0) {
    lines.push("");
    lines.push(...notes);
  }

  const changes = verification.values && verification.formulas && inspection.before
    ? buildWorkbookCellChangeSummary({
      sheetName: inspection.sheetName,
      startCell: `${colToLetter(block.startCol)}${block.startRow}`,
      beforeValues: inspection.before.values,
      beforeFormulas: inspection.before.formulas,
      afterValues: verification.values,
      afterFormulas: verification.formulas,
    })
    : sampledChanges(inspection.sheetName, block, verification, totalCells);

  if (changes.changedCount > 0) {
    lines.push("");
    lines.push(`Changed cell(s): ${changes.changedCount}.`);
  }

  const invalidCount = verification.invalidCells.length;
  return {
    content: [{ type: "text", text: lines.join("\n") }],
    details: {
      kind: "fill_formula",
      blocked: false,
      address: fullAddr,
      formulaErrorCount: verification.formulaErrorCount,
      changes,
      verification: verification.mode,
      ...(notFilled > 0 ? { notWrittenCount: notFilled } : {}),
      ...(invalidCount > 0 ? { invalidCount } : {}),
      ...(verification.manualCalculation ? { calculationMode: verification.manualCalculation } : {}),
    },
    ...(notFilled > 0 ? { isError: true } : {}),
  };
}
