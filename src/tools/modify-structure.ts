/**
 * modify_structure — Insert/delete rows, columns, and sheets.
 *
 * Single tool for all structural changes (sheets, rows, columns).
 */

import { Type, type Static } from "typebox";
import { StringEnum } from "./string-enum.js";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { excelRun } from "../excel/helpers.js";
import { getWorkbookChangeAuditLog } from "../audit/workbook-change-audit.js";
import { MAX_RECOVERY_CELLS, getWorkbookRecoveryLog } from "../workbook/recovery-log.js";
import {
  captureModifyStructureState,
  captureSheetValueDataRange,
  captureValueDataRange,
  columnNumberToLetter,
  isRecoverySheetVisibility,
} from "../workbook/recovery/structure-state.js";
import type { RecoveryModifyStructureState, RecoverySheetVisibility } from "../workbook/recovery-states.js";
import { getErrorMessage } from "../utils/errors.js";
import type { ModifyStructureDetails } from "./tool-details.js";
import {
  CHECKPOINT_SKIPPED_NOTE,
  CHECKPOINT_SKIPPED_REASON,
} from "./recovery-metadata.js";
import { finalizeMutationOperation } from "./mutation/finalize.js";
import { appendMutationResultNote } from "./mutation/result-note.js";
import type { MutationFinalizeDependencies } from "./mutation/types.js";

// Helper for string enum (TypeBox doesn't have a built-in StringEnum)
const schema = Type.Object({
  action: StringEnum(
    [
      "insert_rows",
      "delete_rows",
      "insert_columns",
      "delete_columns",
      "add_sheet",
      "delete_sheet",
      "rename_sheet",
      "duplicate_sheet",
      "hide_sheet",
      "unhide_sheet",
    ],
    { description: "The structural modification to perform." },
  ),
  sheet: Type.Optional(
    Type.String({
      description:
        "Target sheet name. Required for sheet operations and row/column operations on a specific sheet. " +
        "If omitted for row/column ops, uses the active sheet.",
    }),
  ),
  position: Type.Optional(
    Type.Number({
      description:
        "For insert_rows/delete_rows: the 1-indexed row number. " +
        "For insert_columns/delete_columns: the 1-indexed column number. " +
        "For add_sheet: the 0-indexed position to insert the new sheet.",
    }),
  ),
  count: Type.Optional(
    Type.Number({
      description: "Number of rows or columns to insert/delete. Default: 1.",
    }),
  ),
  new_name: Type.Optional(
    Type.String({
      description: 'New name for rename_sheet or add_sheet. Also used for duplicate_sheet target name.',
    }),
  ),
});

type Params = Static<typeof schema>;

interface StructureMutationResult {
  message: string;
  changedCount: number;
  outputAddress?: string;
  summary: string;
  checkpointState?: RecoveryModifyStructureState;
  checkpointUnavailableReason?: string;
}

function targetSheet(context: Excel.RequestContext, sheetName: string | undefined): Excel.Worksheet {
  return sheetName
    ? context.workbook.worksheets.getItem(sheetName)
    : context.workbook.worksheets.getActiveWorksheet();
}

async function insertRows(
  context: Excel.RequestContext,
  params: Params,
  count: number,
): Promise<StructureMutationResult> {
  if (!params.position) throw new Error("position is required for insert_rows");
  const startRow = params.position;
  const endRow = params.position + count - 1;
  const sheet = targetSheet(context, params.sheet);
  sheet.load("id,name");
  await context.sync();

  sheet.getRange(`${startRow}:${endRow}`).insert("Down");
  await context.sync();

  return {
    message: `Inserted ${count} row(s) at row ${startRow} in "${sheet.name}".`,
    changedCount: count,
    outputAddress: `${sheet.name}!${startRow}:${endRow}`,
    summary: `inserted ${count} row(s)`,
    checkpointState: {
      kind: "rows_absent",
      sheetId: sheet.id,
      sheetName: sheet.name,
      position: startRow,
      count,
    },
  };
}

async function deleteRows(
  context: Excel.RequestContext,
  params: Params,
  count: number,
): Promise<StructureMutationResult> {
  if (!params.position) throw new Error("position is required for delete_rows");
  const startRow = params.position;
  const endRow = params.position + count - 1;
  const sheet = targetSheet(context, params.sheet);
  sheet.load("id,name");
  await context.sync();

  const range = sheet.getRange(`${startRow}:${endRow}`);
  const dataRangeCapture = await captureValueDataRange(context, range, MAX_RECOVERY_CELLS);
  range.delete("Up");
  await context.sync();

  const result: StructureMutationResult = {
    message: `Deleted ${count} row(s) starting at row ${startRow} in "${sheet.name}".`,
    changedCount: count,
    outputAddress: `${sheet.name}!${startRow}:${endRow}`,
    summary: `deleted ${count} row(s)`,
  };
  if (dataRangeCapture.status === "too_large") {
    result.checkpointUnavailableReason =
      "Checkpoint capture was skipped for `delete_rows` because deleted row data exceeds recovery size limits.";
  } else {
    result.checkpointState = {
      kind: "rows_present",
      sheetId: sheet.id,
      sheetName: sheet.name,
      position: startRow,
      count,
      ...(dataRangeCapture.status === "captured" ? { dataRange: dataRangeCapture.dataRange } : {}),
    };
  }
  return result;
}

async function insertColumns(
  context: Excel.RequestContext,
  params: Params,
  count: number,
): Promise<StructureMutationResult> {
  if (!params.position) throw new Error("position is required for insert_columns");
  const startLetter = columnNumberToLetter(params.position);
  const endLetter = columnNumberToLetter(params.position + count - 1);
  const sheet = targetSheet(context, params.sheet);
  sheet.load("id,name");
  await context.sync();

  const range = sheet.getRange(`${startLetter}:${startLetter}`);
  for (let index = 0; index < count; index += 1) range.insert("Right");
  await context.sync();

  return {
    message: `Inserted ${count} column(s) at column ${params.position} (${startLetter}) in "${sheet.name}".`,
    changedCount: count,
    outputAddress: `${sheet.name}!${startLetter}:${endLetter}`,
    summary: `inserted ${count} column(s)`,
    checkpointState: {
      kind: "columns_absent",
      sheetId: sheet.id,
      sheetName: sheet.name,
      position: params.position,
      count,
    },
  };
}

async function deleteColumns(
  context: Excel.RequestContext,
  params: Params,
  count: number,
): Promise<StructureMutationResult> {
  if (!params.position) throw new Error("position is required for delete_columns");
  const startLetter = columnNumberToLetter(params.position);
  const endLetter = columnNumberToLetter(params.position + count - 1);
  const sheet = targetSheet(context, params.sheet);
  sheet.load("id,name");
  await context.sync();

  const range = sheet.getRange(`${startLetter}:${endLetter}`);
  const dataRangeCapture = await captureValueDataRange(context, range, MAX_RECOVERY_CELLS);
  range.delete("Left");
  await context.sync();

  const result: StructureMutationResult = {
    message: `Deleted ${count} column(s) starting at column ${params.position} (${startLetter}) in "${sheet.name}".`,
    changedCount: count,
    outputAddress: `${sheet.name}!${startLetter}:${endLetter}`,
    summary: `deleted ${count} column(s)`,
  };
  if (dataRangeCapture.status === "too_large") {
    result.checkpointUnavailableReason =
      "Checkpoint capture was skipped for `delete_columns` because deleted column data exceeds recovery size limits.";
  } else {
    result.checkpointState = {
      kind: "columns_present",
      sheetId: sheet.id,
      sheetName: sheet.name,
      position: params.position,
      count,
      ...(dataRangeCapture.status === "captured" ? { dataRange: dataRangeCapture.dataRange } : {}),
    };
  }
  return result;
}

async function addSheet(
  context: Excel.RequestContext,
  params: Params,
): Promise<StructureMutationResult> {
  const requestedName = params.new_name || `Sheet${Date.now()}`;
  const newSheet = context.workbook.worksheets.add(requestedName);
  if (params.position !== undefined) newSheet.position = params.position;
  newSheet.load("id,name");
  await context.sync();

  return {
    message: `Added sheet "${newSheet.name}".`,
    changedCount: 1,
    outputAddress: newSheet.name,
    summary: `added sheet ${newSheet.name}`,
    checkpointState: {
      kind: "sheet_absent",
      sheetId: newSheet.id,
      sheetName: newSheet.name,
    },
  };
}

async function deleteSheet(
  context: Excel.RequestContext,
  params: Params,
): Promise<StructureMutationResult> {
  if (!params.sheet) throw new Error("sheet name is required for delete_sheet");
  const sheet = context.workbook.worksheets.getItem(params.sheet);
  sheet.load("id,name,position,visibility");
  await context.sync();

  const result: StructureMutationResult = {
    message: `Deleted sheet "${sheet.name}".`,
    changedCount: 1,
    outputAddress: sheet.name,
    summary: `deleted sheet ${sheet.name}`,
  };
  if (isRecoverySheetVisibility(sheet.visibility)) {
    const dataRangeCapture = await captureSheetValueDataRange(context, sheet, MAX_RECOVERY_CELLS);
    if (dataRangeCapture.status === "too_large") {
      result.checkpointUnavailableReason =
        "Checkpoint capture was skipped for `delete_sheet` because deleted sheet data exceeds recovery size limits.";
    } else {
      const visibility: RecoverySheetVisibility = sheet.visibility;
      result.checkpointState = {
        kind: "sheet_present",
        sheetId: sheet.id,
        sheetName: sheet.name,
        position: sheet.position,
        visibility,
        ...(dataRangeCapture.status === "captured" ? { dataRange: dataRangeCapture.dataRange } : {}),
      };
    }
  } else {
    result.checkpointUnavailableReason =
      "Checkpoint capture was skipped for `delete_sheet` (sheet visibility unsupported).";
  }

  sheet.delete();
  await context.sync();
  return result;
}

async function renameSheet(
  context: Excel.RequestContext,
  params: Params,
): Promise<StructureMutationResult> {
  if (!params.sheet) throw new Error("sheet name is required for rename_sheet");
  if (!params.new_name) throw new Error("new_name is required for rename_sheet");
  const previousName = params.sheet;
  const sheet = context.workbook.worksheets.getItem(previousName);
  sheet.name = params.new_name;
  await context.sync();
  return {
    message: `Renamed sheet "${previousName}" to "${params.new_name}".`,
    changedCount: 1,
    outputAddress: params.new_name,
    summary: `renamed sheet ${previousName} to ${params.new_name}`,
  };
}

async function duplicateSheet(
  context: Excel.RequestContext,
  params: Params,
): Promise<StructureMutationResult> {
  if (!params.sheet) throw new Error("sheet name is required for duplicate_sheet");
  if (params.new_name) {
    const existingTarget = context.workbook.worksheets.getItemOrNullObject(params.new_name);
    existingTarget.load("isNullObject");
    await context.sync();
    if (!existingTarget.isNullObject) {
      throw new Error(`A worksheet named "${params.new_name}" already exists.`);
    }
  }

  const source = context.workbook.worksheets.getItem(params.sheet);
  const copy = source.copy("End");
  copy.load("id,name");
  await context.sync();

  if (params.new_name) {
    copy.name = params.new_name;
    await context.sync();
    copy.load("id,name");
    await context.sync();
  }

  const targetName = copy.name;
  let checkpointState: RecoveryModifyStructureState | undefined;
  let checkpointUnavailableReason: string | undefined;
  try {
    const usedRange = copy.getUsedRangeOrNullObject(true);
    usedRange.load("isNullObject");
    await context.sync();
    if (usedRange.isNullObject) {
      checkpointState = { kind: "sheet_absent", sheetId: copy.id, sheetName: targetName };
    } else {
      checkpointUnavailableReason =
        "Checkpoint capture was skipped for `duplicate_sheet` because duplicated sheet contains data.";
    }
  } catch (inspectionError) {
    checkpointUnavailableReason =
      `Checkpoint capture was skipped after \`duplicate_sheet\` completed: ${getErrorMessage(inspectionError)}`;
  }

  return {
    message: `Duplicated "${params.sheet}" as "${targetName}".`,
    changedCount: 1,
    outputAddress: targetName,
    summary: `duplicated sheet ${params.sheet} as ${targetName}`,
    ...(checkpointState !== undefined ? { checkpointState } : {}),
    ...(checkpointUnavailableReason !== undefined ? { checkpointUnavailableReason } : {}),
  };
}

async function setSheetVisibility(
  context: Excel.RequestContext,
  params: Params,
  visibility: "Hidden" | "Visible",
): Promise<StructureMutationResult> {
  if (!params.sheet) throw new Error(`sheet name is required for ${params.action}`);
  const sheet = context.workbook.worksheets.getItem(params.sheet);
  sheet.visibility = visibility;
  await context.sync();
  const verb = visibility === "Hidden" ? "Hidden" : "Unhidden";
  return {
    message: `${verb} sheet "${params.sheet}".`,
    changedCount: 1,
    outputAddress: params.sheet,
    summary: `${verb.toLowerCase()} sheet ${params.sheet}`,
  };
}

async function executeStructureMutation(
  context: Excel.RequestContext,
  params: Params,
): Promise<StructureMutationResult> {
  const count = typeof params.count === "number" && Number.isFinite(params.count) && params.count > 0
    ? Math.floor(params.count)
    : 1;
  switch (params.action) {
    case "insert_rows": return insertRows(context, params, count);
    case "delete_rows": return deleteRows(context, params, count);
    case "insert_columns": return insertColumns(context, params, count);
    case "delete_columns": return deleteColumns(context, params, count);
    case "add_sheet": return addSheet(context, params);
    case "delete_sheet": return deleteSheet(context, params);
    case "rename_sheet": return renameSheet(context, params);
    case "duplicate_sheet": return duplicateSheet(context, params);
    case "hide_sheet": return setSheetVisibility(context, params, "Hidden");
    case "unhide_sheet": return setSheetVisibility(context, params, "Visible");
  }
}

const mutationFinalizeDependencies: MutationFinalizeDependencies = {
  appendAuditEntry: (entry) => getWorkbookChangeAuditLog().append(entry),
};

export function createModifyStructureTool(): AgentTool<typeof schema, ModifyStructureDetails> {
  return {
    name: "modify_structure",
    label: "Modify Structure",
    description:
      "Modify the workbook structure: insert/delete rows and columns, " +
      "add/delete/rename/duplicate/hide/unhide sheets. " +
      "Be careful with deletions — there is no undo.",
    parameters: schema,
    execute: async (
      toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<ModifyStructureDetails>> => {
      try {
        let preMutationCheckpointState: RecoveryModifyStructureState | null = null;
        let checkpointUnavailableReason: string | null = null;

        if (
          (params.action === "rename_sheet" || params.action === "hide_sheet" || params.action === "unhide_sheet") &&
          typeof params.sheet === "string" &&
          params.sheet.trim().length > 0
        ) {
          preMutationCheckpointState = await captureModifyStructureState({
            kind: params.action === "rename_sheet" ? "sheet_name" : "sheet_visibility",
            sheetRef: params.sheet,
          });

          if (!preMutationCheckpointState) {
            checkpointUnavailableReason =
              `Checkpoint capture was skipped for \`${params.action}\` (sheet state unavailable).`;
          }
        }

        const result = await excelRun((context) => executeStructureMutation(context, params));

        const toolResult: AgentToolResult<ModifyStructureDetails> = {
          content: [{ type: "text", text: result.message }],
          details: {
            kind: "modify_structure",
            action: params.action,
          },
        };

        const checkpointAddress = result.outputAddress ?? params.sheet ?? params.action;
        const checkpointState = result.checkpointState ?? preMutationCheckpointState;
        const recoveryUnavailableReason = checkpointState
          ? CHECKPOINT_SKIPPED_REASON
          : (result.checkpointUnavailableReason ?? checkpointUnavailableReason ?? CHECKPOINT_SKIPPED_REASON);

        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "modify_structure",
            toolCallId,
            blocked: false,
            ...(result.outputAddress !== undefined ? { outputAddress: result.outputAddress } : {}),
            changedCount: result.changedCount,
            changes: [],
            summary: result.summary,
          },
          recovery: {
            result: toolResult,
            appendRecoverySnapshot: () => {
              if (!checkpointState) {
                return Promise.resolve(null);
              }

              return getWorkbookRecoveryLog().appendModifyStructure({
                toolName: "modify_structure",
                toolCallId,
                address: checkpointAddress,
                changedCount: result.changedCount,
                modifyStructureState: checkpointState,
              });
            },
            appendResultNote: appendMutationResultNote,
            unavailableReason: recoveryUnavailableReason,
            unavailableNote: CHECKPOINT_SKIPPED_NOTE,
          },
        });

        return toolResult;
      } catch (e) {
        const message = getErrorMessage(e);

        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "modify_structure",
            toolCallId,
            blocked: true,
            ...(params.sheet !== undefined ? { outputAddress: params.sheet } : {}),
            changedCount: 0,
            changes: [],
            summary: `error: ${message}`,
          },
        });

        return {
          content: [{ type: "text", text: `Error: ${message}` }],
          details: {
            kind: "modify_structure",
            action: params.action,
          },
        };
      }
    },
  };
}
