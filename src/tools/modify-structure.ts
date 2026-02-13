/**
 * modify_structure — Insert/delete rows, columns, and sheets.
 *
 * Single tool for all structural changes (sheets, rows, columns).
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun } from "../excel/helpers.js";
import { getWorkbookChangeAuditLog } from "../audit/workbook-change-audit.js";
import { dispatchWorkbookSnapshotCreated } from "../workbook/recovery-events.js";
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
function StringEnum<T extends string[]>(values: [...T], opts?: { description?: string }) {
  return Type.Union(
    values.map((v) => Type.Literal(v)),
    opts,
  );
}

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

type SupportedStructureCheckpointAction =
  | "rename_sheet"
  | "hide_sheet"
  | "unhide_sheet"
  | "insert_rows"
  | "delete_rows"
  | "insert_columns"
  | "delete_columns"
  | "add_sheet"
  | "delete_sheet"
  | "duplicate_sheet";

type PreMutationCapturedStructureCheckpointAction = "rename_sheet" | "hide_sheet" | "unhide_sheet";

function supportedCheckpointActionFor(
  action: Params["action"],
): SupportedStructureCheckpointAction | null {
  if (
    action === "rename_sheet" ||
    action === "hide_sheet" ||
    action === "unhide_sheet" ||
    action === "insert_rows" ||
    action === "delete_rows" ||
    action === "insert_columns" ||
    action === "delete_columns" ||
    action === "add_sheet" ||
    action === "delete_sheet" ||
    action === "duplicate_sheet"
  ) {
    return action;
  }

  return null;
}

function preMutationCheckpointKindFor(
  action: PreMutationCapturedStructureCheckpointAction,
): "sheet_name" | "sheet_visibility" {
  return action === "rename_sheet" ? "sheet_name" : "sheet_visibility";
}

function unsupportedStructureCheckpointReason(action: Params["action"]): string {
  return `Checkpoint capture is not yet supported for modify_structure action \`${action}\`.`;
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
        const checkpointAction = supportedCheckpointActionFor(params.action);
        let preMutationCheckpointState: RecoveryModifyStructureState | null = null;
        let checkpointUnavailableReason = checkpointAction
          ? null
          : unsupportedStructureCheckpointReason(params.action);

        if (
          (params.action === "rename_sheet" || params.action === "hide_sheet" || params.action === "unhide_sheet") &&
          typeof params.sheet === "string" &&
          params.sheet.trim().length > 0
        ) {
          preMutationCheckpointState = await captureModifyStructureState({
            kind: preMutationCheckpointKindFor(params.action),
            sheetRef: params.sheet,
          });

          if (!preMutationCheckpointState) {
            checkpointUnavailableReason =
              `Checkpoint capture was skipped for \`${params.action}\` (sheet state unavailable).`;
          }
        }

        const result = await excelRun<StructureMutationResult>(async (context) => {
          const action = params.action;
          const count = typeof params.count === "number" && Number.isFinite(params.count) && params.count > 0
            ? Math.floor(params.count)
            : 1;

          const getSheet = () => {
            if (params.sheet) {
              return context.workbook.worksheets.getItem(params.sheet);
            }
            return context.workbook.worksheets.getActiveWorksheet();
          };

          switch (action) {
            case "insert_rows": {
              if (!params.position) throw new Error("position is required for insert_rows");
              const startRow = params.position;
              const endRow = params.position + count - 1;
              const sheet = getSheet();
              sheet.load("id,name");
              await context.sync();

              const range = sheet.getRange(`${startRow}:${endRow}`);
              range.insert("Down");
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

            case "delete_rows": {
              if (!params.position) throw new Error("position is required for delete_rows");
              const startRow = params.position;
              const endRow = params.position + count - 1;
              const sheet = getSheet();
              sheet.load("id,name");
              await context.sync();

              const range = sheet.getRange(`${startRow}:${endRow}`);
              const dataRangeCapture = await captureValueDataRange(context, range, MAX_RECOVERY_CELLS);
              range.delete("Up");
              await context.sync();

              if (dataRangeCapture.status === "too_large") {
                return {
                  message: `Deleted ${count} row(s) starting at row ${startRow} in "${sheet.name}".`,
                  changedCount: count,
                  outputAddress: `${sheet.name}!${startRow}:${endRow}`,
                  summary: `deleted ${count} row(s)`,
                  checkpointUnavailableReason: "Checkpoint capture was skipped for `delete_rows` because deleted row data exceeds recovery size limits.",
                };
              }

              return {
                message: `Deleted ${count} row(s) starting at row ${startRow} in "${sheet.name}".`,
                changedCount: count,
                outputAddress: `${sheet.name}!${startRow}:${endRow}`,
                summary: `deleted ${count} row(s)`,
                checkpointState: {
                  kind: "rows_present",
                  sheetId: sheet.id,
                  sheetName: sheet.name,
                  position: startRow,
                  count,
                  ...(dataRangeCapture.status === "captured" ? { dataRange: dataRangeCapture.dataRange } : {}),
                },
              };
            }

            case "insert_columns": {
              if (!params.position) throw new Error("position is required for insert_columns");
              const startLetter = columnNumberToLetter(params.position);
              const endLetter = columnNumberToLetter(params.position + count - 1);
              const sheet = getSheet();
              sheet.load("id,name");
              await context.sync();

              const range = sheet.getRange(`${startLetter}:${startLetter}`);
              for (let index = 0; index < count; index += 1) {
                range.insert("Right");
              }
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

            case "delete_columns": {
              if (!params.position) throw new Error("position is required for delete_columns");
              const startLetter = columnNumberToLetter(params.position);
              const endLetter = columnNumberToLetter(params.position + count - 1);
              const sheet = getSheet();
              sheet.load("id,name");
              await context.sync();

              const range = sheet.getRange(`${startLetter}:${endLetter}`);
              const dataRangeCapture = await captureValueDataRange(context, range, MAX_RECOVERY_CELLS);
              range.delete("Left");
              await context.sync();

              if (dataRangeCapture.status === "too_large") {
                return {
                  message: `Deleted ${count} column(s) starting at column ${params.position} (${startLetter}) in "${sheet.name}".`,
                  changedCount: count,
                  outputAddress: `${sheet.name}!${startLetter}:${endLetter}`,
                  summary: `deleted ${count} column(s)`,
                  checkpointUnavailableReason: "Checkpoint capture was skipped for `delete_columns` because deleted column data exceeds recovery size limits.",
                };
              }

              return {
                message: `Deleted ${count} column(s) starting at column ${params.position} (${startLetter}) in "${sheet.name}".`,
                changedCount: count,
                outputAddress: `${sheet.name}!${startLetter}:${endLetter}`,
                summary: `deleted ${count} column(s)`,
                checkpointState: {
                  kind: "columns_present",
                  sheetId: sheet.id,
                  sheetName: sheet.name,
                  position: params.position,
                  count,
                  ...(dataRangeCapture.status === "captured" ? { dataRange: dataRangeCapture.dataRange } : {}),
                },
              };
            }

            case "add_sheet": {
              const requestedName = params.new_name || `Sheet${Date.now()}`;
              const newSheet = context.workbook.worksheets.add(requestedName);
              if (params.position !== undefined) {
                newSheet.position = params.position;
              }
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

            case "delete_sheet": {
              if (!params.sheet) throw new Error("sheet name is required for delete_sheet");
              const sheet = context.workbook.worksheets.getItem(params.sheet);
              sheet.load("id,name,position,visibility");
              await context.sync();

              const visibility = sheet.visibility;
              let checkpointState: RecoveryModifyStructureState | undefined;
              let checkpointUnavailableReason: string | undefined;

              if (isRecoverySheetVisibility(visibility)) {
                const dataRangeCapture = await captureSheetValueDataRange(context, sheet, MAX_RECOVERY_CELLS);
                if (dataRangeCapture.status === "too_large") {
                  checkpointUnavailableReason = "Checkpoint capture was skipped for `delete_sheet` because deleted sheet data exceeds recovery size limits.";
                } else {
                  const sheetVisibility: RecoverySheetVisibility = visibility;
                  checkpointState = {
                    kind: "sheet_present",
                    sheetId: sheet.id,
                    sheetName: sheet.name,
                    position: sheet.position,
                    visibility: sheetVisibility,
                    ...(dataRangeCapture.status === "captured" ? { dataRange: dataRangeCapture.dataRange } : {}),
                  };
                }
              } else {
                checkpointUnavailableReason = "Checkpoint capture was skipped for `delete_sheet` (sheet visibility unsupported).";
              }

              sheet.delete();
              await context.sync();
              return {
                message: `Deleted sheet "${sheet.name}".`,
                changedCount: 1,
                outputAddress: sheet.name,
                summary: `deleted sheet ${sheet.name}`,
                checkpointState,
                checkpointUnavailableReason,
              };
            }

            case "rename_sheet": {
              if (!params.sheet) throw new Error("sheet name is required for rename_sheet");
              if (!params.new_name) throw new Error("new_name is required for rename_sheet");
              const previousName = params.sheet;
              const newName = params.new_name;
              const sheet = context.workbook.worksheets.getItem(previousName);
              sheet.name = newName;
              await context.sync();
              return {
                message: `Renamed sheet "${previousName}" to "${newName}".`,
                changedCount: 1,
                outputAddress: newName,
                summary: `renamed sheet ${previousName} to ${newName}`,
              };
            }

            case "duplicate_sheet": {
              if (!params.sheet) throw new Error("sheet name is required for duplicate_sheet");
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
              const usedRange = copy.getUsedRangeOrNullObject(true);
              usedRange.load("isNullObject");
              await context.sync();

              const canCreateCheckpoint = usedRange.isNullObject;
              return {
                message: `Duplicated "${params.sheet}" as "${targetName}".`,
                changedCount: 1,
                outputAddress: targetName,
                summary: `duplicated sheet ${params.sheet} as ${targetName}`,
                checkpointState: canCreateCheckpoint
                  ? {
                    kind: "sheet_absent",
                    sheetId: copy.id,
                    sheetName: targetName,
                  }
                  : undefined,
                checkpointUnavailableReason: canCreateCheckpoint
                  ? undefined
                  : "Checkpoint capture was skipped for `duplicate_sheet` because duplicated sheet contains data.",
              };
            }

            case "hide_sheet": {
              if (!params.sheet) throw new Error("sheet name is required for hide_sheet");
              const sheet = context.workbook.worksheets.getItem(params.sheet);
              sheet.visibility = "Hidden";
              await context.sync();
              return {
                message: `Hidden sheet "${params.sheet}".`,
                changedCount: 1,
                outputAddress: params.sheet,
                summary: `hidden sheet ${params.sheet}`,
              };
            }

            case "unhide_sheet": {
              if (!params.sheet) throw new Error("sheet name is required for unhide_sheet");
              const sheet = context.workbook.worksheets.getItem(params.sheet);
              sheet.visibility = "Visible";
              await context.sync();
              return {
                message: `Unhidden sheet "${params.sheet}".`,
                changedCount: 1,
                outputAddress: params.sheet,
                summary: `unhidden sheet ${params.sheet}`,
              };
            }

            default:
              throw new Error(`Unknown action: ${String(action)}`);
          }
        });

        const toolResult: AgentToolResult<ModifyStructureDetails> = {
          content: [{ type: "text", text: result.message }],
          details: {
            kind: "modify_structure",
            action: params.action,
          },
        };

        const checkpointAddress = result.outputAddress ?? params.sheet ?? params.action;
        const checkpointState = result.checkpointState ?? preMutationCheckpointState;
        const recoveryUnavailableReason = checkpointAction && checkpointState
          ? CHECKPOINT_SKIPPED_REASON
          : (checkpointAction
            ? (result.checkpointUnavailableReason ?? checkpointUnavailableReason ?? CHECKPOINT_SKIPPED_REASON)
            : unsupportedStructureCheckpointReason(params.action));

        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "modify_structure",
            toolCallId,
            blocked: false,
            outputAddress: result.outputAddress,
            changedCount: result.changedCount,
            changes: [],
            summary: result.summary,
          },
          recovery: {
            result: toolResult,
            appendRecoverySnapshot: () => {
              if (!checkpointAction || !checkpointState) {
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
            dispatchSnapshotCreated: (checkpoint) => {
              dispatchWorkbookSnapshotCreated({
                snapshotId: checkpoint.id,
                toolName: checkpoint.toolName,
                address: checkpoint.address,
                changedCount: checkpoint.changedCount,
              });
            },
          },
        });

        return toolResult;
      } catch (e: unknown) {
        const message = getErrorMessage(e);

        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "modify_structure",
            toolCallId,
            blocked: true,
            outputAddress: params.sheet,
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
