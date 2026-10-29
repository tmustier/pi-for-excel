/**
 * Turn an Office.js failure into a specific, actionable reason for the agent.
 *
 * Office.js reports sheet protection as "You cannot perform the requested
 * operation." and drops the error code from `message`. The agent cannot act on
 * that, so this maps the error code (and, for protection, the sheet's actual
 * protection state) to a reason that names the cause and the next step.
 */

import { Type } from "typebox";
import { Value } from "typebox/value";

import { excelRun, parseRangeRef } from "../../excel/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const officeErrorSchema = Type.Object({
  code: Type.String(),
  debugInfo: Type.Optional(Type.Object({
    errorLocation: Type.Optional(Type.String()),
  })),
  waitedMs: Type.Optional(Type.Number()),
  earlierStepReachedExcel: Type.Optional(Type.Boolean()),
  outcomeUnknown: Type.Optional(Type.Boolean()),
});

const CELL_EDIT_MODE_CODE = "InvalidOperationInCellEditMode";
const CELL_EDIT_MODE_WAITED_CODE = "excel_cell_edit_mode";

export function cellEditModeReason(waitedMs = 0): string {
  const waited = waitedMs > 0 ? ` Waited ${Math.round(waitedMs / 1000)} s for the user to finish.` : "";
  return "Excel is in cell-edit mode: the user is typing in a cell, and Excel rejects every workbook request " +
    `until they press Enter, Tab or Esc.${waited} Ask the user to finish editing the cell, then retry.`;
}

export interface ExcelFailure {
  code: string | null;
  reason: string;
  maybeApplied: boolean;
}

export function unchangedClause(failure: ExcelFailure, unchanged: string): string {
  return failure.maybeApplied
    ? " — part of this change may already be in the workbook; check it before retrying"
    : ` — ${unchanged}`;
}

export interface ExcelFailureTarget {
  sheet?: string;
  range?: string;
}

const PAYLOAD_LIMIT_CODES = new Set([
  "ResponsePayloadSizeLimitExceeded",
  "RequestPayloadSizeLimitExceeded",
  "PayloadTooLarge",
]);

export interface ExcelErrorInfo {
  code: string | null;
  location: string;
  message: string;
  waitedMs: number;
  earlierStepReachedExcel: boolean;
  outcomeUnknown: boolean;
}

export function parseExcelError(error: unknown): ExcelErrorInfo {
  const message = getErrorMessage(error);
  if (!Value.Check(officeErrorSchema, error)) {
    return { code: null, location: "", message, waitedMs: 0, earlierStepReachedExcel: false, outcomeUnknown: false };
  }
  return {
    code: error.code,
    location: error.debugInfo?.errorLocation ?? "",
    message,
    waitedMs: error.waitedMs ?? 0,
    earlierStepReachedExcel: error.earlierStepReachedExcel ?? false,
    outcomeUnknown: error.outcomeUnknown ?? false,
  };
}

async function describeProtection(sheetName: string | undefined): Promise<string | null> {
  try {
    return await excelRun(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      sheet.protection.load("protected");
      context.workbook.protection.load("protected");
      const sheets = sheetName ? null : context.workbook.worksheets;
      sheets?.load("items/name,items/protection/protected");
      await context.sync();
      const protectedOthers = sheets && !sheet.protection.protected
        ? sheets.items.filter((item) => item.protection.protected).map((item) => `"${item.name}"`)
        : [];
      if (protectedOthers.length > 0) {
        return `Protected sheet(s): ${protectedOthers.join(", ")}. Excel refused the change, most likely because it touches one of them. ` +
          "Pi cannot unprotect a sheet: ask the user to unprotect it (Review → Unprotect Sheet) " +
          "or unlock the target cells, then retry.";
      }
      if (sheet.protection.protected) {
        return `Sheet "${sheet.name}" is protected, so Excel refused the change. ` +
          "Pi cannot unprotect it: ask the user to unprotect the sheet (Review → Unprotect Sheet) " +
          "or unlock the target cells, then retry.";
      }
      if (context.workbook.protection.protected) {
        return "The workbook structure is protected, so sheets cannot be added, deleted, renamed, moved or hidden. " +
          "Ask the user to unprotect the workbook (Review → Protect Workbook), then retry.";
      }
      return null;
    });
  } catch {
    return null;
  }
}

export async function describeExcelFailure(
  error: ExcelErrorInfo,
  target: ExcelFailureTarget = {},
  nothingSent = false,
): Promise<ExcelFailure> {
  const { code, location, message } = error;
  const sheetName = target.sheet || (target.range?.includes("!") ? parseRangeRef(target.range).sheet : undefined);

  if (code === "AccessDenied") {
    const protection = await describeProtection(sheetName);
    return {
      code,
      reason: protection ??
        "Excel denied the change (AccessDenied). The workbook may be read-only, open in view mode, or protected.",
      maybeApplied: false,
    };
  }

  if (error.outcomeUnknown && nothingSent) {
    return {
      code,
      reason: "Excel did not answer before the change was sent, so nothing was written. " +
        "Retry once Excel responds; if it keeps happening, ask the user to reload the Excel page.",
      maybeApplied: false,
    };
  }

  if (error.outcomeUnknown) {
    return {
      code,
      reason: `Excel did not answer (${message}), so whether the change was applied is unknown. ` +
        "Read the affected range before retrying.",
      maybeApplied: true,
    };
  }

  if (code === CELL_EDIT_MODE_CODE || code === CELL_EDIT_MODE_WAITED_CODE) {
    return {
      code,
      reason: cellEditModeReason(error.waitedMs),
      maybeApplied: error.earlierStepReachedExcel,
    };
  }

  if (code !== null && PAYLOAD_LIMIT_CODES.has(code)) {
    return {
      code,
      reason: "The request was larger than Excel accepts in one call (payload size limit). " +
        "Split it into smaller ranges; for a formula column use fill_formula instead of a large values array.",
      maybeApplied: false,
    };
  }

  if (code === "ItemNotFound" && location.startsWith("WorksheetCollection")) {
    let names: string[] | null = null;
    try {
      names = await excelRun(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        return sheets.items.map((sheet) => sheet.name);
      });
    } catch {
    }
    const sheetLabel = sheetName ? `Sheet "${sheetName}"` : "The sheet";
    return {
      code,
      reason: `${sheetLabel} does not exist.` +
        (names ? ` Sheets in this workbook: ${names.map((name) => `"${name}"`).join(", ")}.` : ""),
      maybeApplied: false,
    };
  }

  if (code === "InvalidArgument" && location.endsWith("getRange")) {
    return {
      code,
      reason: `${target.range ? `"${target.range}"` : "The range"} is not a valid A1 address on this sheet.`,
      maybeApplied: false,
    };
  }

  return {
    code,
    reason: code ? `Excel error ${code}: ${message}` : message,
    maybeApplied: false,
  };
}
