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
});

export interface ExcelFailure {
  /** Office.js error code (`AccessDenied`, `ItemNotFound`, ...) when the host gave one. */
  code: string | null;
  /** One or two sentences naming the cause and what to do next. */
  reason: string;
}

export interface ExcelFailureTarget {
  /** Sheet name the call targeted, when known. */
  sheet?: string;
  /** Range or cell reference as the agent wrote it (may include the sheet). */
  range?: string;
}

const PAYLOAD_LIMIT_CODES = new Set([
  "ResponsePayloadSizeLimitExceeded",
  "RequestPayloadSizeLimitExceeded",
  "PayloadTooLarge",
]);

/** A caught Office.js failure, parsed at the catch seam. */
export interface ExcelErrorInfo {
  code: string | null;
  /** `debugInfo.errorLocation`, e.g. "WorksheetCollection.getItem". */
  location: string;
  message: string;
}

/** Parse a caught value (the `catch` seam) into the fields Office.js errors carry. */
export function parseExcelError(error: unknown): ExcelErrorInfo {
  const message = getErrorMessage(error);
  if (!Value.Check(officeErrorSchema, error)) return { code: null, location: "", message };
  return { code: error.code, location: error.debugInfo?.errorLocation ?? "", message };
}

function targetSheetName(target: ExcelFailureTarget): string | undefined {
  if (target.sheet) return target.sheet;
  if (target.range?.includes("!")) return parseRangeRef(target.range).sheet;
  return undefined;
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
      await context.sync();
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

async function listSheetNames(): Promise<string[] | null> {
  try {
    return await excelRun(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      return sheets.items.map((sheet) => sheet.name);
    });
  } catch {
    return null;
  }
}

/** Describe why an Excel call failed. Never throws; probes are best-effort. */
export async function describeExcelFailure(
  error: ExcelErrorInfo,
  target: ExcelFailureTarget = {},
): Promise<ExcelFailure> {
  const { code, location, message } = error;
  const sheetName = targetSheetName(target);

  if (code === "AccessDenied") {
    const protection = await describeProtection(sheetName);
    return {
      code,
      reason: protection ??
        "Excel denied the change (AccessDenied). The workbook may be read-only, open in view mode, or protected.",
    };
  }

  if (code === "InvalidOperationInCellEditMode") {
    return {
      code,
      reason: "Excel is in cell-edit mode (someone is typing in a cell) and refuses every change until that ends. " +
        "Ask the user to press Enter or Esc in the sheet, then retry.",
    };
  }

  if (code !== null && PAYLOAD_LIMIT_CODES.has(code)) {
    return {
      code,
      reason: "The request was larger than Excel accepts in one call (payload size limit). " +
        "Split it into smaller ranges; for a formula column use fill_formula instead of a large values array.",
    };
  }

  if (code === "ItemNotFound" && location.startsWith("WorksheetCollection")) {
    const names = await listSheetNames();
    const sheetLabel = sheetName ? `Sheet "${sheetName}"` : "The sheet";
    return {
      code,
      reason: `${sheetLabel} does not exist.` +
        (names ? ` Sheets in this workbook: ${names.map((name) => `"${name}"`).join(", ")}.` : ""),
    };
  }

  if (code === "InvalidArgument" && location.endsWith("getRange")) {
    return {
      code,
      reason: `${target.range ? `"${target.range}"` : "The range"} is not a valid A1 address on this sheet.`,
    };
  }

  return {
    code,
    reason: code ? `Excel error ${code}: ${message}` : message,
  };
}
