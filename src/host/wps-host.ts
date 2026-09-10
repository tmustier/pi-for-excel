/** WPS Spreadsheets host implementation. */

import {
  getWpsActiveWorkbook,
  getWpsEtApplication,
  getWpsWorkbookFullName,
  getWpsWorkbookName,
} from "./wps/jsapi.js";
import { settingsBackedSessionStorage } from "./session-storage.js";
import {
  createUnknownWorkbookContext,
  getWorkbookContextFromWpsFullName,
  type WorkbookContext,
} from "./workbook-context.js";
import type {
  SpreadsheetHost,
  SpreadsheetHostReadyCallback,
  SpreadsheetHostReadyInfo,
  SpreadsheetHostSessionStorage,
} from "./types.js";

export const WPS_UNSUPPORTED_PHASE_1_MESSAGE =
  "WPS Spreadsheets Phase 2 (NEXSELL-370) supports only workbook overview, read_range, write_cells, and execute_wps_js so far; this operation is not yet supported on WPS.";

export class WpsHost implements SpreadsheetHost {
  readonly kind = "wps";
  readonly displayName = "WPS Spreadsheets";
  readonly sessionStorage: SpreadsheetHostSessionStorage = settingsBackedSessionStorage;

  whenReady(): Promise<SpreadsheetHostReadyInfo> {
    return Promise.resolve({
      kind: "wps",
      nativeHost: "wps",
      nativePlatform: null,
      reason: "wps-jsapi",
    });
  }

  onReady(callback: SpreadsheetHostReadyCallback): () => void {
    let disposed = false;
    queueMicrotask(() => {
      if (disposed) return;
      callback({
        kind: "wps",
        nativeHost: "wps",
        nativePlatform: null,
        reason: "wps-jsapi",
      });
    });

    return () => {
      disposed = true;
    };
  }

  getWorkbookContext(): Promise<WorkbookContext> {
    const app = getWpsEtApplication();
    const workbook = app ? getWpsActiveWorkbook(app) : null;
    if (!workbook) {
      return Promise.resolve(createUnknownWorkbookContext());
    }

    const rawFullName = getWpsWorkbookFullName(workbook);
    const rawWorkbookName = getWpsWorkbookName(workbook);
    const fullName = typeof rawFullName === "string" ? rawFullName : null;
    const workbookName = typeof rawWorkbookName === "string" ? rawWorkbookName : null;

    return getWorkbookContextFromWpsFullName(fullName, workbookName);
  }

  resolveThemeDark(): boolean | null {
    // TODO(NEXSELL-370): map WPS theme APIs if available. Unknown keeps the
    // existing media-query fallback in the taskpane UI.
    return null;
  }
}
