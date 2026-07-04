/**
 * WPS Spreadsheets host scaffold.
 *
 * Phase 1 (NEXSELL-370) only detects WPS and keeps the app hostable. Real WPS
 * workbook identity, theme, and tool implementations are Phase 2 work.
 */

import { settingsBackedSessionStorage } from "./session-storage.js";
import {
  createUnknownWorkbookContext,
  type WorkbookContext,
} from "./workbook-context.js";
import type {
  SpreadsheetHost,
  SpreadsheetHostReadyCallback,
  SpreadsheetHostReadyInfo,
  SpreadsheetHostSessionStorage,
} from "./types.js";

export const WPS_UNSUPPORTED_PHASE_1_MESSAGE =
  "WPS Spreadsheets support is scaffolded only in Phase 1 (NEXSELL-370); this workbook operation is not yet supported on WPS.";

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
    // TODO(NEXSELL-370): derive a stable identity from WPS Application.ActiveWorkbook
    // without persisting raw local paths/URLs. Until then, WPS sessions remain
    // identity-less rather than leaking workbook locations into SettingsStore.
    return Promise.resolve(createUnknownWorkbookContext());
  }

  resolveThemeDark(): boolean | null {
    // TODO(NEXSELL-370): map WPS theme APIs if available. Unknown keeps the
    // existing media-query fallback in the taskpane UI.
    return null;
  }
}
