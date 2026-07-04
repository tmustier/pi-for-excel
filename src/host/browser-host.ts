/** Browser/UI-gallery host used when no spreadsheet host is ready. */

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

export class BrowserHost implements SpreadsheetHost {
  readonly kind = "browser";
  readonly displayName = "Browser";
  readonly sessionStorage: SpreadsheetHostSessionStorage = settingsBackedSessionStorage;

  whenReady(): Promise<SpreadsheetHostReadyInfo> {
    return Promise.resolve({
      kind: "browser",
      nativeHost: null,
      nativePlatform: null,
      reason: "browser",
    });
  }

  onReady(callback: SpreadsheetHostReadyCallback): () => void {
    let disposed = false;
    queueMicrotask(() => {
      if (disposed) return;
      callback({
        kind: "browser",
        nativeHost: null,
        nativePlatform: null,
        reason: "browser",
      });
    });

    return () => {
      disposed = true;
    };
  }

  getWorkbookContext(): Promise<WorkbookContext> {
    return Promise.resolve(createUnknownWorkbookContext());
  }

  resolveThemeDark(): boolean | null {
    return null;
  }
}
