function isHostOfficeHostPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Office.js-backed host implementation. */

import { settingsBackedSessionStorage } from "./session-storage.js";
import { resolveOfficeThemeDark } from "./office-theme.js";
import {
  getWorkbookContextFromDocumentUrl,
  type WorkbookContext,
} from "./workbook-context.js";
import type {
  SpreadsheetHost,
  SpreadsheetHostReadyCallback,
  SpreadsheetHostReadyInfo,
  SpreadsheetHostSessionStorage,
} from "./types.js";

function getOfficeDocumentUrl(): string | null {
  try {
    const office = Reflect.get(globalThis, "Office");
    if (!isHostOfficeHostPayloadShape(office)) return null;

    const ctx = office.context;
    if (!isHostOfficeHostPayloadShape(ctx)) return null;

    const doc = ctx.document;
    if (!isHostOfficeHostPayloadShape(doc)) return null;

    const url = doc.url;
    return typeof url === "string" && url.trim().length > 0 ? url : null;
  } catch {
    return null;
  }
}

function nativeValueToString(value: DynamicValue): string | null {
  if (typeof value === "string" && value.trim().length > 0) {
    return value;
  }

  return null;
}

function toError(error: DynamicValue): Error {
  return error instanceof Error ? error : new Error("Office.onReady failed.");
}

interface OfficeReadyInfoLike {
  host?: DynamicValue;
  platform?: DynamicValue;
}

function fromOfficeReadyInfo(info: OfficeReadyInfoLike): SpreadsheetHostReadyInfo {
  return {
    kind: "office",
    nativeHost: nativeValueToString(info.host),
    nativePlatform: nativeValueToString(info.platform),
    reason: "office-ready",
  };
}

export class OfficeHost implements SpreadsheetHost {
  readonly kind = "office";
  readonly displayName = "Microsoft Excel";
  readonly sessionStorage: SpreadsheetHostSessionStorage = settingsBackedSessionStorage;

  whenReady(): Promise<SpreadsheetHostReadyInfo> {
    if (typeof Office === "undefined") {
      return Promise.reject(new Error("Office.js is unavailable."));
    }

    return new Promise((resolve, reject) => {
      try {
        const readyPromise = Office.onReady((info) => {
          resolve(fromOfficeReadyInfo(info));
        });

        void readyPromise.catch((error: DynamicValue) => {
          reject(error instanceof Error ? error : new Error("Office.onReady failed."));
        });
      } catch (error) {
        reject(toError(error));
      }
    });
  }

  onReady(callback: SpreadsheetHostReadyCallback): () => void {
    if (typeof Office === "undefined") {
      return () => {};
    }

    let disposed = false;
    try {
      const readyPromise = Office.onReady((info) => {
        if (disposed) return;
        callback(fromOfficeReadyInfo(info));
      });

      void readyPromise.catch((error: DynamicValue) => {
        if (!disposed) {
          console.warn("[pi] Office.onReady hook failed:", error);
        }
      });
    } catch (error) {
      console.warn("[pi] Office.onReady hook failed:", error);
    }

    return () => {
      disposed = true;
    };
  }

  getWorkbookContext(): Promise<WorkbookContext> {
    return getWorkbookContextFromDocumentUrl(getOfficeDocumentUrl());
  }

  resolveThemeDark(): boolean | null {
    return resolveOfficeThemeDark();
  }
}
