/** Office.js-backed host implementation. */

import type { DocumentInstanceIdentity } from "./document-instance.js";
import { createOfficeDocumentInstanceIdentity } from "./office-document-instance.js";
import { createOfficeDocumentUrlReader } from "./office-document-url.js";
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
  readonly documentInstance: DocumentInstanceIdentity = createOfficeDocumentInstanceIdentity();
  private readonly documentUrl = createOfficeDocumentUrlReader();

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

  async getWorkbookContext(): Promise<WorkbookContext> {
    return getWorkbookContextFromDocumentUrl(await this.documentUrl.read());
  }

  resolveThemeDark(): boolean | null {
    return resolveOfficeThemeDark();
  }
}
