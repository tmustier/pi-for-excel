import { getWorkbookContext } from "../workbook/context.js";
import type { SpreadsheetHost, SpreadsheetHostKind, SpreadsheetHostReadyInfo } from "./types.js";

function unknownWorkbookContext() {
  return {
    workbookId: null,
    workbookName: null,
    source: "unknown" as const,
  };
}

function hasOfficeHost(): boolean {
  return typeof Office !== "undefined";
}

function hasWpsHost(): boolean {
  const maybeWindow = typeof window !== "undefined" ? window as unknown as { wps?: unknown; WPS?: unknown } : undefined;
  return Boolean(maybeWindow?.wps ?? maybeWindow?.WPS);
}

export function detectSpreadsheetHostKind(): SpreadsheetHostKind {
  if (hasOfficeHost()) return "office";
  if (hasWpsHost()) return "wps";
  return "browser";
}

export function createSpreadsheetHost(kind: SpreadsheetHostKind): SpreadsheetHost {
  if (kind === "office") {
    return {
      kind,
      ready: async (): Promise<SpreadsheetHostReadyInfo> => {
        if (typeof Office === "undefined") {
          return { kind: "browser" };
        }
        const info = await Office.onReady();
        return {
          kind,
          host: typeof info.host === "string" ? info.host : undefined,
          platform: typeof info.platform === "string" ? info.platform : undefined,
        };
      },
      getWorkbookContext,
    };
  }

  if (kind === "wps") {
    return {
      kind,
      ready: () => Promise.resolve({ kind }),
      getWorkbookContext: () => Promise.resolve(unknownWorkbookContext()),
    };
  }

  return {
    kind: "browser",
    ready: () => Promise.resolve({ kind: "browser" }),
    getWorkbookContext: () => Promise.resolve(unknownWorkbookContext()),
  };
}

export type { SpreadsheetHost, SpreadsheetHostKind, SpreadsheetHostReadyInfo } from "./types.js";
