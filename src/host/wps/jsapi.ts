/**
 * Typed, lazy WPS ET (Spreadsheets) JSAPI facade.
 *
 * WPS exposes a synchronous VBA-like object model through `window.wps` and/or a
 * global `Application` object. Keep all direct WPS global access in this file so
 * browser/node tests can install a fake graph without importing WPS at module
 * load time.
 */

import { isRecord } from "../../utils/type-guards.js";

export interface WpsCountedCollection {
  Count?: unknown;
  count?: unknown;
  Item?: (key: string | number) => unknown;
}

export interface WpsRowsOrColumns {
  Count?: unknown;
  count?: unknown;
}

export interface WpsEtRange {
  Address?: unknown;
  Value2?: unknown;
  Value?: (rangeValueDataType?: unknown, value?: unknown) => unknown;
  Formula?: unknown;
  NumberFormat?: unknown;
  Rows?: WpsRowsOrColumns;
  Columns?: WpsRowsOrColumns;
  Row?: unknown;
  Column?: unknown;
  MergeCells?: unknown;
}

export interface WpsEtWorksheet {
  Name?: unknown;
  name?: unknown;
  Visible?: unknown;
  visible?: unknown;
  UsedRange?: WpsEtRange | null;
  Range?: (address: string) => WpsEtRange;
}

export interface WpsEtWorkbook {
  Name?: unknown;
  name?: unknown;
  FullName?: unknown;
  fullName?: unknown;
  Sheets?: WpsCountedCollection;
  Worksheets?: WpsCountedCollection;
}

export interface WpsPluginStorage {
  length?: unknown;
  getItem?: (key: string) => unknown;
  setItem?: (key: string, value: unknown) => void;
  removeItem?: (key: string) => void;
  clear?: () => void;
  key?: (index: number) => string | null;
  Key?: (index: number) => string | null;
}

export interface WpsTaskPane {
  ID?: unknown;
  Visible?: unknown;
  Width?: unknown;
  Height?: unknown;
  DockPosition?: unknown;
  Navigate?: (url: string) => unknown;
  Delete?: () => unknown;
}

export interface WpsEtApplication {
  ActiveWorkbook?: WpsEtWorkbook | null;
  ActiveSheet?: WpsEtWorksheet | null;
  Selection?: WpsEtRange | null;
  Sheets?: WpsCountedCollection;
  Worksheets?: WpsCountedCollection;
  PluginStorage?: WpsPluginStorage;
  Range?: (address: string) => WpsEtRange;
  CreateTaskpane?: (url: string) => WpsTaskPane;
  CreateTaskPane?: (url: string) => WpsTaskPane;
}

export interface WpsGlobal {
  EtApplication?: () => unknown;
  CreateTaskPane?: (url: string) => WpsTaskPane;
  CreateTaskpane?: (url: string) => WpsTaskPane;
  PluginStorage?: WpsPluginStorage;
}

function asWpsEtApplication(value: unknown): WpsEtApplication | null {
  return isRecord(value) ? value : null;
}

function getGlobalMember(key: string): unknown {
  return Reflect.get(globalThis, key);
}

function getWpsGlobal(): WpsGlobal | null {
  const candidate = getGlobalMember("wps");
  return isRecord(candidate) ? candidate : null;
}

/** Resolve the active WPS ET Application lazily at call time. */
export function getWpsEtApplication(): WpsEtApplication | null {
  const wps = getWpsGlobal();
  if (typeof wps?.EtApplication === "function") {
    try {
      const app = asWpsEtApplication(wps.EtApplication());
      if (app) return app;
    } catch {
      // Fall through to the global Application fallback.
    }
  }

  return asWpsEtApplication(getGlobalMember("Application"));
}

export function getWpsGlobalForTaskPane(): WpsGlobal | null {
  return getWpsGlobal();
}
