/**
 * Typed, lazy WPS ET (Spreadsheets) JSAPI adapter.
 *
 * WPS exposes a synchronous VBA-like object model through `window.wps` and/or a
 * global `Application` object. Product versions vary between properties and
 * zero-argument methods, workbook/application collections, and optional range
 * methods. Decode those differences here and keep calls bound to their owner.
 */

type WpsPropertyOrMethod<T> = T | (() => T);

export interface WpsCountedCollection {
  Count?: WpsPropertyOrMethod<unknown>;
  count?: WpsPropertyOrMethod<unknown>;
  Item?: (key: string | number) => unknown;
}

export interface WpsRowsOrColumns {
  Count?: WpsPropertyOrMethod<unknown>;
  count?: WpsPropertyOrMethod<unknown>;
}

export interface WpsEtRange {
  Address?: WpsPropertyOrMethod<unknown>;
  Value2?: WpsPropertyOrMethod<unknown>;
  Value?: (rangeValueDataType?: unknown, value?: unknown) => unknown;
  Formula?: WpsPropertyOrMethod<unknown>;
  NumberFormat?: WpsPropertyOrMethod<unknown>;
  Rows?: WpsPropertyOrMethod<WpsRowsOrColumns | null>;
  Columns?: WpsPropertyOrMethod<WpsRowsOrColumns | null>;
  Row?: WpsPropertyOrMethod<unknown>;
  Column?: WpsPropertyOrMethod<unknown>;
  MergeCells?: WpsPropertyOrMethod<unknown>;
}

export interface WpsEtWorksheet {
  Name?: WpsPropertyOrMethod<unknown>;
  name?: WpsPropertyOrMethod<unknown>;
  Visible?: WpsPropertyOrMethod<unknown>;
  visible?: WpsPropertyOrMethod<unknown>;
  UsedRange?: WpsPropertyOrMethod<WpsEtRange | null>;
  Range?: (address: string) => WpsEtRange;
}

export interface WpsEtWorkbook {
  Name?: WpsPropertyOrMethod<unknown>;
  name?: WpsPropertyOrMethod<unknown>;
  FullName?: WpsPropertyOrMethod<unknown>;
  fullName?: WpsPropertyOrMethod<unknown>;
  Sheets?: WpsPropertyOrMethod<WpsCountedCollection | null>;
  Worksheets?: WpsPropertyOrMethod<WpsCountedCollection | null>;
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
  ActiveWorkbook?: WpsPropertyOrMethod<WpsEtWorkbook | null>;
  ActiveSheet?: WpsPropertyOrMethod<WpsEtWorksheet | null>;
  Selection?: WpsPropertyOrMethod<WpsEtRange | null>;
  Sheets?: WpsPropertyOrMethod<WpsCountedCollection | null>;
  Worksheets?: WpsPropertyOrMethod<WpsCountedCollection | null>;
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

function readPropertyOrMethod<T>(owner: object, value: WpsPropertyOrMethod<T> | undefined): T | undefined {
  if (typeof value !== "function") return value;
  // The union's callable branch is established above; WPS methods require their owner receiver.
  const method = value as () => T;
  return method.call(owner);
}

function parseWpsEtApplication(value: unknown): WpsEtApplication | null {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return null;
  return value;
}

function parseWpsGlobal(value: unknown): WpsGlobal | null {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return null;
  return value;
}

function parseWpsRange(value: unknown): WpsEtRange | null {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return null;
  return value;
}

interface WpsHostGlobals {
  wps?: unknown;
  Application?: unknown;
}

function getWpsHostGlobals(): WpsHostGlobals {
  // WPS injects these non-standard globals before taskpane boot.
  return globalThis as typeof globalThis & WpsHostGlobals;
}

function getWpsGlobal(): WpsGlobal | null {
  return parseWpsGlobal(getWpsHostGlobals().wps);
}

/** Resolve the active WPS ET Application lazily at call time. */
export function getWpsEtApplication(): WpsEtApplication | null {
  const wpsGlobal = getWpsGlobal();
  if (typeof wpsGlobal?.EtApplication === "function") {
    try {
      const app = parseWpsEtApplication(wpsGlobal.EtApplication());
      if (app) return app;
    } catch {
      // Some WPS builds expose a throwing plugin accessor; use the global fallback.
    }
  }

  return parseWpsEtApplication(getWpsHostGlobals().Application);
}

export function getWpsActiveWorkbook(app: WpsEtApplication): WpsEtWorkbook | null {
  return readPropertyOrMethod(app, app.ActiveWorkbook) ?? null;
}

export function getWpsActiveWorksheet(app: WpsEtApplication): WpsEtWorksheet | null {
  return readPropertyOrMethod(app, app.ActiveSheet) ?? null;
}

export function getWpsSelection(app: WpsEtApplication): WpsEtRange | null {
  return readPropertyOrMethod(app, app.Selection) ?? null;
}

export function getWpsWorkbookSheetCollection(
  app: WpsEtApplication,
  workbook: WpsEtWorkbook,
): WpsCountedCollection | null {
  return readPropertyOrMethod(workbook, workbook.Worksheets)
    ?? readPropertyOrMethod(workbook, workbook.Sheets)
    ?? readPropertyOrMethod(app, app.Worksheets)
    ?? readPropertyOrMethod(app, app.Sheets)
    ?? null;
}

export function getWpsCollectionCount(
  collection: WpsCountedCollection | WpsRowsOrColumns | null | undefined,
): number | null {
  if (!collection) return null;
  const raw = readPropertyOrMethod(collection, collection.Count)
    ?? readPropertyOrMethod(collection, collection.count);
  return typeof raw === "number" && Number.isFinite(raw) && raw >= 0
    ? Math.floor(raw)
    : null;
}

export function getWpsCollectionItem(
  collection: WpsCountedCollection,
  key: string | number,
): unknown {
  if (typeof collection.Item !== "function") {
    throw new Error("WPS worksheet collection does not expose Item().");
  }
  return collection.Item(key);
}

export function getWpsWorksheetRange(
  sheet: WpsEtWorksheet,
  address: string,
): WpsEtRange | null {
  if (typeof sheet.Range !== "function") return null;
  return parseWpsRange(sheet.Range(address));
}

export function getWpsUsedRange(sheet: WpsEtWorksheet): WpsEtRange | null {
  return readPropertyOrMethod(sheet, sheet.UsedRange) ?? null;
}

export function getWpsRangeAddress(range: WpsEtRange): unknown {
  return readPropertyOrMethod(range, range.Address);
}

export function getWpsRangeRows(range: WpsEtRange): WpsRowsOrColumns | null {
  return readPropertyOrMethod(range, range.Rows) ?? null;
}

export function getWpsRangeColumns(range: WpsEtRange): WpsRowsOrColumns | null {
  return readPropertyOrMethod(range, range.Columns) ?? null;
}

export function getWpsRangeValues(range: WpsEtRange): unknown {
  const value2 = readPropertyOrMethod(range, range.Value2);
  if (value2 !== undefined) return value2;
  return typeof range.Value === "function" ? range.Value() : undefined;
}

export function getWpsRangeFormula(range: WpsEtRange): unknown {
  return readPropertyOrMethod(range, range.Formula);
}

export function getWpsRangeNumberFormat(range: WpsEtRange): unknown {
  return readPropertyOrMethod(range, range.NumberFormat);
}

export function getWpsWorksheetName(sheet: WpsEtWorksheet): unknown {
  return readPropertyOrMethod(sheet, sheet.Name) ?? readPropertyOrMethod(sheet, sheet.name);
}

export function getWpsWorksheetVisibility(sheet: WpsEtWorksheet): unknown {
  return readPropertyOrMethod(sheet, sheet.Visible) ?? readPropertyOrMethod(sheet, sheet.visible);
}

export function getWpsWorkbookName(workbook: WpsEtWorkbook): unknown {
  return readPropertyOrMethod(workbook, workbook.Name) ?? readPropertyOrMethod(workbook, workbook.name);
}

export function getWpsWorkbookFullName(workbook: WpsEtWorkbook): unknown {
  return readPropertyOrMethod(workbook, workbook.FullName)
    ?? readPropertyOrMethod(workbook, workbook.fullName);
}

export function writeWpsRangeValues(
  range: WpsEtRange,
  values: unknown[][],
  containsFormula: boolean,
): void {
  if (containsFormula) {
    range.Formula = values;
    return;
  }

  if (typeof range.Value === "function" && (range.Value2 === undefined || typeof range.Value2 === "function")) {
    range.Value(undefined, values);
    return;
  }

  range.Value2 = values;
}

export function getWpsGlobalForTaskPane(): WpsGlobal | null {
  return getWpsGlobal();
}
