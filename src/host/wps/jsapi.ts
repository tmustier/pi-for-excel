function isHostWpsJsapiPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Typed, lazy WPS ET (Spreadsheets) JSAPI facade.
 *
 * WPS exposes a synchronous VBA-like object model through `window.wps` and/or a
 * global `Application` object. Keep all direct WPS global access in this file so
 * browser/node tests can install a fake graph without importing WPS at module
 * load time.
 */


export interface WpsCountedCollection {
  Count?: DynamicValue;
  count?: DynamicValue;
  Item?: (key: string | number) => DynamicValue;
}

export interface WpsRowsOrColumns {
  Count?: DynamicValue;
  count?: DynamicValue;
}

export interface WpsEtRange {
  Address?: DynamicValue;
  Value2?: DynamicValue;
  Value?: (rangeValueDataType?: DynamicValue, value?: DynamicValue) => DynamicValue;
  Formula?: DynamicValue;
  NumberFormat?: DynamicValue;
  Rows?: WpsRowsOrColumns;
  Columns?: WpsRowsOrColumns;
  Row?: DynamicValue;
  Column?: DynamicValue;
  MergeCells?: DynamicValue;
}

export interface WpsEtWorksheet {
  Name?: DynamicValue;
  name?: DynamicValue;
  Visible?: DynamicValue;
  visible?: DynamicValue;
  UsedRange?: WpsEtRange | null;
  Range?: (address: string) => WpsEtRange;
}

export interface WpsEtWorkbook {
  Name?: DynamicValue;
  name?: DynamicValue;
  FullName?: DynamicValue;
  fullName?: DynamicValue;
  Sheets?: WpsCountedCollection;
  Worksheets?: WpsCountedCollection;
}

export interface WpsPluginStorage {
  length?: DynamicValue;
  getItem?: (key: string) => DynamicValue;
  setItem?: (key: string, value: DynamicValue) => void;
  removeItem?: (key: string) => void;
  clear?: () => void;
  key?: (index: number) => string | null;
  Key?: (index: number) => string | null;
}

export interface WpsTaskPane {
  ID?: DynamicValue;
  Visible?: DynamicValue;
  Width?: DynamicValue;
  Height?: DynamicValue;
  DockPosition?: DynamicValue;
  Navigate?: (url: string) => DynamicValue;
  Delete?: () => DynamicValue;
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
  EtApplication?: () => DynamicValue;
  CreateTaskPane?: (url: string) => WpsTaskPane;
  CreateTaskpane?: (url: string) => WpsTaskPane;
  PluginStorage?: WpsPluginStorage;
}

function asWpsEtApplication(value: DynamicValue): WpsEtApplication | null {
  return isHostWpsJsapiPayloadShape(value) ? value : null;
}

function getGlobalMember(key: string): DynamicValue {
  return Reflect.get(globalThis, key);
}

function getWpsGlobal(): WpsGlobal | null {
  const candidate = getGlobalMember("wps");
  return isHostWpsJsapiPayloadShape(candidate) ? candidate : null;
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
