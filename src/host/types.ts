/** Host abstraction types for Office, WPS, and browser/UI-test runtimes. */

import type { SettingsStore } from "@earendil-works/pi-web-ui";

import type { SessionWorkbookPartition } from "../workbook/session-association.js";
import type { WorkbookContext } from "./workbook-context.js";

export type SpreadsheetHostKind = "office" | "wps" | "browser";

export interface SpreadsheetHostReadyInfo {
  kind: SpreadsheetHostKind;
  nativeHost: string | null;
  nativePlatform: string | null;
  reason: "office-ready" | "office-unavailable" | "office-timeout" | "wps-jsapi" | "browser";
}

export type SpreadsheetHostReadyCallback = (info: SpreadsheetHostReadyInfo) => void;

export interface SpreadsheetHostSessionStorage {
  getSessionWorkbookId(settings: SettingsStore, sessionId: string): Promise<string | null>;
  linkSessionToWorkbook(settings: SettingsStore, sessionId: string, workbookId: string): Promise<void>;
  setLatestSessionForWorkbook(settings: SettingsStore, workbookId: string, sessionId: string): Promise<void>;
  getLatestSessionForWorkbook(settings: SettingsStore, workbookId: string): Promise<string | null>;
  partitionSessionIdsByWorkbook(
    settings: SettingsStore,
    sessionIds: string[],
    workbookId: string,
  ): Promise<SessionWorkbookPartition>;
}

export interface SpreadsheetHost {
  readonly kind: SpreadsheetHostKind;
  readonly displayName: string;
  readonly sessionStorage: SpreadsheetHostSessionStorage;

  /** Resolve when the host is ready enough for the taskpane to initialize. */
  whenReady(): Promise<SpreadsheetHostReadyInfo>;

  /** Subscribe to a host-ready signal where the host supports one. */
  onReady(callback: SpreadsheetHostReadyCallback): () => void;

  /** Privacy-preserving workbook identity. Never expose or persist raw document URLs. */
  getWorkbookContext(): Promise<WorkbookContext>;

  /** Host theme signal, if available. Returns null when unknown/unsupported. */
  resolveThemeDark(): boolean | null;
}
