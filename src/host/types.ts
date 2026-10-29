import type { SettingsStore } from "../storage/local/settings-store.js";

import type { SessionWorkbookPartition } from "../workbook/session-association.js";
import type { DocumentInstanceIdentity } from "./document-instance.js";
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
  readonly servesOneDocument: boolean;
  readonly sessionStorage: SpreadsheetHostSessionStorage;

  whenReady(): Promise<SpreadsheetHostReadyInfo>;

  onReady(callback: SpreadsheetHostReadyCallback): () => void;

  getWorkbookContext(): Promise<WorkbookContext>;

  readonly documentInstance: DocumentInstanceIdentity | null;

  resolveThemeDark(): boolean | null;
}
