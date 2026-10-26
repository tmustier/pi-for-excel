function isAuditWorkbookChangeAuditPayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Workbook mutation audit log (local, persisted in SettingsStore when available).
 */

import { formatWorkbookLabel, type WorkbookContext, getWorkbookContext } from "../workbook/context.js";
import {
  EXECUTION_MODE_SETTING_KEY,
  normalizeExecutionMode,
  type ExecutionMode,
} from "../execution/mode.js";
import type { WorkbookCellChange } from "./cell-diff.js";
import type { SettingsUpdater } from "../storage/local/settings-store.js";

const AUDIT_SETTING_KEY = "workbook.change-audit.v1";
const MAX_AUDIT_ENTRIES = 500;

export type WorkbookAuditToolName =
  | "write_cells"
  | "fill_formula"
  | "python_transform_range"
  | "format_cells"
  | "conditional_format"
  | "modify_structure"
  | "comments"
  | "charts"
  | "view_settings"
  | "workbook_history"
  | "execute_office_js";

export interface WorkbookChangeAuditEntry {
  id: string;
  at: number;
  toolName: WorkbookAuditToolName;
  toolCallId: string;
  blocked: boolean;
  inputAddress?: string;
  outputAddress?: string;
  changedCount: number;
  changes: WorkbookCellChange[];
  summary?: string;
  executionMode?: ExecutionMode;
  workbookId?: string;
  workbookLabel?: string;
}

export interface AppendWorkbookChangeAuditEntryArgs {
  toolName: WorkbookAuditToolName;
  toolCallId: string;
  blocked: boolean;
  inputAddress?: string;
  outputAddress?: string;
  changedCount: number;
  changes: WorkbookCellChange[];
  summary?: string;
  executionMode?: ExecutionMode;
}

interface WorkbookChangeAuditLogDependencies {
  settings: SettingsUpdater | null;
  getSettingsStore: () => Promise<SettingsUpdater | null>;
  getWorkbookContext: () => Promise<WorkbookContext>;
  now: () => number;
  createId: () => string;
}

interface PersistedWorkbookChangeAuditPayload {
  version: 1;
  entries: WorkbookChangeAuditEntry[];
}

function defaultNow(): number {
  return Date.now();
}

function defaultCreateId(): string {
  const randomUuid = globalThis.crypto?.randomUUID;
  if (typeof randomUuid === "function") {
    return randomUuid.call(globalThis.crypto);
  }

  const randomChunk = Math.floor(Math.random() * 1_000_000)
    .toString(36)
    .padStart(4, "0");

  return `change_${Date.now().toString(36)}_${randomChunk}`;
}

async function defaultGetSettingsStore(): Promise<SettingsUpdater | null> {
  try {
    const storageModule = await import("../storage/local/app-storage.js");
    return storageModule.getAppStorage().settings;
  } catch {
    return null;
  }
}

function isWorkbookAuditToolName(value: unknown): value is WorkbookAuditToolName {
  return (
    value === "write_cells" ||
    value === "fill_formula" ||
    value === "python_transform_range" ||
    value === "format_cells" ||
    value === "conditional_format" ||
    value === "modify_structure" ||
    value === "comments" ||
    value === "charts" ||
    value === "view_settings" ||
    value === "workbook_history" ||
    value === "execute_office_js"
  );
}

function isWorkbookCellChange(value: unknown): value is WorkbookCellChange {
  if (!isAuditWorkbookChangeAuditPayloadShape(value)) return false;

  const beforeFormula = value.beforeFormula;
  const afterFormula = value.afterFormula;

  return (
    typeof value.address === "string" &&
    typeof value.beforeValue === "string" &&
    typeof value.afterValue === "string" &&
    (beforeFormula === undefined || typeof beforeFormula === "string") &&
    (afterFormula === undefined || typeof afterFormula === "string")
  );
}

function parseAuditEntry(value: unknown): WorkbookChangeAuditEntry | null {
  if (!isAuditWorkbookChangeAuditPayloadShape(value)) return null;

  if (!isWorkbookAuditToolName(value.toolName)) return null;
  if (typeof value.toolCallId !== "string") return null;
  if (typeof value.blocked !== "boolean") return null;
  if (typeof value.changedCount !== "number") return null;
  if (!Array.isArray(value.changes) || !value.changes.every((item) => isWorkbookCellChange(item))) return null;

  const id = typeof value.id === "string" ? value.id : defaultCreateId();
  const at = typeof value.at === "number" ? value.at : Date.now();

  const rawExecutionMode = value.executionMode;

  const entry: WorkbookChangeAuditEntry = {
    id,
    at,
    toolName: value.toolName,
    toolCallId: value.toolCallId,
    blocked: value.blocked,
    changedCount: value.changedCount,
    changes: value.changes,
  };

  if (typeof value.inputAddress === "string") entry.inputAddress = value.inputAddress;
  if (typeof value.outputAddress === "string") entry.outputAddress = value.outputAddress;
  if (typeof value.summary === "string") entry.summary = value.summary;
  if (typeof rawExecutionMode === "string") entry.executionMode = normalizeExecutionMode(rawExecutionMode);
  if (typeof value.workbookId === "string") entry.workbookId = value.workbookId;
  if (typeof value.workbookLabel === "string") entry.workbookLabel = value.workbookLabel;

  return entry;
}

function parsePersistedEntries(payload: unknown): WorkbookChangeAuditEntry[] {
  if (!isAuditWorkbookChangeAuditPayloadShape(payload)) return [];

  const entriesRaw = payload.entries;
  if (!Array.isArray(entriesRaw)) return [];

  const entries: WorkbookChangeAuditEntry[] = [];
  for (const item of entriesRaw) {
    const parsed = parseAuditEntry(item);
    if (parsed) {
      entries.push(parsed);
    }
  }

  return entries
    .sort((a, b) => b.at - a.at)
    .slice(0, MAX_AUDIT_ENTRIES);
}

function mergeEntries(
  pending: readonly WorkbookChangeAuditEntry[],
  stored: readonly WorkbookChangeAuditEntry[],
): WorkbookChangeAuditEntry[] {
  const storedIds = new Set(stored.map((entry) => entry.id));
  return [...pending.filter((entry) => !storedIds.has(entry.id)), ...stored]
    .sort((a, b) => b.at - a.at)
    .slice(0, MAX_AUDIT_ENTRIES);
}

function clampLimit(limit: number): number {
  if (!Number.isFinite(limit)) return 50;
  const rounded = Math.floor(limit);
  if (rounded <= 0) return 0;
  if (rounded > MAX_AUDIT_ENTRIES) return MAX_AUDIT_ENTRIES;
  return rounded;
}

export class WorkbookChangeAuditLog {
  private readonly dependencies: WorkbookChangeAuditLogDependencies;
  /** Entries whose write failed; retried with the next append. */
  private pendingEntries: WorkbookChangeAuditEntry[] = [];

  constructor(dependencies: Partial<WorkbookChangeAuditLogDependencies> = {}) {
    const getSettingsStore = dependencies.settings !== undefined
      ? () => Promise.resolve(dependencies.settings ?? null)
      : dependencies.getSettingsStore ?? defaultGetSettingsStore;

    this.dependencies = {
      settings: dependencies.settings ?? null,
      getSettingsStore,
      getWorkbookContext: dependencies.getWorkbookContext ?? getWorkbookContext,
      now: dependencies.now ?? defaultNow,
      createId: dependencies.createId ?? defaultCreateId,
    };
  }

  /**
   * Every taskpane on this computer appends to the same stored log, so each
   * write is an atomic update of what is stored now rather than a write-back
   * of a copy loaded earlier.
   */
  private async persistPending(): Promise<void> {
    const settings = await this.dependencies.getSettingsStore();
    if (!settings || this.pendingEntries.length === 0) return;

    const pending = this.pendingEntries;
    try {
      await settings.update(AUDIT_SETTING_KEY, parsePersistedEntries, (stored): PersistedWorkbookChangeAuditPayload => ({
        version: 1,
        entries: mergeEntries(pending, stored),
      }));
      this.pendingEntries = this.pendingEntries.filter((entry) => !pending.includes(entry));
    } catch (error) {
      console.warn("[pi] Workbook change audit entry could not be saved; retrying with the next change.", error);
    }
  }

  async append(args: AppendWorkbookChangeAuditEntryArgs): Promise<void> {
    let workbookId: string | undefined;
    let workbookLabel: string | undefined;

    try {
      const workbookContext = await this.dependencies.getWorkbookContext();
      if (workbookContext.workbookId) {
        workbookId = workbookContext.workbookId;
        workbookLabel = formatWorkbookLabel(workbookContext);
      }
    } catch {
      // ignore workbook context failures
    }

    let executionMode: ExecutionMode | undefined = args.executionMode;
    if (!executionMode) {
      try {
        const settings = await this.dependencies.getSettingsStore();
        if (settings) {
          const stored = await settings.get(EXECUTION_MODE_SETTING_KEY);
          executionMode = normalizeExecutionMode(stored);
        }
      } catch {
        // ignore execution-mode lookup failures
      }
    }

    const entry: WorkbookChangeAuditEntry = {
      id: this.dependencies.createId(),
      at: this.dependencies.now(),
      toolName: args.toolName,
      toolCallId: args.toolCallId,
      blocked: args.blocked,
      changedCount: args.changedCount,
      changes: args.changes,
      executionMode: executionMode ?? "yolo",
    };

    if (args.inputAddress !== undefined) entry.inputAddress = args.inputAddress;
    if (args.outputAddress !== undefined) entry.outputAddress = args.outputAddress;
    if (args.summary !== undefined) entry.summary = args.summary;
    if (workbookId !== undefined) entry.workbookId = workbookId;
    if (workbookLabel !== undefined) entry.workbookLabel = workbookLabel;

    this.pendingEntries = [entry, ...this.pendingEntries].slice(0, MAX_AUDIT_ENTRIES);
    await this.persistPending();
  }

  async list(limit = 50): Promise<WorkbookChangeAuditEntry[]> {
    const settings = await this.dependencies.getSettingsStore();
    const stored = settings ? parsePersistedEntries(await settings.get(AUDIT_SETTING_KEY)) : [];
    return mergeEntries(this.pendingEntries, stored).slice(0, clampLimit(limit));
  }

  async clear(): Promise<void> {
    this.pendingEntries = [];

    const settings = await this.dependencies.getSettingsStore();
    if (!settings) return;

    await settings.update(AUDIT_SETTING_KEY, parsePersistedEntries, () => undefined);
  }
}

let singleton: WorkbookChangeAuditLog | null = null;

export function getWorkbookChangeAuditLog(
  defaultLog?: WorkbookChangeAuditLog,
): WorkbookChangeAuditLog {
  if (!singleton) {
    singleton = defaultLog ?? new WorkbookChangeAuditLog();
  }

  return singleton;
}
