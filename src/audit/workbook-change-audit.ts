/**
 * Workbook mutation audit log (local, persisted in SettingsStore when available).
 */

import { formatWorkbookLabel, type WorkbookContext, getWorkbookContext } from "../workbook/context.js";
import {
  EXECUTION_MODE_SETTING_KEY,
  normalizeExecutionMode,
  type ExecutionMode,
} from "../execution/mode.js";
import { isRecord } from "../utils/type-guards.js";
import type { WorkbookCellChange } from "./cell-diff.js";

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

interface SettingsStoreLike {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: unknown): Promise<void>;
  delete(key: string): Promise<void>;
}

interface WorkbookChangeAuditLogDependencies {
  getSettingsStore: () => Promise<SettingsStoreLike | null>;
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

function isSettingsStoreLike(value: unknown): value is SettingsStoreLike {
  if (!isRecord(value)) return false;

  return (
    typeof value.get === "function" &&
    typeof value.set === "function" &&
    typeof value.delete === "function"
  );
}

async function defaultGetSettingsStore(): Promise<SettingsStoreLike | null> {
  try {
    const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
    const appStorage = storageModule.getAppStorage();
    const settings = isRecord(appStorage) ? appStorage.settings : null;
    return isSettingsStoreLike(settings) ? settings : null;
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
    value === "view_settings" ||
    value === "workbook_history" ||
    value === "execute_office_js"
  );
}

function isWorkbookCellChange(value: unknown): value is WorkbookCellChange {
  if (!isRecord(value)) return false;

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
  if (!isRecord(value)) return null;

  if (!isWorkbookAuditToolName(value.toolName)) return null;
  if (typeof value.toolCallId !== "string") return null;
  if (typeof value.blocked !== "boolean") return null;
  if (typeof value.changedCount !== "number") return null;
  if (!Array.isArray(value.changes) || !value.changes.every((item) => isWorkbookCellChange(item))) return null;

  const id = typeof value.id === "string" ? value.id : defaultCreateId();
  const at = typeof value.at === "number" ? value.at : Date.now();

  const rawExecutionMode = value.executionMode;

  return {
    id,
    at,
    toolName: value.toolName,
    toolCallId: value.toolCallId,
    blocked: value.blocked,
    inputAddress: typeof value.inputAddress === "string" ? value.inputAddress : undefined,
    outputAddress: typeof value.outputAddress === "string" ? value.outputAddress : undefined,
    changedCount: value.changedCount,
    changes: value.changes,
    summary: typeof value.summary === "string" ? value.summary : undefined,
    executionMode: typeof rawExecutionMode === "string" ? normalizeExecutionMode(rawExecutionMode) : undefined,
    workbookId: typeof value.workbookId === "string" ? value.workbookId : undefined,
    workbookLabel: typeof value.workbookLabel === "string" ? value.workbookLabel : undefined,
  };
}

function parsePersistedEntries(payload: unknown): WorkbookChangeAuditEntry[] {
  if (!isRecord(payload)) return [];

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

function clampLimit(limit: number): number {
  if (!Number.isFinite(limit)) return 50;
  const rounded = Math.floor(limit);
  if (rounded <= 0) return 0;
  if (rounded > MAX_AUDIT_ENTRIES) return MAX_AUDIT_ENTRIES;
  return rounded;
}

export class WorkbookChangeAuditLog {
  private readonly dependencies: WorkbookChangeAuditLogDependencies;
  private loaded = false;
  private entries: WorkbookChangeAuditEntry[] = [];

  constructor(dependencies: Partial<WorkbookChangeAuditLogDependencies> = {}) {
    this.dependencies = {
      getSettingsStore: dependencies.getSettingsStore ?? defaultGetSettingsStore,
      getWorkbookContext: dependencies.getWorkbookContext ?? getWorkbookContext,
      now: dependencies.now ?? defaultNow,
      createId: dependencies.createId ?? defaultCreateId,
    };
  }

  private async ensureLoaded(): Promise<void> {
    if (this.loaded) return;
    this.loaded = true;

    const settings = await this.dependencies.getSettingsStore();
    if (!settings) return;

    try {
      const payload = await settings.get<unknown>(AUDIT_SETTING_KEY);
      this.entries = parsePersistedEntries(payload);
    } catch {
      this.entries = [];
    }
  }

  private async persist(): Promise<void> {
    const settings = await this.dependencies.getSettingsStore();
    if (!settings) return;

    const payload: PersistedWorkbookChangeAuditPayload = {
      version: 1,
      entries: this.entries,
    };

    try {
      await settings.set(AUDIT_SETTING_KEY, payload);
    } catch {
      // ignore persistence failures
    }
  }

  async append(args: AppendWorkbookChangeAuditEntryArgs): Promise<void> {
    await this.ensureLoaded();

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
          const stored = await settings.get<unknown>(EXECUTION_MODE_SETTING_KEY);
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
      inputAddress: args.inputAddress,
      outputAddress: args.outputAddress,
      changedCount: args.changedCount,
      changes: args.changes,
      summary: args.summary,
      executionMode: executionMode ?? "yolo",
      workbookId,
      workbookLabel,
    };

    this.entries = [entry, ...this.entries].slice(0, MAX_AUDIT_ENTRIES);
    await this.persist();
  }

  async list(limit = 50): Promise<WorkbookChangeAuditEntry[]> {
    await this.ensureLoaded();
    return this.entries.slice(0, clampLimit(limit));
  }

  async clear(): Promise<void> {
    await this.ensureLoaded();
    this.entries = [];

    const settings = await this.dependencies.getSettingsStore();
    if (!settings) return;

    try {
      await settings.delete(AUDIT_SETTING_KEY);
    } catch {
      // ignore persistence failures
    }
  }
}

let singleton: WorkbookChangeAuditLog | null = null;

export function getWorkbookChangeAuditLog(): WorkbookChangeAuditLog {
  if (!singleton) {
    singleton = new WorkbookChangeAuditLog();
  }

  return singleton;
}
