/**
 * Manual full-workbook backup flow (#191).
 *
 * This module intentionally stays separate from automatic range-level recovery
 * checkpoints. It captures the current workbook as a compressed Office file
 * only when the user explicitly requests a manual backup.
 */

import { bytesToBase64 } from "../files/encoding.js";
import { getFilesWorkspace, type PendingWorkspaceDownload } from "../files/workspace.js";
import type { WorkspaceFileEntry } from "../files/types.js";
import {
  getOfficeDocumentFileAdapter,
  type OfficeAsyncResultLike,
  type OfficeFileLike,
  type OfficeSliceLike,
} from "../host/office-file.js";
import { formatWorkbookLabel, getWorkbookContext, type WorkbookContext } from "./context.js";

const FULL_BACKUP_PATH_PREFIX = "manual-backups/full-workbook/v1";
const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DEFAULT_SLICE_SIZE_BYTES = 1_048_576; // 1 MB
const DEFAULT_LIST_LIMIT = 20;

interface ManualFullBackupWorkspace {
  listFiles: () => Promise<WorkspaceFileEntry[]>;
  writeBase64File: (path: string, base64: string, mimeTypeHint?: string) => Promise<void>;
  /** Must be called before any `await`, while the user's click still counts. */
  beginDownload: () => PendingWorkspaceDownload;
  deleteFile: (path: string) => Promise<void>;
}

interface ManualFullBackupDependencies {
  getWorkspace: () => ManualFullBackupWorkspace;
  getWorkbookContext: () => Promise<WorkbookContext>;
  captureWorkbookBytes: () => Promise<Uint8Array>;
  now: () => number;
  createSuffix: () => string;
}

export interface ManualFullWorkbookBackup {
  id: string;
  path: string;
  createdAt: number;
  sizeBytes: number;
}

function defaultNow(): number {
  return Date.now();
}

function defaultCreateSuffix(): string {
  const randomUuid = globalThis.crypto?.randomUUID;
  if (typeof randomUuid === "function") {
    return randomUuid.call(globalThis.crypto).slice(0, 8).toLowerCase();
  }

  return Math.floor(Math.random() * 1_000_000)
    .toString(36)
    .padStart(6, "0")
    .slice(0, 8);
}

function toError(error: unknown): Error {
  if (error instanceof Error) {
    return error;
  }

  return new Error(typeof error === "string" ? error : "Unknown error");
}

function formatOfficeAsyncError(result: OfficeAsyncResultLike<unknown>): string {
  const maybeMessage = result.error?.message;
  if (typeof maybeMessage === "string" && maybeMessage.trim().length > 0) {
    return maybeMessage.trim();
  }

  return "Office API call failed.";
}

function assertSucceeded<T>(result: OfficeAsyncResultLike<T>): T {
  const status = typeof result.status === "string" ? result.status.toLowerCase() : "";
  if (status !== "succeeded") {
    throw new Error(formatOfficeAsyncError(result));
  }

  if (result.value === undefined || result.value === null) {
    throw new Error("Office API returned an empty result.");
  }

  return result.value;
}

async function openCompressedWorkbookFile(sliceSize: number): Promise<OfficeFileLike> {
  const document = getOfficeDocumentFileAdapter();
  if (!document) {
    throw new Error(
      "Manual full-workbook backup is unavailable: Office document file API is not available.",
    );
  }

  return new Promise<OfficeFileLike>((resolve, reject) => {
    try {
      document.getFileAsync("compressed", { sliceSize }, (result) => {
        try {
          resolve(assertSucceeded(result));
        } catch (error) {
          reject(toError(error));
        }
      });
    } catch (error) {
      reject(toError(error));
    }
  });
}

async function closeOfficeFile(file: OfficeFileLike): Promise<void> {
  await new Promise<void>((resolve) => {
    try {
      file.closeAsync(() => resolve());
    } catch {
      resolve();
    }
  });
}

async function readSlice(file: OfficeFileLike, index: number): Promise<OfficeSliceLike> {
  return new Promise<OfficeSliceLike>((resolve, reject) => {
    try {
      file.getSliceAsync(index, (result) => {
        try {
          resolve(assertSucceeded(result));
        } catch (error) {
          reject(toError(error));
        }
      });
    } catch (error) {
      reject(toError(error));
    }
  });
}

function normalizeSliceData(data: unknown): Uint8Array {
  if (data instanceof Uint8Array) {
    return data;
  }

  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }

  if (ArrayBuffer.isView(data)) {
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  }

  if (Array.isArray(data)) {
    const items = Array.from<unknown>(data);
    const bytes = new Uint8Array(items.length);

    for (let index = 0; index < items.length; index += 1) {
      const value = items[index];
      if (typeof value !== "number" || !Number.isInteger(value) || value < 0 || value > 255) {
        throw new Error("Workbook backup slice contained invalid byte data.");
      }
      bytes[index] = value;
    }

    return bytes;
  }

  throw new Error("Workbook backup slice data has an unsupported format.");
}

function combineChunks(chunks: Uint8Array[]): Uint8Array {
  const totalBytes = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const out = new Uint8Array(totalBytes);

  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.byteLength;
  }

  return out;
}

export async function captureWorkbookCompressedBytes(
  sliceSize = DEFAULT_SLICE_SIZE_BYTES,
): Promise<Uint8Array> {
  const file = await openCompressedWorkbookFile(sliceSize);

  try {
    const chunks: Uint8Array[] = [];
    for (let index = 0; index < file.sliceCount; index += 1) {
      const slice = await readSlice(file, index);
      chunks.push(normalizeSliceData(slice.data));
    }

    return combineChunks(chunks);
  } finally {
    await closeOfficeFile(file);
  }
}

function sanitizePathSegment(value: string): string {
  const normalized = value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "");

  if (normalized.length === 0) {
    return "workbook";
  }

  return normalized.slice(0, 80);
}

function createBackupId(args: {
  workbookName: string | null;
  at: number;
  suffix: string;
}): string {
  const timestamp = new Date(args.at).toISOString().replace(/[:.]/g, "-");
  const workbookSegment = sanitizePathSegment(args.workbookName ?? "workbook");
  const suffix = sanitizePathSegment(args.suffix).slice(0, 12);
  return `${timestamp}_${workbookSegment}_${suffix}`;
}

function toBackupPath(workbookId: string, backupId: string): string {
  return `${FULL_BACKUP_PATH_PREFIX}/${sanitizePathSegment(workbookId)}/${backupId}.xlsx`;
}

function parseBackupId(fileName: string): string | null {
  if (!fileName.toLowerCase().endsWith(".xlsx")) return null;

  const withoutExt = fileName.slice(0, -5).trim();
  if (withoutExt.length === 0) return null;

  return withoutExt;
}

function toBackupEntry(file: WorkspaceFileEntry): ManualFullWorkbookBackup | null {
  const id = parseBackupId(file.name);
  if (!id) return null;

  return {
    id,
    path: file.path,
    createdAt: file.modifiedAt,
    sizeBytes: file.size,
  };
}

function isManualBackupFile(file: WorkspaceFileEntry, workbookId: string): boolean {
  if (file.sourceKind !== "workspace") return false;
  if (!file.path.startsWith(`${FULL_BACKUP_PATH_PREFIX}/`)) return false;

  const taggedWorkbookId = file.workbookTag?.workbookId;
  return taggedWorkbookId === workbookId;
}

function assertWorkbookIdentity(context: WorkbookContext): {
  workbookId: string;
  workbookLabel: string;
  workbookName: string | null;
} {
  if (!context.workbookId) {
    throw new Error(
      "Current workbook identity is unavailable; cannot create or restore manual full-workbook backups safely.",
    );
  }

  return {
    workbookId: context.workbookId,
    workbookLabel: formatWorkbookLabel(context),
    workbookName: context.workbookName,
  };
}

export class ManualFullWorkbookBackupStore {
  private readonly dependencies: ManualFullBackupDependencies;

  constructor(dependencies: Partial<ManualFullBackupDependencies> = {}) {
    this.dependencies = {
      getWorkspace: dependencies.getWorkspace ?? (() => getFilesWorkspace()),
      getWorkbookContext: dependencies.getWorkbookContext ?? getWorkbookContext,
      captureWorkbookBytes: dependencies.captureWorkbookBytes ?? (() => captureWorkbookCompressedBytes()),
      now: dependencies.now ?? defaultNow,
      createSuffix: dependencies.createSuffix ?? defaultCreateSuffix,
    };
  }

  private async listForWorkbookId(
    workbookId: string,
    limit?: number,
  ): Promise<ManualFullWorkbookBackup[]> {
    const workspace = this.dependencies.getWorkspace();
    const files = await workspace.listFiles();

    const matching = files
      .filter((file) => isManualBackupFile(file, workbookId))
      .map((file) => toBackupEntry(file))
      .filter((entry): entry is ManualFullWorkbookBackup => entry !== null)
      .sort((left, right) => right.createdAt - left.createdAt);

    if (limit === undefined) {
      return matching;
    }

    return matching.slice(0, Math.max(0, limit));
  }

  async listForCurrentWorkbook(limit = DEFAULT_LIST_LIMIT): Promise<ManualFullWorkbookBackup[]> {
    const workbookContext = await this.dependencies.getWorkbookContext();
    const workbookId = workbookContext.workbookId;

    if (!workbookId) {
      return [];
    }

    return this.listForWorkbookId(workbookId, limit);
  }

  /**
   * Opens the download before anything slow happens (see `beginDownload`),
   * runs `produce`, and downloads the file it returns. Nothing produced closes
   * the window again.
   */
  private async withDownload(
    produce: () => Promise<ManualFullWorkbookBackup | null>,
  ): Promise<ManualFullWorkbookBackup | null> {
    const pending = this.dependencies.getWorkspace().beginDownload();
    try {
      const backup = await produce();
      if (!backup) {
        pending.cancel();
        return null;
      }
      await pending.download(backup.path);
      return backup;
    } catch (error) {
      pending.cancel();
      throw error;
    }
  }

  /**
   * Captures the workbook, stores the copy in the Files workspace, and
   * downloads it. `runCapture` lets the caller hold the workbook lock during
   * the capture so a copy never includes half of one of Pi's changes.
   */
  async createAndDownload(options: {
    runCapture?: (capture: () => Promise<ManualFullWorkbookBackup>) => Promise<ManualFullWorkbookBackup>;
  } = {}): Promise<ManualFullWorkbookBackup> {
    const runCapture = options.runCapture ?? ((capture) => capture());
    const created = await this.withDownload(() => runCapture(() => this.create()));
    if (!created) throw new Error("Manual backup was not created.");
    return created;
  }

  private async create(): Promise<ManualFullWorkbookBackup> {
    const workbookContext = await this.dependencies.getWorkbookContext();
    const scoped = assertWorkbookIdentity(workbookContext);

    const bytes = await this.dependencies.captureWorkbookBytes();
    const createdAt = this.dependencies.now();
    const backupId = createBackupId({
      workbookName: scoped.workbookName,
      at: createdAt,
      suffix: this.dependencies.createSuffix(),
    });

    const path = toBackupPath(scoped.workbookId, backupId);
    const workspace = this.dependencies.getWorkspace();

    await workspace.writeBase64File(path, bytesToBase64(bytes), XLSX_MIME_TYPE);

    return {
      id: backupId,
      path,
      createdAt,
      sizeBytes: bytes.byteLength,
    };
  }

  async downloadLatestForCurrentWorkbook(): Promise<ManualFullWorkbookBackup | null> {
    return this.withDownload(async () => {
      const workbookContext = await this.dependencies.getWorkbookContext();
      const scoped = assertWorkbookIdentity(workbookContext);

      const backups = await this.listForWorkbookId(scoped.workbookId, 1);
      return backups[0] ?? null;
    });
  }

  async downloadByIdForCurrentWorkbook(backupId: string): Promise<ManualFullWorkbookBackup | null> {
    const trimmedBackupId = backupId.trim().replace(/^#/, "");
    if (trimmedBackupId.length === 0) {
      return null;
    }

    return this.withDownload(async () => {
      const workbookContext = await this.dependencies.getWorkbookContext();
      const scoped = assertWorkbookIdentity(workbookContext);

      const backups = await this.listForWorkbookId(scoped.workbookId);
      return backups.find((backup) => backup.id === trimmedBackupId) ?? null;
    });
  }

  /**
   * Deletes every manual backup of the current workbook. Identity is checked
   * and the backups counted before `confirm` is asked, so the user is never
   * asked about a deletion that cannot happen; `confirm` is not asked when
   * there is nothing to delete. Returns the number deleted.
   */
  async clearForCurrentWorkbook(confirm: (count: number) => Promise<boolean> = () => Promise.resolve(true)): Promise<number> {
    const workbookContext = await this.dependencies.getWorkbookContext();
    const scoped = assertWorkbookIdentity(workbookContext);

    const backups = await this.listForWorkbookId(scoped.workbookId);
    if (backups.length === 0 || !await confirm(backups.length)) {
      return 0;
    }

    const workspace = this.dependencies.getWorkspace();
    for (const backup of backups) {
      await workspace.deleteFile(backup.path);
    }

    return backups.length;
  }
}

let singleton: ManualFullWorkbookBackupStore | null = null;

export function getManualFullWorkbookBackupStore(
  defaultStore?: ManualFullWorkbookBackupStore,
): ManualFullWorkbookBackupStore {
  if (!singleton) {
    singleton = defaultStore ?? new ManualFullWorkbookBackupStore();
  }

  return singleton;
}
