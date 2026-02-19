/**
 * Files workspace manager.
 *
 * Backend selection strategy:
 * 1) persisted native directory handle (when permission is still granted)
 * 2) OPFS
 * 3) in-memory fallback (non-browser/test environments)
 */

import { formatWorkbookLabel, getWorkbookContext } from "../workbook/context.js";
import { isRecord } from "../utils/type-guards.js";
import { base64ToBytes, bytesToBase64, encodeTextUtf8, truncateBase64, truncateText } from "./encoding.js";
import { MemoryBackend, NativeDirectoryBackend, OpfsBackend, type WorkspaceBackend } from "./backend.js";
import { getBuiltinWorkspaceDoc, isBuiltinWorkspacePath, listBuiltinWorkspaceDocs } from "./builtin-docs.js";
import { resolveSafeBlobUrlMimeType } from "./blob-url-safety.js";
import { inferMimeType, isTextMimeType } from "./mime.js";
import { getWorkspaceBaseName, normalizeWorkspacePath } from "./path.js";
import {
  FILES_WORKSPACE_CHANGED_EVENT,
  type FilesWorkspaceAuditAction,
  type FilesWorkspaceAuditActor,
  type FilesWorkspaceAuditEntry,
  type FilesWorkspaceChangedDetail,
  type WorkspaceBackendKind,
  type WorkspaceBackendStatus,
  type WorkspaceFileEntry,
  type WorkspaceFileLocationKind,
  type WorkspaceFileReadResult,
  type WorkspaceFileWorkbookTag,
  type WorkspaceSnapshot,
} from "./types.js";

const NATIVE_HANDLE_SETTING_KEY = "files.workspace.nativeHandle.v1";
const METADATA_SETTING_KEY = "files.workspace.metadata.v1";
const AUDIT_TRAIL_SETTING_KEY = "files.workspace.audit.v1";
const MAX_AUDIT_ENTRIES = 300;

const DEFAULT_UI_AUDIT_CONTEXT: FilesWorkspaceAuditContext = {
  actor: "user",
  source: "files-dialog",
};

export type WorkspaceReadMode = "auto" | "text" | "base64";

export interface FilesWorkspaceAuditContext {
  actor: FilesWorkspaceAuditActor;
  source: string;
}

export interface WorkspaceListOptions {
  audit?: FilesWorkspaceAuditContext;
}

export interface WorkspaceReadOptions {
  mode?: WorkspaceReadMode;
  maxChars?: number;
  audit?: FilesWorkspaceAuditContext;
  locationKind?: WorkspaceFileLocationKind;
}

export interface WorkspaceMutationOptions {
  audit?: FilesWorkspaceAuditContext;
  locationKind?: WorkspaceFileLocationKind;
}

interface DirectoryPickerHost {
  showDirectoryPicker: () => Promise<FileSystemDirectoryHandle>;
}

interface DirectoryPermissionHandle {
  queryPermission(descriptor?: FileSystemHandlePermissionDescriptor): Promise<PermissionState>;
  requestPermission(descriptor?: FileSystemHandlePermissionDescriptor): Promise<PermissionState>;
}

interface PersistedWorkspaceMetadata {
  version: 1;
  byPath: Record<string, WorkspaceFileWorkbookTag>;
}

interface PersistedAuditTrail {
  version: 1;
  entries: FilesWorkspaceAuditEntry[];
}

function isDirectoryPickerHost(value: unknown): value is DirectoryPickerHost {
  if (!isRecord(value)) return false;
  return typeof value.showDirectoryPicker === "function";
}

function isDirectoryHandle(value: unknown): value is FileSystemDirectoryHandle {
  if (!isRecord(value)) return false;
  return (
    value.kind === "directory" &&
    typeof value.getDirectoryHandle === "function" &&
    typeof value.getFileHandle === "function" &&
    typeof value.queryPermission === "function"
  );
}

function isDirectoryPermissionHandle(value: unknown): value is DirectoryPermissionHandle {
  if (!isRecord(value)) return false;

  return (
    typeof value.queryPermission === "function" &&
    typeof value.requestPermission === "function"
  );
}

function dispatchWorkspaceChanged(detail: FilesWorkspaceChangedDetail): void {
  if (typeof document === "undefined") return;
  document.dispatchEvent(new CustomEvent<FilesWorkspaceChangedDetail>(FILES_WORKSPACE_CHANGED_EVENT, { detail }));
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const out = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(out).set(bytes);
  return out;
}

function closeWindowSafely(windowHandle: Window | null): void {
  if (!windowHandle || windowHandle.closed) {
    return;
  }

  try {
    windowHandle.close();
  } catch {
    // ignore close errors
  }
}

function tryOpenDownloadWindow(url: string, pendingWindow: Window | null): boolean {
  if (pendingWindow && !pendingWindow.closed) {
    try {
      pendingWindow.location.replace(url);
      return true;
    } catch {
      closeWindowSafely(pendingWindow);
    }
  }

  return window.open(url, "_blank") !== null;
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function isMissingWorkspaceFileError(error: unknown): boolean {
  if (!(error instanceof Error)) return false;

  if (error.name === "NotFoundError") {
    return true;
  }

  const message = error.message.toLowerCase();

  if (message.includes("not found") || message.includes("not be found")) {
    return true;
  }

  if (message.includes("cannot find") || message.includes("can't find") || message.includes("can not find")) {
    return true;
  }

  if (message.includes("no such file") || message.includes("does not exist")) {
    return true;
  }

  return /\b(can\s?not|cannot|can't)\s+be\s+found\b/u.test(message);
}

function isWorkspaceBackendKind(value: unknown): value is WorkspaceBackendKind {
  return value === "native-directory" || value === "opfs" || value === "memory";
}

function isFilesWorkspaceAuditActor(value: unknown): value is FilesWorkspaceAuditActor {
  return value === "assistant" || value === "user" || value === "system";
}

function isFilesWorkspaceAuditAction(value: unknown): value is FilesWorkspaceAuditAction {
  return (
    value === "list" ||
    value === "read" ||
    value === "write" ||
    value === "delete" ||
    value === "rename" ||
    value === "import" ||
    value === "connect_native" ||
    value === "disconnect_native" ||
    value === "clear_audit"
  );
}

function sanitizeOptionalPath(value: unknown): string | undefined {
  if (!isNonEmptyString(value)) return undefined;

  try {
    return normalizeWorkspacePath(value);
  } catch {
    return undefined;
  }
}

function parseWorkbookTag(value: unknown): WorkspaceFileWorkbookTag | null {
  if (!isRecord(value)) return null;

  const workbookId = typeof value.workbookId === "string"
    ? value.workbookId.trim()
    : "";
  if (workbookId.length === 0) return null;

  const workbookLabel = typeof value.workbookLabel === "string"
    ? value.workbookLabel.trim()
    : "";
  if (workbookLabel.length === 0) return null;

  const taggedAt = isFiniteNumber(value.taggedAt)
    ? value.taggedAt
    : Date.now();

  return {
    workbookId,
    workbookLabel,
    taggedAt,
  };
}

function parsePersistedMetadata(value: unknown): Map<string, WorkspaceFileWorkbookTag> {
  const byPath = new Map<string, WorkspaceFileWorkbookTag>();
  if (!isRecord(value)) return byPath;

  const rawByPath = value.byPath;
  if (!isRecord(rawByPath)) return byPath;

  for (const [rawPath, rawTag] of Object.entries(rawByPath)) {
    const normalizedPath = sanitizeOptionalPath(rawPath);
    if (!normalizedPath) continue;

    const tag = parseWorkbookTag(rawTag);
    if (!tag) continue;

    byPath.set(normalizedPath, tag);
  }

  return byPath;
}

function parseAuditEntry(value: unknown): FilesWorkspaceAuditEntry | null {
  if (!isRecord(value)) return null;

  if (!isFilesWorkspaceAuditAction(value.action)) return null;
  if (!isFilesWorkspaceAuditActor(value.actor)) return null;
  if (!isWorkspaceBackendKind(value.backend)) return null;
  if (!isNonEmptyString(value.source)) return null;

  const at = isFiniteNumber(value.at) ? value.at : Date.now();
  const id = isNonEmptyString(value.id) ? value.id : createAuditEntryId();

  const path = sanitizeOptionalPath(value.path);
  const fromPath = sanitizeOptionalPath(value.fromPath);
  const toPath = sanitizeOptionalPath(value.toPath);

  const bytes = isFiniteNumber(value.bytes) ? value.bytes : undefined;
  const workbookId = isNonEmptyString(value.workbookId) ? value.workbookId.trim() : undefined;
  const workbookLabel = isNonEmptyString(value.workbookLabel) ? value.workbookLabel.trim() : undefined;

  return {
    id,
    at,
    action: value.action,
    actor: value.actor,
    source: value.source.trim(),
    backend: value.backend,
    path,
    fromPath,
    toPath,
    bytes,
    workbookId,
    workbookLabel,
  };
}

function parsePersistedAuditTrail(value: unknown): FilesWorkspaceAuditEntry[] {
  if (!isRecord(value)) return [];

  const entriesRaw = value.entries;
  if (!Array.isArray(entriesRaw)) return [];

  const parsedEntries: FilesWorkspaceAuditEntry[] = [];
  for (const entryRaw of entriesRaw) {
    const parsed = parseAuditEntry(entryRaw);
    if (parsed) parsedEntries.push(parsed);
  }

  return parsedEntries
    .sort((a, b) => b.at - a.at)
    .slice(0, MAX_AUDIT_ENTRIES);
}

function createAuditEntryId(): string {
  const randomUuid = globalThis.crypto?.randomUUID;
  if (typeof randomUuid === "function") {
    return randomUuid.call(globalThis.crypto);
  }

  const randomChunk = Math.floor(Math.random() * 1_000_000)
    .toString(36)
    .padStart(4, "0");

  return `audit_${Date.now().toString(36)}_${randomChunk}`;
}

interface SettingsStoreLike {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: unknown): Promise<void>;
  delete(key: string): Promise<void>;
}

function isSettingsStoreLike(value: unknown): value is SettingsStoreLike {
  if (!isRecord(value)) return false;

  return (
    typeof value.get === "function" &&
    typeof value.set === "function" &&
    typeof value.delete === "function"
  );
}

async function getSettingsStore(): Promise<SettingsStoreLike | null> {
  try {
    const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
    const appStorage = storageModule.getAppStorage();
    const settings = isRecord(appStorage) ? appStorage.settings : null;
    return isSettingsStoreLike(settings) ? settings : null;
  } catch {
    return null;
  }
}

async function readPersistedNativeHandle(): Promise<FileSystemDirectoryHandle | null> {
  const settings = await getSettingsStore();
  if (!settings) return null;

  try {
    const stored = await settings.get<unknown>(NATIVE_HANDLE_SETTING_KEY);
    return isDirectoryHandle(stored) ? stored : null;
  } catch {
    return null;
  }
}

async function persistNativeHandle(handle: FileSystemDirectoryHandle | null): Promise<void> {
  const settings = await getSettingsStore();
  if (!settings) return;

  try {
    if (handle) {
      await settings.set(NATIVE_HANDLE_SETTING_KEY, handle);
    } else {
      await settings.delete(NATIVE_HANDLE_SETTING_KEY);
    }
  } catch {
    // ignore persistence failures; fallback still works for the current session.
  }
}

async function queryReadWritePermission(
  handle: FileSystemDirectoryHandle,
): Promise<PermissionState | "unsupported"> {
  if (!isDirectoryPermissionHandle(handle)) {
    return "unsupported";
  }

  try {
    return await handle.queryPermission({ mode: "readwrite" });
  } catch {
    return "unsupported";
  }
}

async function requestReadWritePermission(
  handle: FileSystemDirectoryHandle,
): Promise<PermissionState | "unsupported"> {
  if (!isDirectoryPermissionHandle(handle)) {
    return "unsupported";
  }

  try {
    return await handle.requestPermission({ mode: "readwrite" });
  } catch {
    return "unsupported";
  }
}

function backendLabel(kind: WorkspaceBackendStatus["kind"]): string {
  switch (kind) {
    case "native-directory":
      return "Local folder";
    case "opfs":
      return "Sandboxed workspace";
    case "memory":
      return "Session memory";
  }
}

const NOTES_PREFIX = "notes/";
const NOTES_INDEX_PATH = "notes/index.md";
const IMPORTS_PREFIX = "imports/";
const SCRATCH_PREFIX = "scratch/";
const ASSISTANT_DOCS_PREFIX = "assistant-docs/";
const DEFAULT_CONTEXT_PREVIEW_LIMIT = 3;
const SCRATCH_CLEANUP_MAX_AGE_MS = 24 * 60 * 60 * 1000;

export interface WorkspaceContextSummary {
  summary: string;
  hasRelevantFiles: boolean;
  relevantSignature: string;
}

interface WorkspaceContextSummaryArgs {
  snapshot: WorkspaceSnapshot;
  currentWorkbookId: string | null;
  previewLimit?: number;
}

function buildContextFileSignature(file: WorkspaceFileEntry): string {
  const workbookId = file.workbookTag?.workbookId ?? "";
  return `${file.path}:${file.size}:${file.modifiedAt}:${workbookId}`;
}

function locationPriority(locationKind: WorkspaceFileLocationKind | undefined): number {
  switch (locationKind) {
    case "workspace":
      return 0;
    case "native-directory":
      return 1;
    case "builtin-doc":
      return 2;
    default:
      return 3;
  }
}

function sortFilesByPath(files: readonly WorkspaceFileEntry[]): WorkspaceFileEntry[] {
  return [...files].sort((left, right) => {
    const byPath = left.path.localeCompare(right.path);
    if (byPath !== 0) {
      return byPath;
    }

    const byLocation = locationPriority(left.locationKind) - locationPriority(right.locationKind);
    if (byLocation !== 0) {
      return byLocation;
    }

    return right.modifiedAt - left.modifiedAt;
  });
}

function formatPathPreview(files: readonly WorkspaceFileEntry[], limit: number): string {
  const safeLimit = Math.max(1, limit);
  const visible = files.slice(0, safeLimit).map((file) => file.path);
  if (visible.length === 0) {
    return "";
  }

  const suffix = files.length > visible.length ? ", …" : "";
  return ` (${visible.join(", ")}${suffix})`;
}

function isWorkspaceSourceFile(file: WorkspaceFileEntry): boolean {
  return file.sourceKind === "workspace";
}

export function buildWorkspaceContextSummary(args: WorkspaceContextSummaryArgs): WorkspaceContextSummary {
  const previewLimit = args.previewLimit ?? DEFAULT_CONTEXT_PREVIEW_LIMIT;
  const workspaceFiles = args.snapshot.files.filter(isWorkspaceSourceFile);

  const notesFiles = sortFilesByPath(workspaceFiles.filter((file) => file.path.startsWith(NOTES_PREFIX)));
  const notesIndexFile = notesFiles.find((file) => file.path === NOTES_INDEX_PATH) ?? null;

  const importFiles = sortFilesByPath(workspaceFiles.filter((file) => file.path.startsWith(IMPORTS_PREFIX)));

  const currentWorkbookFiles = args.currentWorkbookId
    ? sortFilesByPath(workspaceFiles.filter((file) => {
      if (file.workbookTag?.workbookId !== args.currentWorkbookId) {
        return false;
      }

      if (file.path.startsWith(NOTES_PREFIX)) return false;
      if (file.path.startsWith(IMPORTS_PREFIX)) return false;
      if (file.path.startsWith(SCRATCH_PREFIX)) return false;
      if (file.path.startsWith(ASSISTANT_DOCS_PREFIX)) return false;
      return true;
    }))
    : [];

  const hasRelevantFiles =
    notesFiles.length > 0 ||
    importFiles.length > 0 ||
    currentWorkbookFiles.length > 0;

  const signatureParts = [
    `backend:${args.snapshot.backend.kind}`,
    `workbook:${args.currentWorkbookId ?? "none"}`,
    `notes-files:${notesFiles.map((file) => buildContextFileSignature(file)).join(",")}`,
    `imports:${importFiles.map((file) => buildContextFileSignature(file)).join(",")}`,
    `workbook-files:${currentWorkbookFiles.map((file) => buildContextFileSignature(file)).join(",")}`,
  ];

  const lines: string[] = ["### Workspace"];

  if (notesFiles.length > 0) {
    const noteNoun = notesFiles.length === 1 ? "file" : "files";
    if (notesIndexFile) {
      lines.push(`- notes/: ${notesFiles.length} ${noteNoun}. Read notes/index.md first.`);
    } else {
      lines.push(`- notes/: ${notesFiles.length} ${noteNoun}. No notes/index.md yet.`);
    }
  }

  if (currentWorkbookFiles.length > 0) {
    const noun = currentWorkbookFiles.length === 1 ? "file" : "files";
    lines.push(
      `- Current workbook artifacts: ${currentWorkbookFiles.length} ${noun}${formatPathPreview(currentWorkbookFiles, previewLimit)}.`,
    );
  }

  if (importFiles.length > 0) {
    const noun = importFiles.length === 1 ? "file" : "files";
    lines.push(`- imports/: ${importFiles.length} ${noun}${formatPathPreview(importFiles, previewLimit)}.`);
  }

  if (!hasRelevantFiles) {
    lines.push("- No relevant workspace files for this workbook.");
  }

  lines.push("- Auto-context omits scratch/ and assistant-docs/. Use files list for full details.");

  return {
    summary: lines.join("\n"),
    hasRelevantFiles,
    relevantSignature: signatureParts.join("|"),
  };
}

export interface FilesWorkspaceOptions {
  initialBackend?: WorkspaceBackend;
  initialWorkspaceBackend?: WorkspaceBackend;
}

export class FilesWorkspace {
  private backend: WorkspaceBackend | null = null;
  private backendPromise: Promise<WorkspaceBackend> | null = null;
  private nativeHandle: FileSystemDirectoryHandle | null = null;

  private workspaceStorageBackend: WorkspaceBackend | null = null;
  private workspaceStorageBackendPromise: Promise<WorkspaceBackend> | null = null;

  private metadataLoaded = false;
  private readonly metadataByPath = new Map<string, WorkspaceFileWorkbookTag>();

  private auditLoaded = false;
  private auditEntries: FilesWorkspaceAuditEntry[] = [];

  private readonly scratchCleanupByBackend = new WeakMap<WorkspaceBackend, Promise<void>>();

  constructor(options: FilesWorkspaceOptions = {}) {
    if (options.initialBackend) {
      this.backend = options.initialBackend;
      if (options.initialBackend.kind !== "native-directory") {
        this.workspaceStorageBackend = options.initialBackend;
      }
    }

    if (options.initialWorkspaceBackend) {
      this.workspaceStorageBackend = options.initialWorkspaceBackend;
    }
  }

  private async initializeBackend(): Promise<WorkspaceBackend> {
    const persistedNative = await readPersistedNativeHandle();
    if (persistedNative) {
      const permission = await queryReadWritePermission(persistedNative);
      if (permission === "granted") {
        this.nativeHandle = persistedNative;
        return new NativeDirectoryBackend(persistedNative);
      }
    }

    if (typeof navigator !== "undefined" && navigator.storage && typeof navigator.storage.getDirectory === "function") {
      return new OpfsBackend();
    }

    return new MemoryBackend();
  }

  private initializeWorkspaceStorageBackend(): Promise<WorkspaceBackend> {
    if (typeof navigator !== "undefined" && navigator.storage && typeof navigator.storage.getDirectory === "function") {
      return Promise.resolve(new OpfsBackend());
    }

    return Promise.resolve(new MemoryBackend());
  }

  private async getWorkspaceStorageBackend(activeBackend?: WorkspaceBackend): Promise<WorkspaceBackend> {
    if (this.workspaceStorageBackend) {
      await this.ensureScratchCleanupAttempted(this.workspaceStorageBackend);
      return this.workspaceStorageBackend;
    }

    const backend = activeBackend ?? await this.getBackend();
    if (backend.kind !== "native-directory") {
      this.workspaceStorageBackend = backend;
      await this.ensureScratchCleanupAttempted(backend);
      return backend;
    }

    if (!this.workspaceStorageBackendPromise) {
      this.workspaceStorageBackendPromise = this.initializeWorkspaceStorageBackend();
    }

    const workspaceBackend = await this.workspaceStorageBackendPromise;
    this.workspaceStorageBackendPromise = null;
    this.workspaceStorageBackend = workspaceBackend;
    await this.ensureScratchCleanupAttempted(workspaceBackend);
    return workspaceBackend;
  }

  private async removeWorkbookTags(paths: readonly string[]): Promise<void> {
    await this.ensureMetadataLoaded();

    let changed = false;
    for (const path of paths) {
      if (this.metadataByPath.delete(path)) {
        changed = true;
      }
    }

    if (changed) {
      await this.persistMetadata();
    }
  }

  private async cleanupExpiredScratchFiles(backend: WorkspaceBackend): Promise<void> {
    const expirationCutoff = Date.now() - SCRATCH_CLEANUP_MAX_AGE_MS;

    const files = await backend.listFiles();
    const staleScratchPaths = files
      .filter((file) => file.path.startsWith(SCRATCH_PREFIX) && file.modifiedAt < expirationCutoff)
      .map((file) => file.path);

    if (staleScratchPaths.length === 0) {
      return;
    }

    const deletedPaths: string[] = [];

    for (const path of staleScratchPaths) {
      try {
        await backend.deleteFile(path);
        deletedPaths.push(path);
      } catch {
        // best-effort cleanup: ignore per-file delete failures
      }
    }

    if (deletedPaths.length === 0) {
      return;
    }

    await this.removeWorkbookTags(deletedPaths);
    dispatchWorkspaceChanged({ reason: "delete" });
  }

  private async ensureScratchCleanupAttempted(backend: WorkspaceBackend): Promise<void> {
    let cleanupPromise = this.scratchCleanupByBackend.get(backend);

    if (!cleanupPromise) {
      cleanupPromise = this.cleanupExpiredScratchFiles(backend).catch(() => {
        // best-effort cleanup: fail open on startup/first interaction
      });
      this.scratchCleanupByBackend.set(backend, cleanupPromise);
    }

    await cleanupPromise;
  }

  private async getBackend(): Promise<WorkspaceBackend> {
    if (this.backend) {
      await this.ensureScratchCleanupAttempted(this.backend);
      if (this.backend.kind !== "native-directory" && !this.workspaceStorageBackend) {
        this.workspaceStorageBackend = this.backend;
      }
      return this.backend;
    }

    if (!this.backendPromise) {
      this.backendPromise = this.initializeBackend();
    }

    const backend = await this.backendPromise;
    this.backend = backend;
    this.backendPromise = null;

    await this.ensureScratchCleanupAttempted(backend);
    if (backend.kind !== "native-directory") {
      this.workspaceStorageBackend = backend;
    }
    return backend;
  }

  private replaceBackend(nextBackend: WorkspaceBackend): void {
    this.backend = nextBackend;
    this.backendPromise = null;
    if (nextBackend.kind !== "native-directory") {
      this.workspaceStorageBackend = nextBackend;
      this.workspaceStorageBackendPromise = null;
    }
    dispatchWorkspaceChanged({ reason: "backend" });
  }

  private async ensureMetadataLoaded(): Promise<void> {
    if (this.metadataLoaded) return;
    this.metadataLoaded = true;

    const settings = await getSettingsStore();
    if (!settings) return;

    try {
      const raw = await settings.get<unknown>(METADATA_SETTING_KEY);
      const parsed = parsePersistedMetadata(raw);
      this.metadataByPath.clear();
      for (const [path, tag] of parsed) {
        this.metadataByPath.set(path, tag);
      }
    } catch {
      this.metadataByPath.clear();
    }
  }

  private async persistMetadata(): Promise<void> {
    const settings = await getSettingsStore();
    if (!settings) return;

    const byPath: Record<string, WorkspaceFileWorkbookTag> = {};
    for (const [path, tag] of this.metadataByPath) {
      byPath[path] = tag;
    }

    const payload: PersistedWorkspaceMetadata = {
      version: 1,
      byPath,
    };

    try {
      await settings.set(METADATA_SETTING_KEY, payload);
    } catch {
      // ignore persistence failures
    }
  }

  private async ensureAuditLoaded(): Promise<void> {
    if (this.auditLoaded) return;
    this.auditLoaded = true;

    const settings = await getSettingsStore();
    if (!settings) return;

    try {
      const raw = await settings.get<unknown>(AUDIT_TRAIL_SETTING_KEY);
      this.auditEntries = parsePersistedAuditTrail(raw);
    } catch {
      this.auditEntries = [];
    }
  }

  private async persistAuditTrail(): Promise<void> {
    const settings = await getSettingsStore();
    if (!settings) return;

    const payload: PersistedAuditTrail = {
      version: 1,
      entries: this.auditEntries,
    };

    try {
      await settings.set(AUDIT_TRAIL_SETTING_KEY, payload);
    } catch {
      // ignore persistence failures
    }
  }

  private async resolveActiveWorkbookTag(): Promise<WorkspaceFileWorkbookTag | null> {
    try {
      const context = await getWorkbookContext();
      if (!context.workbookId) return null;

      return {
        workbookId: context.workbookId,
        workbookLabel: formatWorkbookLabel(context),
        taggedAt: Date.now(),
      };
    } catch {
      return null;
    }
  }

  private async setWorkbookTagForPath(path: string, fallbackTag?: WorkspaceFileWorkbookTag): Promise<void> {
    await this.ensureMetadataLoaded();

    const resolvedTag = await this.resolveActiveWorkbookTag();
    const nextTag = resolvedTag ?? fallbackTag;
    if (!nextTag) return;

    this.metadataByPath.set(path, {
      workbookId: nextTag.workbookId,
      workbookLabel: nextTag.workbookLabel,
      taggedAt: Date.now(),
    });

    await this.persistMetadata();
  }

  private async removeWorkbookTag(path: string): Promise<void> {
    await this.ensureMetadataLoaded();

    if (!this.metadataByPath.delete(path)) {
      return;
    }

    await this.persistMetadata();
  }

  private async moveWorkbookTag(oldPath: string, newPath: string): Promise<void> {
    await this.ensureMetadataLoaded();

    const previousTag = this.metadataByPath.get(oldPath);
    this.metadataByPath.delete(oldPath);

    await this.setWorkbookTagForPath(newPath, previousTag);
    await this.persistMetadata();
  }

  private async pruneStaleWorkbookTags(currentPaths: Set<string>): Promise<void> {
    await this.ensureMetadataLoaded();

    let changed = false;
    for (const path of this.metadataByPath.keys()) {
      if (currentPaths.has(path)) continue;
      this.metadataByPath.delete(path);
      changed = true;
    }

    if (changed) {
      await this.persistMetadata();
    }
  }

  private async withWorkbookTags(entries: WorkspaceFileEntry[]): Promise<WorkspaceFileEntry[]> {
    await this.ensureMetadataLoaded();

    return entries.map((entry) => {
      const tag = this.metadataByPath.get(entry.path);
      if (!tag) return entry;

      return {
        ...entry,
        workbookTag: tag,
      };
    });
  }

  private async workspacePathExists(
    path: string,
    backend: WorkspaceBackend,
  ): Promise<boolean> {
    try {
      await backend.readFile(path);
      return true;
    } catch (error: unknown) {
      if (isMissingWorkspaceFileError(error)) {
        return false;
      }

      throw error;
    }
  }

  private async appendAuditEntry(args: {
    action: FilesWorkspaceAuditAction;
    backend: WorkspaceBackendKind;
    context: FilesWorkspaceAuditContext;
    path?: string;
    fromPath?: string;
    toPath?: string;
    bytes?: number;
  }): Promise<void> {
    await this.ensureAuditLoaded();

    const workbookTag = await this.resolveActiveWorkbookTag();
    const entry: FilesWorkspaceAuditEntry = {
      id: createAuditEntryId(),
      at: Date.now(),
      action: args.action,
      actor: args.context.actor,
      source: args.context.source,
      backend: args.backend,
      path: args.path,
      fromPath: args.fromPath,
      toPath: args.toPath,
      bytes: args.bytes,
      workbookId: workbookTag?.workbookId,
      workbookLabel: workbookTag?.workbookLabel,
    };

    this.auditEntries = [entry, ...this.auditEntries].slice(0, MAX_AUDIT_ENTRIES);
    await this.persistAuditTrail();
  }

  private withLocationKind(
    entries: readonly WorkspaceFileEntry[],
    locationKind: WorkspaceFileLocationKind,
  ): WorkspaceFileEntry[] {
    return entries.map((entry) => ({
      ...entry,
      locationKind,
    }));
  }

  private async resolveSourceBackends(): Promise<{
    activeBackend: WorkspaceBackend;
    workspaceBackend: WorkspaceBackend;
    nativeBackend: WorkspaceBackend | null;
  }> {
    const activeBackend = await this.getBackend();
    const workspaceBackend = await this.getWorkspaceStorageBackend(activeBackend);

    return {
      activeBackend,
      workspaceBackend,
      nativeBackend: activeBackend.kind === "native-directory" ? activeBackend : null,
    };
  }

  private async resolveMutationTarget(args: {
    requestedLocationKind?: WorkspaceFileLocationKind;
    defaultToWorkspace: boolean;
    path?: string;
  }): Promise<{
    backend: WorkspaceBackend;
    locationKind: WorkspaceFileLocationKind;
  }> {
    const { activeBackend, workspaceBackend, nativeBackend } = await this.resolveSourceBackends();

    if (args.requestedLocationKind === "workspace") {
      return {
        backend: workspaceBackend,
        locationKind: "workspace",
      };
    }

    if (args.requestedLocationKind === "native-directory") {
      if (!nativeBackend) {
        throw new Error("Connected folder is not available.");
      }

      return {
        backend: nativeBackend,
        locationKind: "native-directory",
      };
    }

    if (args.requestedLocationKind === "builtin-doc") {
      throw new Error("Built-in docs are read-only.");
    }

    if (args.path) {
      const workspaceHasPath = await this.workspacePathExists(args.path, workspaceBackend);
      const nativeHasPath = nativeBackend
        ? await this.workspacePathExists(args.path, nativeBackend)
        : false;

      if (workspaceHasPath && nativeHasPath) {
        throw new Error(
          `File '${args.path}' exists in both uploaded files and the connected folder. ` +
          "Select the file from Files and run the action there.",
        );
      }

      if (workspaceHasPath) {
        return {
          backend: workspaceBackend,
          locationKind: "workspace",
        };
      }

      if (nativeHasPath && nativeBackend) {
        return {
          backend: nativeBackend,
          locationKind: "native-directory",
        };
      }
    }

    if (args.defaultToWorkspace) {
      return {
        backend: workspaceBackend,
        locationKind: "workspace",
      };
    }

    if (activeBackend.kind === "native-directory") {
      return {
        backend: activeBackend,
        locationKind: "native-directory",
      };
    }

    return {
      backend: activeBackend,
      locationKind: "workspace",
    };
  }

  private async readRawFile(args: {
    path: string;
    requestedLocationKind?: WorkspaceFileLocationKind;
  }): Promise<{
    result: WorkspaceFileReadResult;
    locationKind: WorkspaceFileLocationKind;
    auditBackend: WorkspaceBackendKind;
  }> {
    const { activeBackend, workspaceBackend, nativeBackend } = await this.resolveSourceBackends();
    const builtinResult = getBuiltinWorkspaceDoc(args.path);

    if (args.requestedLocationKind === "workspace") {
      return {
        result: await workspaceBackend.readFile(args.path),
        locationKind: "workspace",
        auditBackend: workspaceBackend.kind,
      };
    }

    if (args.requestedLocationKind === "native-directory") {
      if (!nativeBackend) {
        throw new Error("Connected folder is not available.");
      }

      return {
        result: await nativeBackend.readFile(args.path),
        locationKind: "native-directory",
        auditBackend: nativeBackend.kind,
      };
    }

    if (args.requestedLocationKind === "builtin-doc") {
      if (!builtinResult) {
        throw new Error(`File not found: ${args.path}`);
      }

      return {
        result: builtinResult,
        locationKind: "builtin-doc",
        auditBackend: activeBackend.kind,
      };
    }

    try {
      return {
        result: await activeBackend.readFile(args.path),
        locationKind: activeBackend.kind === "native-directory" ? "native-directory" : "workspace",
        auditBackend: activeBackend.kind,
      };
    } catch (activeError: unknown) {
      if (!isMissingWorkspaceFileError(activeError)) {
        throw activeError;
      }

      if (activeBackend !== workspaceBackend) {
        try {
          return {
            result: await workspaceBackend.readFile(args.path),
            locationKind: "workspace",
            auditBackend: workspaceBackend.kind,
          };
        } catch (workspaceError: unknown) {
          if (!isMissingWorkspaceFileError(workspaceError)) {
            throw workspaceError;
          }
        }
      }

      if (builtinResult) {
        return {
          result: builtinResult,
          locationKind: "builtin-doc",
          auditBackend: activeBackend.kind,
        };
      }

      throw activeError;
    }
  }

  isNativeDirectoryPickerSupported(): boolean {
    if (typeof window === "undefined") return false;
    return isDirectoryPickerHost(window);
  }

  async connectNativeDirectory(options: WorkspaceMutationOptions = {}): Promise<void> {
    if (typeof window === "undefined" || !isDirectoryPickerHost(window)) {
      throw new Error("Native directory picker is not supported in this environment.");
    }

    const handle = await window.showDirectoryPicker();
    const permission = await requestReadWritePermission(handle);
    if (permission !== "granted") {
      throw new Error("Permission to the selected folder was not granted.");
    }

    const currentBackend = await this.getBackend();
    if (currentBackend.kind !== "native-directory") {
      this.workspaceStorageBackend = currentBackend;
      this.workspaceStorageBackendPromise = null;
    }

    this.nativeHandle = handle;
    await persistNativeHandle(handle);

    this.replaceBackend(new NativeDirectoryBackend(handle));

    await this.appendAuditEntry({
      action: "connect_native",
      backend: "native-directory",
      context: options.audit ?? DEFAULT_UI_AUDIT_CONTEXT,
    });
  }

  async disconnectNativeDirectory(options: WorkspaceMutationOptions = {}): Promise<void> {
    this.nativeHandle = null;
    await persistNativeHandle(null);

    const fallback = await this.getWorkspaceStorageBackend();

    this.replaceBackend(fallback);

    await this.appendAuditEntry({
      action: "disconnect_native",
      backend: fallback.kind,
      context: options.audit ?? DEFAULT_UI_AUDIT_CONTEXT,
    });
  }

  async getBackendStatus(): Promise<WorkspaceBackendStatus> {
    const backend = await this.getBackend();
    const nativeSupported = this.isNativeDirectoryPickerSupported();

    return {
      kind: backend.kind,
      label: backendLabel(backend.kind),
      nativeSupported,
      nativeConnected: backend.kind === "native-directory",
      nativeDirectoryName: this.nativeHandle?.name,
    };
  }

  async listFiles(options: WorkspaceListOptions = {}): Promise<WorkspaceFileEntry[]> {
    const { activeBackend, workspaceBackend, nativeBackend } = await this.resolveSourceBackends();

    const workspaceFiles = await workspaceBackend.listFiles();
    const taggedWorkspaceFiles = await this.withWorkbookTags(workspaceFiles);
    const workspaceEntries = this.withLocationKind(taggedWorkspaceFiles, "workspace");

    const currentPaths = new Set(taggedWorkspaceFiles.map((file) => file.path));
    await this.pruneStaleWorkbookTags(currentPaths);

    const nativeEntries = nativeBackend
      ? this.withLocationKind(await nativeBackend.listFiles(), "native-directory")
      : [];

    const occupiedPaths = new Set([
      ...workspaceEntries.map((file) => file.path),
      ...nativeEntries.map((file) => file.path),
    ]);

    const builtinEntries = this.withLocationKind(listBuiltinWorkspaceDocs(), "builtin-doc")
      .filter((file) => !occupiedPaths.has(file.path));

    const mergedFiles = sortFilesByPath([...workspaceEntries, ...nativeEntries, ...builtinEntries]);

    if (options.audit) {
      await this.appendAuditEntry({
        action: "list",
        backend: activeBackend.kind,
        context: options.audit,
      });
      dispatchWorkspaceChanged({ reason: "audit" });
    }

    return mergedFiles;
  }

  async getSnapshot(): Promise<WorkspaceSnapshot> {
    const [backend, files] = await Promise.all([
      this.getBackendStatus(),
      this.listFiles(),
    ]);

    const signature = files
      .map((file) => {
        const workbookSignature = file.workbookTag?.workbookId ?? "";
        return `${file.path}:${file.size}:${file.modifiedAt}:${workbookSignature}`;
      })
      .join("|");

    return {
      backend,
      files,
      signature,
    };
  }

  async readFile(path: string, opts: WorkspaceReadOptions = {}): Promise<WorkspaceFileReadResult> {
    const normalizedPath = normalizeWorkspacePath(path);
    const rawReadResult = await this.readRawFile({
      path: normalizedPath,
      requestedLocationKind: opts.locationKind,
    });

    const withTags =
      rawReadResult.locationKind === "workspace" &&
      rawReadResult.result.sourceKind === "workspace"
        ? await this.withWorkbookTags([rawReadResult.result])
        : [rawReadResult.result];

    const taggedResult = withTags[0];
    const result: WorkspaceFileReadResult = taggedResult
      ? {
        ...rawReadResult.result,
        workbookTag: taggedResult.workbookTag,
        locationKind: rawReadResult.locationKind,
      }
      : {
        ...rawReadResult.result,
        locationKind: rawReadResult.locationKind,
      };

    const mode = opts.mode ?? "auto";
    const maxChars = opts.maxChars ?? 20000;

    let resolved: WorkspaceFileReadResult;

    if (mode === "text") {
      if (result.text === undefined) {
        throw new Error(
          `File '${normalizedPath}' is binary (${result.mimeType}). Read it with mode=\"base64\" instead.`,
        );
      }

      const truncated = truncateText(result.text, maxChars);
      resolved = {
        ...result,
        text: truncated.text,
        base64: undefined,
        truncated: truncated.truncated,
      };
    } else if (mode === "base64") {
      const base64Content = result.base64 ?? bytesToBase64(encodeTextUtf8(result.text ?? ""));
      const truncated = truncateBase64(base64Content, maxChars);

      resolved = {
        ...result,
        text: undefined,
        base64: truncated.base64,
        truncated: truncated.truncated,
      };
    } else if (result.text !== undefined) {
      const truncated = truncateText(result.text, maxChars);
      resolved = {
        ...result,
        text: truncated.text,
        base64: undefined,
        truncated: truncated.truncated,
      };
    } else {
      const base64Content = result.base64 ?? "";
      const truncated = truncateBase64(base64Content, maxChars);
      resolved = {
        ...result,
        text: undefined,
        base64: truncated.base64,
        truncated: truncated.truncated,
      };
    }

    if (opts.audit) {
      await this.appendAuditEntry({
        action: "read",
        backend: rawReadResult.auditBackend,
        context: opts.audit,
        path: normalizedPath,
        bytes: resolved.size,
      });
      dispatchWorkspaceChanged({ reason: "audit" });
    }

    return resolved;
  }

  async writeTextFile(
    path: string,
    text: string,
    mimeTypeHint?: string,
    options: WorkspaceMutationOptions = {},
  ): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);
    const target = await this.resolveMutationTarget({
      requestedLocationKind: options.locationKind,
      defaultToWorkspace: false,
      path: normalizedPath,
    });

    if (isBuiltinWorkspacePath(normalizedPath)) {
      const hasWorkspaceCollision = await this.workspacePathExists(normalizedPath, target.backend);
      if (!hasWorkspaceCollision) {
        throw new Error(`'${normalizedPath}' is a built-in doc and cannot be modified.`);
      }
    }

    const bytes = encodeTextUtf8(text);

    await target.backend.writeBytes(
      normalizedPath,
      bytes,
      mimeTypeHint ?? inferMimeType(getWorkspaceBaseName(normalizedPath), "text/plain"),
    );

    if (target.locationKind === "workspace") {
      await this.setWorkbookTagForPath(normalizedPath);
    }

    await this.appendAuditEntry({
      action: "write",
      backend: target.backend.kind,
      context: options.audit ?? DEFAULT_UI_AUDIT_CONTEXT,
      path: normalizedPath,
      bytes: bytes.byteLength,
    });

    dispatchWorkspaceChanged({ reason: "write" });
  }

  async writeBase64File(
    path: string,
    base64: string,
    mimeTypeHint?: string,
    options: WorkspaceMutationOptions = {},
  ): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);
    const target = await this.resolveMutationTarget({
      requestedLocationKind: options.locationKind,
      defaultToWorkspace: false,
      path: normalizedPath,
    });

    if (isBuiltinWorkspacePath(normalizedPath)) {
      const hasWorkspaceCollision = await this.workspacePathExists(normalizedPath, target.backend);
      if (!hasWorkspaceCollision) {
        throw new Error(`'${normalizedPath}' is a built-in doc and cannot be modified.`);
      }
    }

    const bytes = base64ToBytes(base64);

    await target.backend.writeBytes(
      normalizedPath,
      bytes,
      mimeTypeHint ?? inferMimeType(getWorkspaceBaseName(normalizedPath)),
    );

    if (target.locationKind === "workspace") {
      await this.setWorkbookTagForPath(normalizedPath);
    }

    await this.appendAuditEntry({
      action: "write",
      backend: target.backend.kind,
      context: options.audit ?? DEFAULT_UI_AUDIT_CONTEXT,
      path: normalizedPath,
      bytes: bytes.byteLength,
    });

    dispatchWorkspaceChanged({ reason: "write" });
  }

  async deleteFile(path: string, options: WorkspaceMutationOptions = {}): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);
    const target = await this.resolveMutationTarget({
      requestedLocationKind: options.locationKind,
      defaultToWorkspace: false,
      path: normalizedPath,
    });

    if (isBuiltinWorkspacePath(normalizedPath)) {
      const hasWorkspaceCollision = await this.workspacePathExists(normalizedPath, target.backend);
      if (!hasWorkspaceCollision) {
        throw new Error(`'${normalizedPath}' is a built-in doc and cannot be deleted.`);
      }
    }

    await target.backend.deleteFile(normalizedPath);
    if (target.locationKind === "workspace") {
      await this.removeWorkbookTag(normalizedPath);
    }

    await this.appendAuditEntry({
      action: "delete",
      backend: target.backend.kind,
      context: options.audit ?? DEFAULT_UI_AUDIT_CONTEXT,
      path: normalizedPath,
    });

    dispatchWorkspaceChanged({ reason: "delete" });
  }

  async renameFile(
    oldPath: string,
    newPath: string,
    options: WorkspaceMutationOptions = {},
  ): Promise<void> {
    const normalizedOldPath = normalizeWorkspacePath(oldPath);
    const normalizedNewPath = normalizeWorkspacePath(newPath);

    const target = await this.resolveMutationTarget({
      requestedLocationKind: options.locationKind,
      defaultToWorkspace: false,
      path: normalizedOldPath,
    });

    if (isBuiltinWorkspacePath(normalizedOldPath)) {
      const hasWorkspaceCollision = await this.workspacePathExists(normalizedOldPath, target.backend);
      if (!hasWorkspaceCollision) {
        throw new Error(`'${normalizedOldPath}' is a built-in doc and cannot be renamed.`);
      }
    }

    if (isBuiltinWorkspacePath(normalizedNewPath)) {
      throw new Error(`'${normalizedNewPath}' is reserved for a built-in doc.`);
    }

    await target.backend.renameFile(normalizedOldPath, normalizedNewPath);
    if (target.locationKind === "workspace") {
      await this.moveWorkbookTag(normalizedOldPath, normalizedNewPath);
    }

    await this.appendAuditEntry({
      action: "rename",
      backend: target.backend.kind,
      context: options.audit ?? DEFAULT_UI_AUDIT_CONTEXT,
      fromPath: normalizedOldPath,
      toPath: normalizedNewPath,
    });

    dispatchWorkspaceChanged({ reason: "rename" });
  }

  async importFiles(files: Iterable<File>, options: WorkspaceMutationOptions = {}): Promise<number> {
    const target = await this.resolveMutationTarget({
      requestedLocationKind: options.locationKind,
      defaultToWorkspace: true,
    });

    let imported = 0;
    let importedBytes = 0;

    for (const file of files) {
      const relativePath = typeof file.webkitRelativePath === "string"
        ? file.webkitRelativePath.trim()
        : "";

      const preferredPath = relativePath.length > 0
        ? relativePath
        : file.name;

      const normalizedPath = normalizeWorkspacePath(preferredPath);
      if (isBuiltinWorkspacePath(normalizedPath)) {
        const hasWorkspaceCollision = await this.workspacePathExists(normalizedPath, target.backend);
        if (!hasWorkspaceCollision) {
          throw new Error(`'${normalizedPath}' is reserved for a built-in doc.`);
        }
      }

      const bytes = new Uint8Array(await file.arrayBuffer());
      await target.backend.writeBytes(
        normalizedPath,
        bytes,
        inferMimeType(file.name, file.type),
      );

      if (target.locationKind === "workspace") {
        await this.setWorkbookTagForPath(normalizedPath);
      }

      imported += 1;
      importedBytes += bytes.byteLength;
    }

    if (imported > 0) {
      await this.appendAuditEntry({
        action: "import",
        backend: target.backend.kind,
        context: options.audit ?? DEFAULT_UI_AUDIT_CONTEXT,
        bytes: importedBytes,
      });

      dispatchWorkspaceChanged({ reason: "import" });
    }

    return imported;
  }

  async listAuditEntries(limit = 40): Promise<FilesWorkspaceAuditEntry[]> {
    await this.ensureAuditLoaded();

    const safeLimit = Math.max(0, Math.min(limit, MAX_AUDIT_ENTRIES));
    return this.auditEntries.slice(0, safeLimit);
  }

  async clearAuditTrail(_options: WorkspaceMutationOptions = {}): Promise<void> {
    await this.ensureAuditLoaded();

    this.auditEntries = [];
    await this.persistAuditTrail();

    dispatchWorkspaceChanged({ reason: "audit" });
  }

  async downloadFile(
    path: string,
    options: { locationKind?: WorkspaceFileLocationKind } = {},
  ): Promise<void> {
    if (typeof document === "undefined") {
      throw new Error("Downloads are not available in this environment.");
    }

    const pendingWindow = window.open("", "_blank");

    try {
      const normalizedPath = normalizeWorkspacePath(path);
      const readResult = await this.readRawFile({
        path: normalizedPath,
        requestedLocationKind: options.locationKind,
      });
      const result = readResult.result;

      const bytes = result.base64
        ? base64ToBytes(result.base64)
        : encodeTextUtf8(result.text ?? "");

      const mimeType = result.mimeType && isTextMimeType(result.mimeType)
        ? result.mimeType
        : inferMimeType(result.name, result.mimeType);

      // Sanitize script-capable MIME types (HTML, SVG, JS) to prevent
      // active-content execution at the app origin when opened via blob URL.
      const safeMimeType = resolveSafeBlobUrlMimeType(mimeType);

      const blob = new Blob([toArrayBuffer(bytes)], { type: safeMimeType });
      const url = URL.createObjectURL(blob);

      // Open a blank tab synchronously before async file IO (above) to keep
      // user activation in Office WKWebView. Then navigate it to the blob URL.
      // Fall back to <a download> when popups are blocked.
      const openedInWindow = tryOpenDownloadWindow(url, pendingWindow);
      if (!openedInWindow) {
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = result.name;
        anchor.rel = "noopener";
        anchor.style.display = "none";
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
      }

      // Delay revocation so the opened window can finish loading the blob.
      setTimeout(() => URL.revokeObjectURL(url), 60_000);
    } catch (error: unknown) {
      closeWindowSafely(pendingWindow);
      throw error;
    }
  }

  async getContextSummary(currentWorkbookId: string | null): Promise<WorkspaceContextSummary> {
    const snapshot = await this.getSnapshot();
    return buildWorkspaceContextSummary({
      snapshot,
      currentWorkbookId,
    });
  }
}

let workspaceSingleton: FilesWorkspace | null = null;

export function getFilesWorkspace(): FilesWorkspace {
  if (!workspaceSingleton) {
    workspaceSingleton = new FilesWorkspace();
  }

  return workspaceSingleton;
}
