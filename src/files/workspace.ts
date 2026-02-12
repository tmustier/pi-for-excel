/**
 * Files workspace manager.
 *
 * Backend selection strategy:
 * 1) persisted native directory handle (when permission is still granted)
 * 2) OPFS
 * 3) in-memory fallback (non-browser/test environments)
 */

import { isRecord } from "../utils/type-guards.js";
import { base64ToBytes, bytesToBase64, encodeTextUtf8, truncateBase64, truncateText } from "./encoding.js";
import { MemoryBackend, NativeDirectoryBackend, OpfsBackend, type WorkspaceBackend } from "./backend.js";
import { formatBytes, inferMimeType, isTextMimeType } from "./mime.js";
import { getWorkspaceBaseName, normalizeWorkspacePath } from "./path.js";
import {
  FILES_WORKSPACE_CHANGED_EVENT,
  type FilesWorkspaceChangedDetail,
  type WorkspaceBackendStatus,
  type WorkspaceFileEntry,
  type WorkspaceFileReadResult,
  type WorkspaceSnapshot,
} from "./types.js";

const NATIVE_HANDLE_SETTING_KEY = "files.workspace.nativeHandle.v1";

export type WorkspaceReadMode = "auto" | "text" | "base64";

export interface WorkspaceReadOptions {
  mode?: WorkspaceReadMode;
  maxChars?: number;
}

interface DirectoryPickerHost {
  showDirectoryPicker: () => Promise<FileSystemDirectoryHandle>;
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

function dispatchWorkspaceChanged(detail: FilesWorkspaceChangedDetail): void {
  if (typeof document === "undefined") return;
  document.dispatchEvent(new CustomEvent<FilesWorkspaceChangedDetail>(FILES_WORKSPACE_CHANGED_EVENT, { detail }));
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const out = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(out).set(bytes);
  return out;
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
  try {
    return await handle.queryPermission({ mode: "readwrite" });
  } catch {
    return "unsupported";
  }
}

async function requestReadWritePermission(
  handle: FileSystemDirectoryHandle,
): Promise<PermissionState | "unsupported"> {
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

export class FilesWorkspace {
  private backend: WorkspaceBackend | null = null;
  private backendPromise: Promise<WorkspaceBackend> | null = null;
  private nativeHandle: FileSystemDirectoryHandle | null = null;

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

  private async getBackend(): Promise<WorkspaceBackend> {
    if (this.backend) return this.backend;

    if (!this.backendPromise) {
      this.backendPromise = this.initializeBackend();
    }

    const backend = await this.backendPromise;
    this.backend = backend;
    this.backendPromise = null;
    return backend;
  }

  private replaceBackend(nextBackend: WorkspaceBackend): void {
    this.backend = nextBackend;
    this.backendPromise = null;
    dispatchWorkspaceChanged({ reason: "backend" });
  }

  isNativeDirectoryPickerSupported(): boolean {
    if (typeof window === "undefined") return false;
    return isDirectoryPickerHost(window);
  }

  async connectNativeDirectory(): Promise<void> {
    if (typeof window === "undefined" || !isDirectoryPickerHost(window)) {
      throw new Error("Native directory picker is not supported in this environment.");
    }

    const handle = await window.showDirectoryPicker();
    const permission = await requestReadWritePermission(handle);
    if (permission !== "granted") {
      throw new Error("Permission to the selected folder was not granted.");
    }

    this.nativeHandle = handle;
    await persistNativeHandle(handle);
    this.replaceBackend(new NativeDirectoryBackend(handle));
  }

  async disconnectNativeDirectory(): Promise<void> {
    this.nativeHandle = null;
    await persistNativeHandle(null);

    const fallback =
      typeof navigator !== "undefined" && navigator.storage && typeof navigator.storage.getDirectory === "function"
        ? new OpfsBackend()
        : new MemoryBackend();

    this.replaceBackend(fallback);
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

  async listFiles(): Promise<WorkspaceFileEntry[]> {
    const backend = await this.getBackend();
    return backend.listFiles();
  }

  async getSnapshot(): Promise<WorkspaceSnapshot> {
    const [backend, files] = await Promise.all([
      this.getBackendStatus(),
      this.listFiles(),
    ]);

    const signature = files
      .map((file) => `${file.path}:${file.size}:${file.modifiedAt}`)
      .join("|");

    return {
      backend,
      files,
      signature,
    };
  }

  async readFile(path: string, opts: WorkspaceReadOptions = {}): Promise<WorkspaceFileReadResult> {
    const normalizedPath = normalizeWorkspacePath(path);
    const backend = await this.getBackend();
    const result = await backend.readFile(normalizedPath);

    const mode = opts.mode ?? "auto";
    const maxChars = opts.maxChars ?? 20000;

    if (mode === "text") {
      if (result.text === undefined) {
        throw new Error(
          `File '${normalizedPath}' is binary (${result.mimeType}). Read it with mode=\"base64\" instead.`,
        );
      }

      const truncated = truncateText(result.text, maxChars);
      return {
        ...result,
        text: truncated.text,
        base64: undefined,
        truncated: truncated.truncated,
      };
    }

    if (mode === "base64") {
      const base64Content = result.base64 ?? bytesToBase64(encodeTextUtf8(result.text ?? ""));
      const truncated = truncateBase64(base64Content, maxChars);

      return {
        ...result,
        text: undefined,
        base64: truncated.base64,
        truncated: truncated.truncated,
      };
    }

    // auto mode
    if (result.text !== undefined) {
      const truncated = truncateText(result.text, maxChars);
      return {
        ...result,
        text: truncated.text,
        base64: undefined,
        truncated: truncated.truncated,
      };
    }

    const base64Content = result.base64 ?? "";
    const truncated = truncateBase64(base64Content, maxChars);
    return {
      ...result,
      text: undefined,
      base64: truncated.base64,
      truncated: truncated.truncated,
    };
  }

  async writeTextFile(path: string, text: string, mimeTypeHint?: string): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);
    const bytes = encodeTextUtf8(text);
    const backend = await this.getBackend();

    await backend.writeBytes(
      normalizedPath,
      bytes,
      mimeTypeHint ?? inferMimeType(getWorkspaceBaseName(normalizedPath), "text/plain"),
    );

    dispatchWorkspaceChanged({ reason: "write" });
  }

  async writeBase64File(path: string, base64: string, mimeTypeHint?: string): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);
    const bytes = base64ToBytes(base64);
    const backend = await this.getBackend();

    await backend.writeBytes(
      normalizedPath,
      bytes,
      mimeTypeHint ?? inferMimeType(getWorkspaceBaseName(normalizedPath)),
    );

    dispatchWorkspaceChanged({ reason: "write" });
  }

  async deleteFile(path: string): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);
    const backend = await this.getBackend();

    await backend.deleteFile(normalizedPath);
    dispatchWorkspaceChanged({ reason: "delete" });
  }

  async renameFile(oldPath: string, newPath: string): Promise<void> {
    const normalizedOldPath = normalizeWorkspacePath(oldPath);
    const normalizedNewPath = normalizeWorkspacePath(newPath);
    const backend = await this.getBackend();

    await backend.renameFile(normalizedOldPath, normalizedNewPath);
    dispatchWorkspaceChanged({ reason: "rename" });
  }

  async importFiles(files: Iterable<File>): Promise<number> {
    const backend = await this.getBackend();
    let imported = 0;

    for (const file of files) {
      const preferredPath = file.webkitRelativePath.trim().length > 0
        ? file.webkitRelativePath
        : file.name;

      const normalizedPath = normalizeWorkspacePath(preferredPath);
      const bytes = new Uint8Array(await file.arrayBuffer());
      await backend.writeBytes(
        normalizedPath,
        bytes,
        inferMimeType(file.name, file.type),
      );
      imported += 1;
    }

    if (imported > 0) {
      dispatchWorkspaceChanged({ reason: "import" });
    }

    return imported;
  }

  async downloadFile(path: string): Promise<void> {
    if (typeof document === "undefined") {
      throw new Error("Downloads are not available in this environment.");
    }

    const normalizedPath = normalizeWorkspacePath(path);
    const result = await (await this.getBackend()).readFile(normalizedPath);

    const bytes = result.base64
      ? base64ToBytes(result.base64)
      : encodeTextUtf8(result.text ?? "");

    const mimeType = result.mimeType && isTextMimeType(result.mimeType)
      ? result.mimeType
      : inferMimeType(result.name, result.mimeType);

    const blob = new Blob([toArrayBuffer(bytes)], { type: mimeType });
    const url = URL.createObjectURL(blob);

    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = result.name;
    anchor.rel = "noopener";
    anchor.style.display = "none";

    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  }

  async getContextSummary(maxFiles = 20): Promise<string | null> {
    const snapshot = await this.getSnapshot();

    if (snapshot.files.length === 0) return null;

    const lines: string[] = [];
    lines.push(`### Workspace Files (${snapshot.backend.label})`);

    const visible = snapshot.files.slice(0, maxFiles);
    for (const file of visible) {
      lines.push(`- ${file.path} (${formatBytes(file.size)}, ${file.kind})`);
    }

    const remaining = snapshot.files.length - visible.length;
    if (remaining > 0) {
      lines.push(`- … and ${remaining} more`);
    }

    return lines.join("\n");
  }
}

let workspaceSingleton: FilesWorkspace | null = null;

export function getFilesWorkspace(): FilesWorkspace {
  if (!workspaceSingleton) {
    workspaceSingleton = new FilesWorkspace();
  }

  return workspaceSingleton;
}
