/**
 * Workspace storage backends.
 */

import { base64ToBytes, bytesToBase64, decodeTextUtf8, encodeTextUtf8 } from "./encoding.js";
import { inferFileKind, inferMimeType } from "./mime.js";
import { getWorkspaceBaseName, normalizeWorkspacePath, splitWorkspacePath } from "./path.js";
import type { WorkspaceBackendKind, WorkspaceFileEntry, WorkspaceFileReadResult } from "./types.js";

export interface WorkspaceBackend {
  readonly kind: WorkspaceBackendKind;
  readonly label: string;

  listFiles(): Promise<WorkspaceFileEntry[]>;
  readFile(path: string): Promise<WorkspaceFileReadResult>;
  writeBytes(path: string, bytes: Uint8Array, mimeTypeHint?: string): Promise<void>;
  deleteFile(path: string): Promise<void>;
  renameFile(oldPath: string, newPath: string): Promise<void>;
}

async function fileToBytes(file: File): Promise<Uint8Array> {
  const buffer = await file.arrayBuffer();
  return new Uint8Array(buffer);
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const out = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(out).set(bytes);
  return out;
}

function mapFileToEntry(path: string, file: File): WorkspaceFileEntry {
  const mimeType = inferMimeType(file.name, file.type);
  return {
    path,
    name: file.name,
    size: file.size,
    modifiedAt: file.lastModified,
    mimeType,
    kind: inferFileKind(file.name, mimeType),
    sourceKind: "workspace",
    readOnly: false,
  };
}

function mapFileToReadResult(path: string, file: File, bytes: Uint8Array): WorkspaceFileReadResult {
  const entry = mapFileToEntry(path, file);

  if (entry.kind === "text") {
    return {
      ...entry,
      text: decodeTextUtf8(bytes),
    };
  }

  return {
    ...entry,
    base64: bytesToBase64(bytes),
  };
}

async function resolveParentDirectoryHandle(args: {
  root: FileSystemDirectoryHandle;
  pathParts: string[];
  createDirectories: boolean;
}): Promise<FileSystemDirectoryHandle> {
  let current = args.root;

  for (let i = 0; i < args.pathParts.length - 1; i += 1) {
    const segment = args.pathParts[i];
    current = await current.getDirectoryHandle(segment, { create: args.createDirectories });
  }

  return current;
}

async function resolveFileHandle(args: {
  root: FileSystemDirectoryHandle;
  path: string;
  createFile: boolean;
  createDirectories: boolean;
}): Promise<FileSystemFileHandle> {
  const normalizedPath = normalizeWorkspacePath(args.path);
  const pathParts = splitWorkspacePath(normalizedPath);
  const fileName = pathParts[pathParts.length - 1];

  if (!fileName) {
    throw new Error("Path must include a file name.");
  }

  const parent = await resolveParentDirectoryHandle({
    root: args.root,
    pathParts,
    createDirectories: args.createDirectories,
  });

  return parent.getFileHandle(fileName, { create: args.createFile });
}

async function removeFileHandle(root: FileSystemDirectoryHandle, path: string): Promise<void> {
  const normalizedPath = normalizeWorkspacePath(path);
  const pathParts = splitWorkspacePath(normalizedPath);
  const fileName = pathParts[pathParts.length - 1];

  if (!fileName) {
    throw new Error("Path must include a file name.");
  }

  const parent = await resolveParentDirectoryHandle({
    root,
    pathParts,
    createDirectories: false,
  });

  await parent.removeEntry(fileName, { recursive: false });
}

async function listDirectoryRecursively(args: {
  root: FileSystemDirectoryHandle;
  pathPrefix: string;
  out: WorkspaceFileEntry[];
}): Promise<void> {
  for await (const [name, handle] of args.root.entries()) {
    const path = args.pathPrefix.length > 0 ? `${args.pathPrefix}/${name}` : name;

    if (handle.kind === "file") {
      const file = await handle.getFile();
      args.out.push(mapFileToEntry(path, file));
      continue;
    }

    if (handle.kind === "directory") {
      await listDirectoryRecursively({
        root: handle,
        pathPrefix: path,
        out: args.out,
      });
    }
  }
}

function sortEntries(entries: WorkspaceFileEntry[]): WorkspaceFileEntry[] {
  return entries.sort((a, b) => a.path.localeCompare(b.path));
}

export class NativeDirectoryBackend implements WorkspaceBackend {
  readonly kind = "native-directory";
  readonly label = "Local folder";

  private readonly root: FileSystemDirectoryHandle;

  constructor(root: FileSystemDirectoryHandle) {
    this.root = root;
  }

  async listFiles(): Promise<WorkspaceFileEntry[]> {
    const out: WorkspaceFileEntry[] = [];
    await listDirectoryRecursively({
      root: this.root,
      pathPrefix: "",
      out,
    });

    return sortEntries(out);
  }

  async readFile(path: string): Promise<WorkspaceFileReadResult> {
    const normalizedPath = normalizeWorkspacePath(path);
    const handle = await resolveFileHandle({
      root: this.root,
      path: normalizedPath,
      createFile: false,
      createDirectories: false,
    });

    const file = await handle.getFile();
    const bytes = await fileToBytes(file);
    return mapFileToReadResult(normalizedPath, file, bytes);
  }

  async writeBytes(path: string, bytes: Uint8Array, _mimeTypeHint?: string): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);
    const handle = await resolveFileHandle({
      root: this.root,
      path: normalizedPath,
      createFile: true,
      createDirectories: true,
    });

    const writable = await handle.createWritable();
    await writable.write(toArrayBuffer(bytes));
    await writable.close();
  }

  async deleteFile(path: string): Promise<void> {
    await removeFileHandle(this.root, path);
  }

  async renameFile(oldPath: string, newPath: string): Promise<void> {
    const normalizedOldPath = normalizeWorkspacePath(oldPath);
    const normalizedNewPath = normalizeWorkspacePath(newPath);

    if (normalizedOldPath === normalizedNewPath) return;

    const oldHandle = await resolveFileHandle({
      root: this.root,
      path: normalizedOldPath,
      createFile: false,
      createDirectories: false,
    });

    const file = await oldHandle.getFile();
    const bytes = await fileToBytes(file);

    await this.writeBytes(normalizedNewPath, bytes, inferMimeType(getWorkspaceBaseName(normalizedNewPath), file.type));
    await this.deleteFile(normalizedOldPath);
  }
}

async function getOpfsRoot(): Promise<FileSystemDirectoryHandle> {
  const storage = navigator.storage;
  if (!storage || typeof storage.getDirectory !== "function") {
    throw new Error("Origin Private File System is not available in this environment.");
  }

  return storage.getDirectory();
}

export class OpfsBackend implements WorkspaceBackend {
  readonly kind = "opfs";
  readonly label = "Sandboxed workspace";

  async listFiles(): Promise<WorkspaceFileEntry[]> {
    const root = await getOpfsRoot();
    const out: WorkspaceFileEntry[] = [];

    await listDirectoryRecursively({
      root,
      pathPrefix: "",
      out,
    });

    return sortEntries(out);
  }

  async readFile(path: string): Promise<WorkspaceFileReadResult> {
    const root = await getOpfsRoot();
    const normalizedPath = normalizeWorkspacePath(path);

    const handle = await resolveFileHandle({
      root,
      path: normalizedPath,
      createFile: false,
      createDirectories: false,
    });

    const file = await handle.getFile();
    const bytes = await fileToBytes(file);
    return mapFileToReadResult(normalizedPath, file, bytes);
  }

  async writeBytes(path: string, bytes: Uint8Array, _mimeTypeHint?: string): Promise<void> {
    const root = await getOpfsRoot();
    const normalizedPath = normalizeWorkspacePath(path);

    const handle = await resolveFileHandle({
      root,
      path: normalizedPath,
      createFile: true,
      createDirectories: true,
    });

    const writable = await handle.createWritable();
    await writable.write(toArrayBuffer(bytes));
    await writable.close();
  }

  async deleteFile(path: string): Promise<void> {
    const root = await getOpfsRoot();
    await removeFileHandle(root, path);
  }

  async renameFile(oldPath: string, newPath: string): Promise<void> {
    const normalizedOldPath = normalizeWorkspacePath(oldPath);
    const normalizedNewPath = normalizeWorkspacePath(newPath);

    if (normalizedOldPath === normalizedNewPath) return;

    const file = await this.readFile(normalizedOldPath);
    if (!file.base64 && !file.text) {
      throw new Error("Could not read file during rename.");
    }

    const bytes = file.base64
      ? base64ToBytes(file.base64)
      : encodeTextUtf8(file.text ?? "");

    await this.writeBytes(normalizedNewPath, bytes, file.mimeType);
    await this.deleteFile(normalizedOldPath);
  }
}

interface MemoryFileRecord {
  bytes: Uint8Array;
  mimeType: string;
  modifiedAt: number;
}

export class MemoryBackend implements WorkspaceBackend {
  readonly kind = "memory";
  readonly label = "Session memory";

  private readonly files = new Map<string, MemoryFileRecord>();

  listFiles(): Promise<WorkspaceFileEntry[]> {
    const out: WorkspaceFileEntry[] = [];

    for (const [path, record] of this.files) {
      const name = getWorkspaceBaseName(path);
      out.push({
        path,
        name,
        size: record.bytes.byteLength,
        modifiedAt: record.modifiedAt,
        mimeType: record.mimeType,
        kind: inferFileKind(name, record.mimeType),
        sourceKind: "workspace",
        readOnly: false,
      });
    }

    return Promise.resolve(sortEntries(out));
  }

  readFile(path: string): Promise<WorkspaceFileReadResult> {
    const normalizedPath = normalizeWorkspacePath(path);
    const record = this.files.get(normalizedPath);
    if (!record) {
      return Promise.reject(new Error(`File not found: ${normalizedPath}`));
    }

    const name = getWorkspaceBaseName(normalizedPath);
    const kind = inferFileKind(name, record.mimeType);

    if (kind === "text") {
      return Promise.resolve({
        path: normalizedPath,
        name,
        size: record.bytes.byteLength,
        modifiedAt: record.modifiedAt,
        mimeType: record.mimeType,
        kind,
        sourceKind: "workspace",
        readOnly: false,
        text: decodeTextUtf8(record.bytes),
      });
    }

    return Promise.resolve({
      path: normalizedPath,
      name,
      size: record.bytes.byteLength,
      modifiedAt: record.modifiedAt,
      mimeType: record.mimeType,
      kind,
      sourceKind: "workspace",
      readOnly: false,
      base64: bytesToBase64(record.bytes),
    });
  }

  writeBytes(path: string, bytes: Uint8Array, mimeTypeHint?: string): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);
    const name = getWorkspaceBaseName(normalizedPath);

    this.files.set(normalizedPath, {
      bytes,
      mimeType: inferMimeType(name, mimeTypeHint),
      modifiedAt: Date.now(),
    });

    return Promise.resolve();
  }

  deleteFile(path: string): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);
    this.files.delete(normalizedPath);
    return Promise.resolve();
  }

  renameFile(oldPath: string, newPath: string): Promise<void> {
    const normalizedOldPath = normalizeWorkspacePath(oldPath);
    const normalizedNewPath = normalizeWorkspacePath(newPath);

    if (normalizedOldPath === normalizedNewPath) {
      return Promise.resolve();
    }

    const record = this.files.get(normalizedOldPath);
    if (!record) {
      return Promise.reject(new Error(`File not found: ${normalizedOldPath}`));
    }

    this.files.delete(normalizedOldPath);
    this.files.set(normalizedNewPath, {
      bytes: record.bytes,
      mimeType: inferMimeType(getWorkspaceBaseName(normalizedNewPath), record.mimeType),
      modifiedAt: Date.now(),
    });

    return Promise.resolve();
  }
}
