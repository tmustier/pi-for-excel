/**
 * Minimal File System Access API typings used by the files workspace.
 *
 * TypeScript's bundled DOM lib does not always include these interfaces,
 * so we declare the subset we need.
 */

interface FileSystemHandlePermissionDescriptor {
  mode?: "read" | "readwrite";
}

interface PiFileSystemWritableFileStream {
  write(data: PiFileSystemWriteChunkType): Promise<void>;
  close(): Promise<void>;
}

type PiFileSystemWriteChunkType =
  | ArrayBuffer
  | Blob
  | string
  | ArrayBufferView
  | {
    type: "write";
    position?: number;
    data: string | Blob | BufferSource;
  }
  | {
    type: "truncate";
    size: number;
  }
  | {
    type: "seek";
    position: number;
  };

interface FileSystemHandle {
  readonly kind: "file" | "directory";
  readonly name: string;
}

interface FileSystemFileHandle extends FileSystemHandle {
  readonly kind: "file";
  getFile(): Promise<File>;
  createWritable(options?: { keepExistingData?: boolean }): Promise<PiFileSystemWritableFileStream>;
}

type FileSystemEntryHandle = FileSystemFileHandle | FileSystemDirectoryHandle;

interface FileSystemDirectoryHandle extends FileSystemHandle {
  readonly kind: "directory";
  values(): AsyncIterableIterator<FileSystemEntryHandle>;
  keys(): AsyncIterableIterator<string>;
  entries(): AsyncIterableIterator<[string, FileSystemEntryHandle]>;
  [Symbol.asyncIterator](): AsyncIterableIterator<[string, FileSystemEntryHandle]>;
  getFileHandle(name: string, options?: { create?: boolean }): Promise<FileSystemFileHandle>;
  getDirectoryHandle(name: string, options?: { create?: boolean }): Promise<FileSystemDirectoryHandle>;
  removeEntry(name: string, options?: { recursive?: boolean }): Promise<void>;
  queryPermission(descriptor?: FileSystemHandlePermissionDescriptor): Promise<PermissionState>;
  requestPermission(descriptor?: FileSystemHandlePermissionDescriptor): Promise<PermissionState>;
}

interface DirectoryPickerHost {
  showDirectoryPicker(options?: {
    id?: string;
    mode?: "read" | "readwrite";
    startIn?: "desktop" | "documents" | "downloads" | "music" | "pictures" | "videos";
  }): Promise<FileSystemDirectoryHandle>;
}

interface Window {
  showDirectoryPicker?: DirectoryPickerHost["showDirectoryPicker"];
}

interface StorageManager {
  getDirectory(): Promise<FileSystemDirectoryHandle>;
}
