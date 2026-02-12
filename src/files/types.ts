/**
 * Workspace file subsystem shared types.
 */

export type WorkspaceBackendKind = "native-directory" | "opfs" | "memory";

export type WorkspaceFileKind = "text" | "binary";

export type WorkspaceFileSourceKind = "workspace" | "builtin-doc";

export interface WorkspaceFileWorkbookTag {
  workbookId: string;
  workbookLabel: string;
  taggedAt: number;
}

export interface WorkspaceFileEntry {
  path: string;
  name: string;
  size: number;
  modifiedAt: number;
  mimeType: string;
  kind: WorkspaceFileKind;
  sourceKind: WorkspaceFileSourceKind;
  readOnly: boolean;
  workbookTag?: WorkspaceFileWorkbookTag;
}

export interface WorkspaceFileReadResult extends WorkspaceFileEntry {
  text?: string;
  base64?: string;
  truncated?: boolean;
}

export interface WorkspaceBackendStatus {
  kind: WorkspaceBackendKind;
  label: string;
  nativeSupported: boolean;
  nativeConnected: boolean;
  nativeDirectoryName?: string;
}

export interface WorkspaceSnapshot {
  backend: WorkspaceBackendStatus;
  files: WorkspaceFileEntry[];
  signature: string;
}

export type FilesWorkspaceAuditActor = "assistant" | "user" | "system";

export type FilesWorkspaceAuditAction =
  | "list"
  | "read"
  | "write"
  | "delete"
  | "rename"
  | "import"
  | "connect_native"
  | "disconnect_native"
  | "clear_audit";

export interface FilesWorkspaceAuditEntry {
  id: string;
  at: number;
  action: FilesWorkspaceAuditAction;
  actor: FilesWorkspaceAuditActor;
  source: string;
  backend: WorkspaceBackendKind;
  path?: string;
  fromPath?: string;
  toPath?: string;
  bytes?: number;
  workbookId?: string;
  workbookLabel?: string;
}

export const FILES_WORKSPACE_CHANGED_EVENT = "pi:files-workspace-changed";

export interface FilesWorkspaceChangedDetail {
  reason: "write" | "delete" | "rename" | "import" | "backend" | "audit";
}
