/**
 * Workspace file subsystem shared types.
 *
 * The persisted records (workbook tags, audit entries) and their enums are
 * `Static<>` of the schemas in `./persisted-schemas.ts`.
 */

import type { WorkspaceBackendKind, WorkspaceFileWorkbookTag } from "./persisted-schemas.js";

export type {
  FilesWorkspaceAuditAction,
  FilesWorkspaceAuditActor,
  FilesWorkspaceAuditEntry,
  WorkspaceBackendKind,
  WorkspaceFileWorkbookTag,
} from "./persisted-schemas.js";

export type WorkspaceFileKind = "text" | "binary";

export type WorkspaceFileSourceKind = "workspace" | "builtin-doc";

export type WorkspaceFileLocationKind = "workspace" | "native-directory" | "builtin-doc";

export interface WorkspaceFileEntry {
  path: string;
  name: string;
  size: number;
  modifiedAt: number;
  mimeType: string;
  kind: WorkspaceFileKind;
  sourceKind: WorkspaceFileSourceKind;
  locationKind?: WorkspaceFileLocationKind;
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

export const FILES_WORKSPACE_CHANGED_EVENT = "pi:files-workspace-changed";

export interface FilesWorkspaceChangedDetail {
  reason: "write" | "delete" | "rename" | "import" | "backend" | "audit";
}
