import { type Static, Type } from "typebox";

import { normalizeWorkspacePath } from "./path.js";

export const WorkspaceBackendKindSchema = Type.Union([
  Type.Literal("native-directory"),
  Type.Literal("opfs"),
  Type.Literal("memory"),
]);

export const FilesWorkspaceAuditActorSchema = Type.Union([
  Type.Literal("assistant"),
  Type.Literal("user"),
  Type.Literal("system"),
]);

export const FilesWorkspaceAuditActionSchema = Type.Union([
  Type.Literal("list"),
  Type.Literal("read"),
  Type.Literal("write"),
  Type.Literal("delete"),
  Type.Literal("rename"),
  Type.Literal("import"),
  Type.Literal("connect_native"),
  Type.Literal("disconnect_native"),
  Type.Literal("clear_audit"),
]);

const NonEmptyStringSchema = Type.String({ minLength: 1 });

function isNormalizedWorkspacePath(value: string): boolean {
  try {
    return normalizeWorkspacePath(value) === value;
  } catch {
    return false;
  }
}

export const NormalizedWorkspacePathSchema = Type.Refine(
  Type.String(),
  isNormalizedWorkspacePath,
  () => "path must be a normalized workspace path",
);

export const WorkspaceFileWorkbookTagSchema = Type.Object({
  workbookId: NonEmptyStringSchema,
  workbookLabel: NonEmptyStringSchema,
  taggedAt: Type.Number(),
});

export const FilesWorkspaceAuditEntrySchema = Type.Object({
  id: NonEmptyStringSchema,
  at: Type.Number(),
  action: FilesWorkspaceAuditActionSchema,
  actor: FilesWorkspaceAuditActorSchema,
  source: NonEmptyStringSchema,
  backend: WorkspaceBackendKindSchema,
  path: Type.Optional(NormalizedWorkspacePathSchema),
  fromPath: Type.Optional(NormalizedWorkspacePathSchema),
  toPath: Type.Optional(NormalizedWorkspacePathSchema),
  bytes: Type.Optional(Type.Number()),
  workbookId: Type.Optional(NonEmptyStringSchema),
  workbookLabel: Type.Optional(NonEmptyStringSchema),
});

export const PersistedWorkspaceMetadataSchema = Type.Object({
  version: Type.Literal(1),
  byPath: Type.Record(Type.String(), Type.Unknown()),
});

export const PersistedAuditTrailSchema = Type.Object({
  version: Type.Literal(1),
  entries: Type.Array(Type.Unknown()),
});

export type WorkspaceBackendKind = Static<typeof WorkspaceBackendKindSchema>;
export type FilesWorkspaceAuditActor = Static<typeof FilesWorkspaceAuditActorSchema>;
export type FilesWorkspaceAuditAction = Static<typeof FilesWorkspaceAuditActionSchema>;
export type WorkspaceFileWorkbookTag = Static<typeof WorkspaceFileWorkbookTagSchema>;
export type FilesWorkspaceAuditEntry = Static<typeof FilesWorkspaceAuditEntrySchema>;
export type PersistedWorkspaceMetadata = Static<typeof PersistedWorkspaceMetadataSchema>;
export type PersistedAuditTrail = Static<typeof PersistedAuditTrailSchema>;
