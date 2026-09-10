/**
 * Which workbook a recovery checkpoint belongs to.
 *
 * Saved workbooks use the path-derived `workbookId`. A workbook that has never
 * been saved has no such identity, so checkpoints for it are scoped to a token
 * stored inside the document instead (see `src/host/document-instance.ts`).
 *
 * The two never alias: once a workbook has a path-derived identity, checkpoints
 * taken under its document token are no longer visible. Save As and Save a Copy
 * duplicate the token, and the save boundary already clears checkpoints, so the
 * token only ever needs to cover the stretch before the first save.
 */

import { getCurrentSpreadsheetHost } from "../host/current.js";
import type { DocumentInstanceIdentity } from "../host/document-instance.js";
import { formatWorkbookLabel, getWorkbookContext, type WorkbookContext } from "./context.js";

export const UNSAVED_WORKBOOK_ID_PREFIX = "doc_instance:";

export interface RecoveryWorkbookScope {
  /** Identity checkpoints are stored under and matched against exactly. */
  workbookId: string;
  workbookLabel: string;
  /** True when `workbookId` is a document token rather than a path-derived identity. */
  unsaved: boolean;
}

export interface RecoveryScopeResolver {
  /** Current scope without side effects; `null` when no identity exists yet. */
  resolveForRead(): Promise<RecoveryWorkbookScope | null>;
  /** Current scope, creating a document token for a never-saved workbook when the host can. */
  resolveForAppend(): Promise<RecoveryWorkbookScope | null>;
}

export interface RecoveryScopeResolverDependencies {
  getWorkbookContext: () => Promise<WorkbookContext>;
  getDocumentInstance: () => DocumentInstanceIdentity | null;
}

export function defaultGetDocumentInstance(): DocumentInstanceIdentity | null {
  return getCurrentSpreadsheetHost().documentInstance;
}

export function createRecoveryScopeResolver(
  dependencies: Partial<RecoveryScopeResolverDependencies> = {},
): RecoveryScopeResolver {
  const getContext = dependencies.getWorkbookContext ?? getWorkbookContext;
  const getDocumentInstance = dependencies.getDocumentInstance ?? defaultGetDocumentInstance;

  const resolve = async (mode: "read" | "append"): Promise<RecoveryWorkbookScope | null> => {
    const context = await getContext();
    if (context.workbookId) {
      return { workbookId: context.workbookId, workbookLabel: formatWorkbookLabel(context), unsaved: false };
    }

    const documentInstance = getDocumentInstance();
    if (!documentInstance) return null;

    const token = mode === "append" ? await documentInstance.ensure() : documentInstance.read();
    if (!token) return null;

    return {
      workbookId: `${UNSAVED_WORKBOOK_ID_PREFIX}${token}`,
      workbookLabel: formatWorkbookLabel(context),
      unsaved: true,
    };
  };

  return {
    resolveForRead: () => resolve("read"),
    resolveForAppend: () => resolve("append"),
  };
}
