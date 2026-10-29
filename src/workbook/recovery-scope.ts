import { getCurrentSpreadsheetHost } from "../host/current.js";
import type { DocumentInstanceIdentity } from "../host/document-instance.js";
import { formatWorkbookLabel, getWorkbookContext, type WorkbookContext } from "./context.js";

export const UNSAVED_WORKBOOK_ID_PREFIX = "doc_instance:";

export interface RecoveryWorkbookScope {
  workbookId: string;
  workbookLabel: string;
  unsaved: boolean;
}

export interface RecoveryScopeResolver {
  resolveForRead(): Promise<RecoveryWorkbookScope | null>;
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
