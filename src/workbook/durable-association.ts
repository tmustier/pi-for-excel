import { Type, type Static } from "typebox";
import { Value } from "typebox/value";
import type { SpreadsheetHost } from "../host/types.js";
import type { SettingsWriter } from "../storage/local/settings-store.js";

const associationSchema = Type.Object({
  scopeId: Type.String({ minLength: 1 }),
  canonicalId: Type.Union([Type.String({ minLength: 1 }), Type.Null()]),
  databaseId: Type.Union([Type.String({ minLength: 1 }), Type.Null()]),
}, { additionalProperties: false });
type Association = Static<typeof associationSchema>;
const DOCUMENT_KEY = "pi.durableWorkbook.v1";
const PREFIX = "workbook.durable.v1.";

/** Independent of the Durable database: a retained generation detects loss. */
export interface DurableDocumentAssociation {
  read(): string | null;
  write(value: string): Promise<boolean>;
}

export function officeDurableDocumentAssociation(): DurableDocumentAssociation | null {
  if (typeof Office === "undefined" || !Office.context?.document?.settings) return null;
  const settings = Office.context.document.settings;
  return {
    read: () => {
      const value: unknown = settings.get(DOCUMENT_KEY);
      if (value === undefined || value === null) return null;
      if (typeof value !== "string") throw new Error("Invalid workbook Durable association");
      return value;
    },
    write: (value) => new Promise((resolve) => {
      const previous: unknown = settings.get(DOCUMENT_KEY);
      settings.set(DOCUMENT_KEY, value);
      settings.saveAsync((result) => {
        const saved = result.status === Office.AsyncResultStatus.Succeeded;
        if (!saved) {
          if (previous === undefined || previous === null) settings.remove(DOCUMENT_KEY);
          else settings.set(DOCUMENT_KEY, previous);
        }
        resolve(saved);
      });
    }),
  };
}

export interface DurableWorkbookAssociation {
  readonly scopeId: string;
  readonly databaseName: string;
  readonly expectedDatabaseId: string | null;
  /** False means loss of all browser-origin data cannot be detected independently. */
  readonly documentBacked: boolean;
  /** Must succeed before admitting work. Never substitutes a new generation. */
  rememberGeneration(databaseId: string): Promise<void>;
}

/** One resolver per pane. No raw URLs, shared "unknown workbook", or save-first gate.
 * Document tokens can be copied by Save As. Only an unclaimed first save adopts
 * the token database; another canonical path receives a separate scope. The old
 * database is retained for explicit history access, never deleted or replayed.
 */
export function createDurableWorkbookAssociationResolver(options: {
  host: Pick<SpreadsheetHost, "kind" | "getWorkbookContext" | "documentInstance">;
  settings: SettingsWriter;
  document?: DurableDocumentAssociation | null;
  createId?: () => string;
}): () => Promise<DurableWorkbookAssociation> {
  const createId = options.createId ?? (() => crypto.randomUUID());
  const paneScope = `pane:${createId()}`;
  const document = options.document ?? null;
  let last: Association | undefined;
  let documentBacked = false;

  return async () => {
    const workbook = await options.host.getWorkbookContext();
    const documentText = document?.read();
    let storedDocument: Association | undefined;
    if (documentText) {
      const decoded: unknown = JSON.parse(documentText);
      if (!Value.Check(associationSchema, decoded)) throw new Error("Invalid workbook Durable association");
      storedDocument = decoded;
    }
    const token = workbook.workbookId === null ? await options.host.documentInstance?.ensure() : null;
    const fallbackScope = options.host.kind === "browser" ? "browser-library" : paneScope;
    const scope = workbook.workbookId ?? (token ? `doc_instance:${token}` : fallbackScope);
    const local: unknown = await options.settings.get(`${PREFIX}${scope}`);
    if (local !== undefined && local !== null && !Value.Check(associationSchema, local)) {
      throw new Error("Invalid local Durable workbook association");
    }
    const localAssociation = Value.Check(associationSchema, local) ? local : undefined;
    const candidate = storedDocument ?? last;
    const sameDocument = candidate && ((candidate.canonicalId === workbook.workbookId
      && (workbook.workbookId !== null || candidate.scopeId === scope))
      || (candidate.canonicalId === null && workbook.workbookId !== null));
    // A saved copy inherits the document token and settings, not ownership of
    // pending work. Its different canonical identity therefore cannot adopt it.
    const inherited = sameDocument ? candidate : undefined;
    if (localAssociation && inherited && (localAssociation.scopeId !== inherited.scopeId
      || (localAssociation.databaseId !== null && inherited.databaseId !== null
        && localAssociation.databaseId !== inherited.databaseId))) {
      throw new Error("Conflicting Durable workbook generations; review retained histories before continuing");
    }
    const association: Association = {
      ...(localAssociation ?? inherited ?? { scopeId: scope, databaseId: null }),
      databaseId: localAssociation?.databaseId ?? inherited?.databaseId ?? null,
      canonicalId: workbook.workbookId,
    };
    await options.settings.set(`${PREFIX}${scope}`, association);
    // Also retain the original scope's expectation after a first save alias.
    await options.settings.set(`${PREFIX}${association.scopeId}`, association);
    documentBacked = document ? await document.write(JSON.stringify(association)) : false;
    last = association;
    return {
      scopeId: association.scopeId,
      databaseName: `pi-for-excel-durable:${association.scopeId}`,
      expectedDatabaseId: association.databaseId,
      get documentBacked() { return documentBacked; },
      rememberGeneration: async (databaseId) => {
        if (association.databaseId !== null && association.databaseId !== databaseId) {
          throw new Error("Durable database generation changed");
        }
        const retained = { ...association, databaseId };
        await options.settings.set(`${PREFIX}${scope}`, retained);
        await options.settings.set(`${PREFIX}${association.scopeId}`, retained);
        if (document) documentBacked = await document.write(JSON.stringify(retained));
        last = retained;
      },
    };
  };
}
