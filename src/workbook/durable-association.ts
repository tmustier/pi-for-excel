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
const RETAINED_KEY = "workbook.durable.retained.v1";
const retainedSchema = Type.Array(Type.Object({
  ...associationSchema.properties,
  name: Type.String(),
  aliases: Type.Array(Type.String()),
}));
export type RetainedDurableAssociation = Static<typeof retainedSchema>[number];

export async function listRetainedDurableAssociations(settings: SettingsWriter): Promise<RetainedDurableAssociation[]> {
  const value = await settings.get(RETAINED_KEY);
  if (value == null) return [];
  if (!Value.Check(retainedSchema, value)) throw new Error("Invalid retained Durable associations");
  return value;
}

/** Remove aliases only when they still point at the explicitly deleted generation. */
export async function forgetRetainedDurableAssociation(settings: SettingsWriter, scopeId: string, databaseId: string): Promise<void> {
  const rows = await listRetainedDurableAssociations(settings);
  const row = rows.find((item) => item.scopeId === scopeId && item.databaseId === databaseId);
  if (!row) throw new Error("Retained generation changed");
  for (const alias of row.aliases) {
    const value = await settings.get(`${PREFIX}${alias}`);
    if (Value.Check(associationSchema, value) && value.scopeId === scopeId && value.databaseId === databaseId) {
      await settings.set(`${PREFIX}${alias}`, null);
    }
  }
  await settings.set(RETAINED_KEY, rows.filter((item) => item !== row));
}

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
      if (previous === value) { resolve(true); return; }
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
    const expected = (await listRetainedDurableAssociations(options.settings))
      .find((row) => row.scopeId === association.scopeId)?.databaseId;
    if (expected && association.databaseId && expected !== association.databaseId) {
      throw new Error("Conflicting retained Durable database generation");
    }
    association.databaseId ??= expected ?? null;
    const retain = async (value: Association): Promise<void> => {
      const rows = await listRetainedDurableAssociations(options.settings);
      const previous = rows.find((row) => row.scopeId === value.scopeId);
      if (previous?.databaseId && value.databaseId && previous.databaseId !== value.databaseId) {
        throw new Error("Retained Durable database generation changed");
      }
      // Host display names are not identity. Strip path/query syntax defensively.
      const name = workbook.workbookName?.split(/[?#]/)[0]?.split(/[\\/]/).at(-1) || "Unidentified workbook";
      const row = { ...value, name, aliases: [...new Set([...(previous?.aliases ?? []), scope, value.scopeId])] };
      // A copied document token may predate the catalogue. Keep its known
      // generation discoverable without adopting it or inventing its old name.
      if (candidate?.databaseId && candidate.scopeId !== value.scopeId
        && !rows.some((item) => item.scopeId === candidate.scopeId)) {
        rows.push({ ...candidate, name: "Retained workbook", aliases: [candidate.scopeId,
          ...(candidate.canonicalId ? [candidate.canonicalId] : [])] });
      }
      await options.settings.set(RETAINED_KEY, [...rows.filter((item) => item.scopeId !== value.scopeId), row]);
    };
    await retain(association);
    await options.settings.set(`${PREFIX}${scope}`, association);
    // Also retain the original scope's expectation after a first save alias.
    await options.settings.set(`${PREFIX}${association.scopeId}`, association);
    documentBacked = document ? (document.read() === JSON.stringify(association) || await document.write(JSON.stringify(association))) : false;
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
        await retain(retained);
        await options.settings.set(`${PREFIX}${scope}`, retained);
        await options.settings.set(`${PREFIX}${association.scopeId}`, retained);
        association.databaseId = databaseId;
        if (document) documentBacked = document.read() === JSON.stringify(retained) || await document.write(JSON.stringify(retained));
        last = retained;
      },
    };
  };
}
