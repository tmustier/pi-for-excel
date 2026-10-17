import type { Context, JsonValue } from "@earendil-works/chord";
import { BACKGROUND_CONTEXT } from "@earendil-works/chord/context";
import { applyImmutableBatches } from "@earendil-works/chord/delta";
import { StorageRequestError } from "@earendil-works/pi-durable";
import type {
  ConversationId, ConversationQuery, ConversationRecord, Cursor, DocumentAddress, DocumentContent,
  DocumentCreate, DocumentId, DocumentPoint, DocumentQuery, DocumentRecord, EntryId, EntryQuery,
  EntryRecord, Id, Page, ScanOrder, Seq, Storage, StorageWrite, StoredDocument, SubmissionId,
  SubmissionQuery, SubmissionRecord, TaskId, TaskQuery, TaskRecord,
} from "@earendil-works/pi-durable";
import { Type, type Static, type TSchema } from "typebox";
import { Value } from "typebox/value";
import {
  completed, digest, counterSchema, conversationSchema, entrySchema, taskSchema,
  submissionSchema, documentSchema, contentSchema, journalSchema, writesSchema,
} from "./storage-codec";
import { acquireStorageLock, DurableStorageError, requestStoragePersistence } from "./storage-lifecycle";

export { DurableStorageError, deleteDurableWorkbookStorage } from "./storage-lifecycle";

const STORES = ["meta", "journal", "records", "revisions", "requests"];
// Physical batching policies, never limits on history or on returned content.
const CHECKPOINT_INTERVAL = 64;
const READ_PAGE = 64;
const natural = Type.Integer({ minimum: 0, maximum: Number.MAX_SAFE_INTEGER });
const manifestSchema = Type.Object({
  sequence: natural, highWatermark: counterSchema, records: natural, revisions: natural,
  checkpoint: natural, requests: natural,
});
type Manifest = Static<typeof manifestSchema>;
const envelopeSchema = Type.Object({ text: Type.String(), digest: Type.String() });
type Envelope = Static<typeof envelopeSchema>;
const storedRowSchema = Type.Union([envelopeSchema, journalSchema]);
type StoredRow = Static<typeof storedRowSchema>;
type TableWrite = Extract<StorageWrite, { type: "conversation" | "entry" | "task" | "submission" }>;
type RecordRow = { seq: Seq; write: TableWrite | { type: "document"; value: DocumentRecord } };
type TableValues = {
  conversation: ConversationRecord; entry: EntryRecord; document: DocumentRecord;
  task: TaskRecord<JsonValue, JsonValue, JsonValue>; submission: SubmissionRecord;
};
const recordSchema = Type.Unsafe<RecordRow>(Type.Object({ seq: counterSchema, write: Type.Union([
  Type.Object({ type: Type.Literal("conversation"), value: conversationSchema }),
  Type.Object({ type: Type.Literal("entry"), value: entrySchema }),
  Type.Object({ type: Type.Literal("task"), value: taskSchema }),
  Type.Object({ type: Type.Literal("submission"), value: submissionSchema }),
  Type.Object({ type: Type.Literal("document"), value: Type.Intersect([
    documentSchema, Type.Object({ createdAt: counterSchema, retiredAt: Type.Optional(counterSchema) }),
  ]) }),
]) }));
type Revision = { id: DocumentId; seq: Seq; deltasSinceBase: number; content: DocumentContent };
const revisionSchema = Type.Unsafe<Revision>(Type.Object({
  id: counterSchema, seq: counterSchema, deltasSinceBase: natural, content: contentSchema,
}));
const requestSchema = Type.Object({ key: Type.String(), id: counterSchema });

function corrupt(message: string): DurableStorageError {
  return new DurableStorageError("corrupt", message);
}

async function encode<T>(value: T): Promise<Envelope> {
  const text = JSON.stringify(value);
  return { text, digest: await digest(text) };
}

async function decode<S extends TSchema>(schema: S, raw: StoredRow | undefined): Promise<Static<S>> {
  if (!Value.Check(envelopeSchema, raw) || await digest(raw.text) !== raw.digest) {
    throw corrupt("Durable storage checksum mismatch");
  }
  let parsed: unknown;
  try { parsed = JSON.parse(raw.text); } catch { throw corrupt("Invalid durable storage JSON"); }
  if (!Value.Check(schema, parsed)) throw corrupt("Invalid durable storage record");
  return parsed;
}

function scopeKey(scope: DocumentRecord["scope"]): string {
  return JSON.stringify(scope.kind === "session" ? [scope.kind]
    : [scope.kind, scope.kind === "conversation" ? scope.conversationId : scope.taskId]);
}

function addressKey(address: DocumentAddress): string {
  return JSON.stringify([address.kind, scopeKey(address.scope),
    address.key === undefined ? ["singleton"] : ["family", address.key]]);
}

function alive(record: DocumentRecord, at: DocumentPoint): boolean {
  return at === "current" ? record.retiredAt === undefined
    : record.createdAt <= at && (record.retiredAt === undefined || at < record.retiredAt);
}

function currentOnly(record: DocumentRecord): boolean {
  return record.scope.kind !== "conversation" || record.history === "latest";
}

function scanStart(order: ScanOrder | undefined, cursor: Cursor | undefined, fallback: ScanOrder): {
  order: ScanOrder; after: number | undefined;
} {
  if (order !== undefined && order !== "ascending" && order !== "descending") throw new StorageRequestError("Invalid scan order");
  if (cursor === undefined) return { order: order ?? fallback, after: undefined };
  const stored = cursor.order ?? fallback;
  if (typeof cursor.after !== "number" || !Number.isSafeInteger(cursor.after)
    || (stored !== "ascending" && stored !== "descending")) throw new StorageRequestError("Invalid storage cursor");
  if (order !== undefined && order !== stored) throw new StorageRequestError(`The cursor continues a ${stored} scan; the query asks for ${order}`);
  return { order: stored, after: cursor.after };
}

function page<T extends { id: number }>(items: T[], limit: number, order: ScanOrder): Page<T, Cursor> {
  if (items.length <= limit) return { items };
  items.pop();
  const last = items.at(-1);
  return last ? { items, next: { after: last.id, order } } : { items };
}

/** The public Storage contract plus explicit lifecycle information for its host. */
export interface DurableIndexedDbStorage extends Storage {
  readonly databaseId: string;
  readonly persistence: "persistent" | "best-effort" | "unsupported";
  checkpoint(): Promise<{ sequence: number; pruned: number }>;
}

/** No complete-database mirror: memory is a page, a commit batch, or one materialized document. */
class IndexedDbStorage implements DurableIndexedDbStorage {
  private tail: Promise<void> = Promise.resolve();
  private closing = false;
  private shutdown: Promise<void> | undefined;
  private readonly db: IDBDatabase;
  private readonly release: () => Promise<void>;
  readonly databaseId: string;
  readonly persistence: DurableIndexedDbStorage["persistence"];
  private manifest: Manifest;

  constructor(
    db: IDBDatabase,
    release: () => Promise<void>,
    databaseId: string,
    persistence: DurableIndexedDbStorage["persistence"],
    manifest: Manifest,
  ) {
    this.db = db;
    this.release = release;
    this.databaseId = databaseId;
    this.persistence = persistence;
    this.manifest = manifest;
    db.onversionchange = () => { void this.close(); };
    window.addEventListener("pagehide", this.pagehide);
  }

  private readonly pagehide = (event: PageTransitionEvent): void => {
    if (event.persisted) window.addEventListener("pageshow", () => window.location.reload(), { once: true });
    void this.close();
  };

  private serialize<T>(operation: () => Promise<T>, context = BACKGROUND_CONTEXT): Promise<T> {
    if (this.closing) return Promise.reject(new Error("Durable storage is closed"));
    const result = this.tail.then(async () => {
      context.abortSignal?.throwIfAborted();
      const value = await operation();
      context.abortSignal?.throwIfAborted();
      return value;
    });
    this.tail = result.then(() => {}, () => {});
    return result;
  }

  private async get<S extends TSchema>(store: string, key: IDBValidKey, schema: S): Promise<Static<S> | undefined> {
    const tx = this.db.transaction(store, "readonly");
    const done = completed(tx);
    const request = tx.objectStore(store).get(key);
    await done;
    const raw: unknown = request.result;
    if (raw === undefined) return undefined;
    if (!Value.Check(schema, raw)) throw corrupt("Invalid Durable storage row");
    return raw;
  }

  private async record(id: number): Promise<RecordRow | undefined> {
    const raw = await this.get("records", id, envelopeSchema);
    if (raw === undefined) return undefined;
    const row = await decode(recordSchema, raw);
    if (row.write.value.id !== id || row.seq > this.manifest.sequence) throw corrupt("Invalid durable record identity");
    return row;
  }

  /** Bounded cursor page; never getAll() over a user-sized history. */
  private async rows(store: string, index: string | undefined, range: IDBKeyRange | undefined,
    direction: IDBCursorDirection, limit = READ_PAGE): Promise<{ key: IDBValidKey; value: StoredRow }[]> {
    const tx = this.db.transaction(store, "readonly");
    const done = completed(tx);
    const source = index ? tx.objectStore(store).index(index) : tx.objectStore(store);
    const request = source.openCursor(range, direction);
    const rows: { key: IDBValidKey; value: StoredRow }[] = [];
    let invalid = false;
    request.onsuccess = () => {
      const cursor = request.result;
      if (!cursor) return;
      const raw: unknown = cursor.value;
      if (!Value.Check(storedRowSchema, raw)) { invalid = true; return; }
      rows.push({ key: cursor.key, value: raw });
      if (rows.length < limit) cursor.continue();
    };
    await done;
    if (invalid) throw corrupt("Invalid Durable storage cursor row");
    return rows;
  }

  private async *records(type: RecordRow["write"]["type"], order: ScanOrder, after?: number): AsyncGenerator<RecordRow> {
    if (after !== undefined && ((order === "ascending" && after >= Number.MAX_SAFE_INTEGER)
      || (order === "descending" && after <= 0))) return;
    let range = order === "ascending"
      ? IDBKeyRange.bound([type, after ?? 0], [type, Number.MAX_SAFE_INTEGER], after !== undefined)
      : IDBKeyRange.bound([type, 0], [type, after ?? Number.MAX_SAFE_INTEGER], false, after !== undefined);
    while (true) {
      const rows = await this.rows("records", "type", range, order === "ascending" ? "next" : "prev");
      for (const raw of rows) {
        const row = await decode(recordSchema, raw.value);
        if (row.write.type !== type || JSON.stringify(raw.key) !== JSON.stringify([type, row.write.value.id])) {
          throw corrupt("Invalid durable record index");
        }
        yield row;
      }
      const last = rows.at(-1);
      if (!last || rows.length < READ_PAGE) return;
      if (indexedDB.cmp(last.key, [type, order === "ascending" ? Number.MAX_SAFE_INTEGER : 0]) === 0) return;
      range = order === "ascending" ? IDBKeyRange.bound(last.key, [type, Number.MAX_SAFE_INTEGER], true)
        : IDBKeyRange.bound([type, 0], last.key, false, true);
    }
  }

  private async tablePage<T extends RecordRow["write"]["type"]>(type: T, order: ScanOrder | undefined,
    limit: number, cursor: Cursor | undefined,
    predicate: (value: TableValues[T]) => boolean,
  ): Promise<Page<TableValues[T], Cursor>> {
    if (!Number.isSafeInteger(limit) || limit < 1) throw new StorageRequestError("Invalid scan limit");
    const start = scanStart(order, cursor, "ascending");
    const values: TableValues[T][] = [];
    for await (const row of this.records(type, start.order, start.after)) {
      // records verifies the discriminant before yielding; TS cannot narrow a generic discriminant.
      const value = row.write.value as TableValues[T];
      if (predicate(value)) values.push(value);
      if (values.length > limit) break;
    }
    return page(values, limit, start.order);
  }

  conversation(id: ConversationId, context: Context): Promise<ConversationRecord | undefined> {
    return this.serialize(async () => { const row = await this.record(id); return row?.write.type === "conversation" ? row.write.value : undefined; }, context);
  }

  scanConversations(query: ConversationQuery, limit: number, cursor: Cursor | undefined, context: Context): Promise<Page<ConversationRecord, Cursor>> {
    return this.serialize(() => this.tablePage("conversation", query.order, limit, cursor, (v) =>
      (query.ownerConversationId === undefined || v.owner?.conversationId === query.ownerConversationId)
      && (query.ownerTaskId === undefined || v.owner?.taskId === query.ownerTaskId)), context);
  }

  private async *ancestry(id: ConversationId, upper: number): AsyncGenerator<{ id: ConversationId; upper: number }> {
    // Session supplies acyclic ancestry. Reject corrupt cycles using monotonic fork cutoffs' IDs is
    // insufficient (IDs may be inserted out of order), so use a visited set of ancestry only.
    const visited = new Set<ConversationId>();
    while (true) {
      if (visited.has(id)) throw corrupt("Cyclic durable conversation ancestry");
      visited.add(id);
      const row = await this.record(id);
      if (row?.write.type !== "conversation") throw new StorageRequestError(`Unknown conversation: ${id}`);
      yield { id, upper };
      const parent = row.write.value.parent;
      if (!parent) return;
      upper = Math.min(upper, parent.at);
      id = parent.conversationId;
    }
  }

  private async *visible(query: EntryQuery, order: ScanOrder, after?: number): AsyncGenerator<EntryRecord> {
    const min = Math.max(query.minEntryId ?? 0, order === "ascending" && after !== undefined ? after + 1 : 0);
    const max = Math.min(query.maxEntryId ?? Number.MAX_SAFE_INTEGER,
      order === "descending" && after !== undefined ? after - 1 : Number.MAX_SAFE_INTEGER);
    const segments = [];
    for await (const segment of this.ancestry(query.conversationId, max)) segments.push(segment);
    if (order === "ascending") segments.reverse();
    for (const segment of segments) {
      if (min > segment.upper) continue;
      let range = IDBKeyRange.bound([segment.id, min], [segment.id, segment.upper]);
      while (true) {
        const rows = await this.rows("records", "entries", range, order === "ascending" ? "next" : "prev");
        for (const raw of rows) {
          const row = await decode(recordSchema, raw.value);
          if (row.write.type !== "entry" || row.write.value.conversationId !== segment.id
            || JSON.stringify(raw.key) !== JSON.stringify([segment.id, row.write.value.id])) throw corrupt("Invalid durable entry index");
          yield row.write.value;
        }
        const last = rows.at(-1);
        if (!last || rows.length < READ_PAGE) break;
        if (indexedDB.cmp(last.key, [segment.id, order === "ascending" ? segment.upper : min]) === 0) break;
        range = order === "ascending" ? IDBKeyRange.bound(last.key, [segment.id, segment.upper], true)
          : IDBKeyRange.bound([segment.id, min], last.key, false, true);
      }
    }
  }

  entry(id: EntryId, context: Context): Promise<{ entry: EntryRecord; commitSeq: Seq } | undefined>;
  entry(conversationId: ConversationId, id: EntryId, context: Context): Promise<{ entry: EntryRecord; commitSeq: Seq } | undefined>;
  entry(id: EntryId | ConversationId, idOrContext: EntryId | Context, context?: Context): Promise<{ entry: EntryRecord; commitSeq: Seq } | undefined> {
    return this.serialize(async () => {
      if (context !== undefined && typeof idOrContext !== "number") throw new StorageRequestError("Storage.entry() requires an entry ID");
      const entryId = typeof idOrContext === "number" ? idOrContext : id;
      if (context !== undefined) {
        let found = false;
        // Both IDs are validated by the caller's branded overload.
        for await (const segment of this.ancestry(id as ConversationId, entryId)) {
          const row = await this.record(entryId);
          if (row?.write.type === "entry" && row.write.value.conversationId === segment.id && entryId <= segment.upper) { found = true; break; }
        }
        if (!found) return undefined;
      }
      const row = await this.record(entryId);
      return row?.write.type === "entry" ? { entry: row.write.value, commitSeq: row.seq } : undefined;
    }, context ?? (typeof idOrContext === "number" ? BACKGROUND_CONTEXT : idOrContext));
  }

  scanEntries(query: EntryQuery, limit: number, cursor: Cursor | undefined, context: Context): Promise<Page<EntryRecord, Cursor>> {
    return this.serialize(async () => {
      if (!Number.isSafeInteger(limit) || limit < 1) throw new StorageRequestError("Invalid scan limit");
      const start = scanStart(query.order, cursor, "descending");
      const items = [];
      for await (const entry of this.visible(query, start.order, start.after)) {
        items.push(entry);
        if (items.length > limit) break;
      }
      return page(items, limit, start.order);
    }, context);
  }

  findLatestHeadMarker(conversationId: ConversationId, at: EntryId | undefined, context: Context): Promise<(EntryRecord & { readonly head: EntryId }) | undefined> {
    return this.serialize(async () => {
      for await (const entry of this.visible({ conversationId, ...(at === undefined ? {} : { maxEntryId: at }) }, "descending")) {
        if (entry.head !== undefined) return { ...entry, head: entry.head };
      }
      return undefined;
    }, context);
  }

  task(id: TaskId, context: Context): Promise<TaskRecord<JsonValue, JsonValue, JsonValue> | undefined> {
    return this.serialize(async () => { const row = await this.record(id); return row?.write.type === "task" ? row.write.value : undefined; }, context);
  }

  scanTasks(query: TaskQuery, limit: number, cursor: Cursor | undefined, context: Context): Promise<Page<TaskRecord<JsonValue, JsonValue, JsonValue>, Cursor>> {
    return this.serialize(() => this.tablePage("task", query.order, limit, cursor, (v) =>
      (query.conversationId === undefined || v.conversationId === query.conversationId)
      && (query.kind === undefined || v.kind === query.kind) && (query.status === undefined || v.state.status === query.status)
      && (query.abortRequested === undefined || v.abortRequested === query.abortRequested)
      && (query.background === undefined || v.background === query.background)), context);
  }

  submission(id: SubmissionId, context: Context): Promise<SubmissionRecord | undefined> {
    return this.serialize(async () => { const row = await this.record(id); return row?.write.type === "submission" ? row.write.value : undefined; }, context);
  }

  scanSubmissions(query: SubmissionQuery, limit: number, cursor: Cursor | undefined, context: Context): Promise<Page<SubmissionRecord, Cursor>> {
    return this.serialize(() => this.tablePage("submission", query.order, limit, cursor, (v) =>
      (query.conversationId === undefined || v.conversationId === query.conversationId)
      && (query.status === undefined || v.status === query.status)), context);
  }

  submissionByRequest(conversationId: ConversationId, requestId: string, context: Context): Promise<SubmissionRecord | undefined> {
    return this.serialize(async () => {
      const id = await this.requestOwner(JSON.stringify([conversationId, requestId]));
      if (id === undefined) return undefined;
      const row = await this.record(id);
      if (row?.write.type !== "submission" || row.write.value.requestId !== requestId || row.write.value.conversationId !== conversationId) throw corrupt("Invalid submission index");
      return row.write.value;
    }, context);
  }

  private async requestOwner(key: string): Promise<number | undefined> {
    const raw = await this.get("requests", key, envelopeSchema);
    if (raw === undefined) return undefined;
    const pointer = await decode(requestSchema, raw);
    if (pointer.key !== key) throw corrupt("Invalid submission request identity");
    return pointer.id;
  }

  private async find(address: DocumentAddress, at: DocumentPoint): Promise<DocumentRecord | undefined> {
    const key = addressKey(address);
    let range = IDBKeyRange.bound([key, 0], [key, Number.MAX_SAFE_INTEGER]);
    while (true) {
      const rows = await this.rows("records", "address", range, "next");
      for (const raw of rows) {
        const row = await decode(recordSchema, raw.value);
        if (row.write.type !== "document" || addressKey(row.write.value) !== key) throw corrupt("Invalid document address index");
        if (alive(row.write.value, at)) return row.write.value;
      }
      const last = rows.at(-1);
      if (!last || rows.length < READ_PAGE) return undefined;
      if (indexedDB.cmp(last.key, [key, Number.MAX_SAFE_INTEGER]) === 0) return undefined;
      range = IDBKeyRange.bound(last.key, [key, Number.MAX_SAFE_INTEGER], true);
    }
  }

  findDocument(address: DocumentAddress, at: DocumentPoint, context: Context): Promise<DocumentRecord | undefined> {
    return this.serialize(() => this.find(address, at), context);
  }

  private async materialize(id: DocumentId, at: DocumentPoint): Promise<StoredDocument | undefined> {
    const row = await this.record(id);
    if (row?.write.type !== "document") return undefined;
    const record = row.write.value;
    if (at !== "current" && currentOnly(record)) throw new StorageRequestError(`Document ${id} does not retain historical content`);
    if (!alive(record, at)) return undefined;
    const upper = at === "current" ? this.manifest.sequence : at;
    const bases = await this.rows("revisions", "base", IDBKeyRange.bound([id, 0], [id, upper]), "prev", 1);
    const rawBase = bases[0];
    if (!rawBase) throw corrupt(`Document ${id} is missing a required base`);
    const base = await decode(revisionSchema, rawBase.value);
    if (base.id !== id || base.content.kind !== "base" || base.seq > upper) throw corrupt("Invalid document base");
    let value = base.content.value;
    let deltasSinceBase = base.deltasSinceBase;
    let lastSeq = base.seq;
    while (lastSeq < upper) {
      const rows = await this.rows("revisions", undefined, IDBKeyRange.bound([id, lastSeq], [id, upper], true), "next");
      for (const raw of rows) {
        const revision = await decode(revisionSchema, raw.value);
        if (revision.id !== id || revision.seq <= lastSeq || revision.content.kind !== "delta"
          || revision.content.version !== base.content.version || revision.deltasSinceBase !== deltasSinceBase + 1) throw corrupt("Invalid document revision chain");
        value = applyImmutableBatches(value, [revision.content.ops]);
        deltasSinceBase = revision.deltasSinceBase;
        lastSeq = revision.seq;
      }
      if (rows.length < READ_PAGE) break;
    }
    return { record, version: base.content.version, value, deltasSinceBase };
  }

  document(id: DocumentId, at: DocumentPoint, context: Context): Promise<StoredDocument | undefined> {
    return this.serialize(() => this.materialize(id, at), context);
  }

  scanDocuments(query: DocumentQuery, limit: number, cursor: Cursor | undefined, context: Context): Promise<Page<DocumentRecord, Cursor>> {
    return this.serialize(() => this.tablePage("document", "ascending", limit, cursor, (v) =>
      scopeKey(v.scope) === scopeKey(query.scope) && (query.kind === undefined || v.kind === query.kind) && alive(v, query.at)), context);
  }

  async commit(writes: readonly StorageWrite[], _context: Context): Promise<Seq> {
    const detached = structuredClone(writes);
    return this.serialize(() => this.persist(detached));
  }

  private async persist(input: readonly StorageWrite[], migrationSeq?: number): Promise<Seq> {
    const sequence = migrationSeq ?? this.manifest.sequence + 1;
    if (!Number.isSafeInteger(sequence) || sequence <= this.manifest.sequence) throw new Error("Durable sequence space is exhausted");
    // Safe-integer check above establishes the erased sequence brand.
    const seq = sequence as Seq;
    const writes: StorageWrite[] = [];
    const touched = new Set(input.filter((w) => w.type.startsWith("document.")).map((w) =>
      w.type === "document.create" || w.type === "document.copy" ? w.record.id
        : w.type === "document.change" || w.type === "document.retire" ? w.id : w.value.id));
    for (const write of input) {
      if (write.type !== "document.copy") { writes.push(write); continue; }
      try {
        if (touched.has(write.source.id)) throw new Error("Fork source document is changed in the copy batch");
        const source = await this.materialize(write.source.id, write.source.at);
        if (!source || source.record.scope.kind !== "conversation" || write.record.scope.kind !== "conversation"
          || source.record.kind !== write.record.kind || source.record.key !== write.record.key
          || source.record.history !== write.record.history || source.record.fork !== write.record.fork) throw new Error("Fork source does not match copied record");
        writes.push({ type: "document.create", record: write.record, content: { kind: "base", version: source.version, value: source.value } });
      } catch (cause) { throw new Error(`Document copy ${write.record.id} was rejected`, { cause }); }
    }
    if (!Value.Check(writesSchema, writes)) throw new Error("Invalid durable writes");
    const changed = new Map<number, RecordRow>();
    const claims = new Map<number, RecordRow["write"]["type"]>();
    const actions = new Map<DocumentId, { create?: DocumentCreate; content?: DocumentContent; retire: boolean }>();
    const requests = new Map<string, { before: number | undefined; after: number | undefined }>();
    const request = async (key: string): Promise<{ before: number | undefined; after: number | undefined }> => {
      let state = requests.get(key);
      if (!state) {
        const owner = await this.requestOwner(key);
        state = { before: owner, after: owner };
        requests.set(key, state);
      }
      return state;
    };
    let additions = 0;
    let highWatermark = this.manifest.highWatermark;
    for (const write of writes) {
      if (write.type !== "document.change" && write.type !== "document.retire") {
        if (write.type === "document.copy") throw new Error("Unresolved document copy");
        const type = write.type === "document.create" ? "document" : write.type;
        const id = write.type === "document.create" ? write.record.id : write.value.id;
        const existing = await this.record(id);
        const earlier = claims.get(id);
        if (existing && (existing.write.type !== type || type === "conversation" || type === "entry" || type === "document")) throw new Error(`ID ${id} already belongs to ${existing.write.type}`);
        if (earlier && (earlier !== type || type === "conversation" || type === "entry" || type === "document")) throw new Error(`ID ${id} is written more than once or as two record types`);
        if (!existing && !earlier) additions++;
        claims.set(id, type);
        highWatermark = Math.max(highWatermark, id);
      }
      if (write.type === "submission") {
        const previous = changed.get(write.value.id) ?? await this.record(write.value.id);
        if (previous?.write.type === "submission" && previous.write.value.requestId !== undefined) {
          const prior = previous.write.value;
          const pointer = await request(JSON.stringify([prior.conversationId, prior.requestId]));
          if (pointer.after === prior.id) pointer.after = undefined;
        }
        if (write.value.requestId !== undefined) {
          const pointer = await request(JSON.stringify([write.value.conversationId, write.value.requestId]));
          pointer.after = write.value.id;
        }
      }
      if (write.type === "conversation" || write.type === "entry" || write.type === "task" || write.type === "submission") {
        changed.set(write.value.id, { seq, write });
      } else {
        const id = write.type === "document.create" ? write.record.id : write.id;
        const action = actions.get(id) ?? { retire: false };
        if (write.type === "document.retire") {
          if (action.retire) throw new Error(`Document ${id} is retired more than once`);
          action.retire = true;
        } else {
          if (action.content) throw new Error(`Document ${id} has more than one content command`);
          action.content = write.content;
          if (write.type === "document.create") action.create = write.record;
        }
        actions.set(id, action);
      }
    }
    const liveCounts = new Map<string, number>();
    const revisions: Revision[] = [];
    for (const [id, action] of actions) {
      const existing = await this.record(id);
      const previous = existing?.write.type === "document" ? existing.write.value : undefined;
      if (!action.create && !previous) throw new Error(`Unknown document: ${id}`);
      if (previous?.retiredAt !== undefined) throw new Error(`Document ${id} is retired`);
      const record = action.create ? { ...action.create, createdAt: seq } : previous;
      if (!record) throw new Error(`Unknown document: ${id}`);
      const key = addressKey(record);
      const current = await this.find(record, "current");
      let live = liveCounts.get(key) ?? (current ? 1 : 0);
      if (action.retire && current?.id === id) live--;
      if (action.create && !action.retire) live++;
      liveCounts.set(key, live);
      changed.set(id, { seq, write: { type: "document", value: action.retire ? { ...record, retiredAt: seq } : record } });
      if (action.content) {
        let content = action.content;
        let deltasSinceBase = 0;
        if (content.kind === "delta") {
          const latest = (await this.rows("revisions", undefined,
            IDBKeyRange.bound([id, 0], [id, this.manifest.sequence]), "prev", 1))[0];
          if (!latest) throw corrupt(`Document ${id} delta has no base`);
          const revision = await decode(revisionSchema, latest.value);
          if (revision.id !== id || JSON.stringify(latest.key) !== JSON.stringify([id, revision.seq])) throw corrupt("Invalid document revision identity");
          if (revision.content.version !== content.version) throw new Error(`Document ${id} version transition requires a base`);
          deltasSinceBase = revision.deltasSinceBase + 1;
          // Physical checkpoint does not reset the Session's logical delta count.
          if (deltasSinceBase % CHECKPOINT_INTERVAL === 0) {
            const stored = await this.materialize(id, "current");
            if (!stored) throw corrupt(`Document ${id} delta has no base`);
            content = { kind: "base", version: content.version, value: applyImmutableBatches(stored.value, [content.ops]) };
          }
        }
        revisions.push({ id, seq, content, deltasSinceBase });
      }
    }
    for (const live of liveCounts.values()) if (live > 1) throw new Error("Document address already has a current incarnation");
    const next = { ...this.manifest, sequence, highWatermark, records: this.manifest.records + additions,
      revisions: this.manifest.revisions + revisions.length };
    const encodedRequests = [];
    for (const [key, pointer] of requests) {
      next.requests += Number(pointer.after !== undefined) - Number(pointer.before !== undefined);
      encodedRequests.push({ key, value: pointer.after === undefined ? undefined : await encode({ key, id: pointer.after }) });
    }
    const encodedRecords = [];
    for (const row of changed.values()) {
      const write = row.write;
      encodedRecords.push({ ...await encode(row), id: write.value.id, type: [write.type, write.value.id],
        ...(write.type === "entry" ? { entries: [write.value.conversationId, write.value.id] } : {}),
        ...(write.type === "document" ? { address: [addressKey(write.value), write.value.id] } : {}),
      });
    }
    const encodedRevisions = [];
    for (const revision of revisions) encodedRevisions.push({ ...await encode(revision), key: [revision.id, seq],
      ...(revision.content.kind === "base" ? { base: [revision.id, seq] } : {}),
    });
    const text = JSON.stringify(writes);
    const journal = { seq, writes: text, digest: await digest(`${seq}:${text}`) };
    const prune = migrationSeq !== undefined || sequence - next.checkpoint >= CHECKPOINT_INTERVAL;
    if (prune) next.checkpoint = sequence;
    const metadata = await encode(next);
    const tx = this.db.transaction(STORES, "readwrite", { durability: "strict" });
    const done = completed(tx);
    try {
      for (const pointer of encodedRequests) {
        if (pointer.value === undefined) tx.objectStore("requests").delete(pointer.key);
        else tx.objectStore("requests").put(pointer.value, pointer.key);
      }
      for (const row of encodedRecords) tx.objectStore("records").put(row);
      for (const row of encodedRevisions) tx.objectStore("revisions").add(row);
      if (migrationSeq === undefined) tx.objectStore("journal").add(journal);
      if (prune) tx.objectStore("journal").delete(IDBKeyRange.upperBound(sequence));
      tx.objectStore("meta").put(metadata, "manifest");
    } catch (error) { tx.abort(); await done.catch(() => {}); throw error; }
    await done;
    this.manifest = next;
    return seq;
  }

  mintId<I extends Id<string>>(): Promise<I> {
    return this.serialize(async () => {
      const id = this.manifest.highWatermark + 1;
      if (!Number.isSafeInteger(id)) throw new Error("Durable ID space is exhausted");
      const next = { ...this.manifest, highWatermark: id };
      const encoded = await encode(next);
      const tx = this.db.transaction("meta", "readwrite", { durability: "strict" });
      const done = completed(tx);
      tx.objectStore("meta").put(encoded, "manifest");
      await done;
      this.manifest = next;
      // Reservation has committed and the safe integer invariant was checked above.
      return id as I;
    });
  }

  checkpoint(): Promise<{ sequence: number; pruned: number }> {
    return this.serialize(async () => {
      await this.verifyJournal();
      const next = { ...this.manifest, checkpoint: this.manifest.sequence };
      const encoded = await encode(next);
      const tx = this.db.transaction(["meta", "journal"], "readwrite", { durability: "strict" });
      const done = completed(tx);
      const count = tx.objectStore("journal").count();
      tx.objectStore("journal").clear();
      tx.objectStore("meta").put(encoded, "manifest");
      await done;
      this.manifest = next;
      return { sequence: next.sequence, pruned: count.result };
    });
  }

  private async verifyJournal(): Promise<void> {
    let last = this.manifest.checkpoint;
    while (true) {
      const rows = await this.rows("journal", undefined, IDBKeyRange.lowerBound(last, true), "next", 1);
      const raw = rows[0];
      if (!raw) break;
      if (!Value.Check(journalSchema, raw.value)) throw corrupt("Invalid durable journal");
      const row = raw.value;
      if (row.seq !== last + 1 || row.seq !== raw.key || await digest(`${row.seq}:${row.writes}`) !== row.digest) throw corrupt("Corrupt durable journal");
      last = row.seq;
    }
    if (last !== this.manifest.sequence) throw corrupt("Incomplete durable journal");
  }

  initialize(): Promise<void> {
    return this.serialize(() => this.restore());
  }

  private async restore(): Promise<void> {
    const migration = await this.get("meta", "migration", Type.Literal(true));
    if (migration === true) {
      const rawSequence = await this.get("meta", "sequence", counterSchema);
      const target = rawSequence ?? 0;
      const watermark = await this.get("meta", "highWatermark", counterSchema);
      if (target < this.manifest.sequence) throw corrupt("Legacy migration sequence regressed");
      while (this.manifest.sequence < target) {
        const rows = await this.rows("journal", undefined, IDBKeyRange.lowerBound(this.manifest.sequence, true), "next", 1);
        const raw = rows[0];
        if (!raw || !Value.Check(journalSchema, raw.value)) throw corrupt("Incomplete legacy journal");
        const row = raw.value;
        if (row.seq !== this.manifest.sequence + 1 || row.seq !== raw.key || await digest(`${row.seq}:${row.writes}`) !== row.digest) throw corrupt("Corrupt legacy journal");
        const parsed: unknown = JSON.parse(row.writes);
        if (!Value.Check(writesSchema, parsed)) throw corrupt("Invalid legacy writes");
        await this.persist(parsed, row.seq);
      }
      const next = { ...this.manifest, highWatermark: Math.max(this.manifest.highWatermark, watermark ?? 1) };
      const encoded = await encode(next);
      const tx = this.db.transaction("meta", "readwrite", { durability: "strict" });
      const done = completed(tx);
      tx.objectStore("meta").put(encoded, "manifest");
      tx.objectStore("meta").delete("migration");
      await done;
      this.manifest = next;
    }
    const tx = this.db.transaction(["records", "revisions", "requests"], "readonly");
    const done = completed(tx);
    const records = tx.objectStore("records").count();
    const revisions = tx.objectStore("revisions").count();
    const requests = tx.objectStore("requests").count();
    await done;
    if (records.result !== this.manifest.records || revisions.result !== this.manifest.revisions
      || requests.result !== this.manifest.requests) throw corrupt("Durable storage records are missing");
    await this.verifyJournal();
  }

  close(_context = BACKGROUND_CONTEXT): Promise<void> {
    if (this.shutdown) return this.shutdown;
    this.closing = true;
    this.shutdown = this.tail.then(async () => {
      window.removeEventListener("pagehide", this.pagehide);
      this.db.close();
      await this.release();
    });
    return this.shutdown;
  }
}

async function openDatabase(name: string, expectedDatabaseId?: string): Promise<IDBDatabase> {
  const empty = await encode({ sequence: 0, highWatermark: 1, records: 0, revisions: 0, checkpoint: 0, requests: 0 });
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(name, 2);
    let failure: DurableStorageError | undefined;
    request.onupgradeneeded = (event) => {
      const db = request.result;
      const tx = request.transaction;
      if (!tx) return;
      if (failure || (event.oldVersion === 0 && expectedDatabaseId !== undefined)) {
        failure ??= new DurableStorageError("missing", "Expected Durable storage is missing; it may have been evicted or cleared");
        tx.abort();
        return;
      }
      if (event.oldVersion === 0) {
        db.createObjectStore("meta");
        db.createObjectStore("journal", { keyPath: "seq" });
      } else if (!db.objectStoreNames.contains("meta") || !db.objectStoreNames.contains("journal")) {
        failure = corrupt("Legacy Durable schema is incomplete"); tx.abort(); return;
      }
      const records = db.createObjectStore("records", { keyPath: "id" });
      records.createIndex("type", "type", { unique: true });
      records.createIndex("entries", "entries", { unique: true });
      records.createIndex("address", "address", { unique: true });
      db.createObjectStore("requests");
      const revisions = db.createObjectStore("revisions", { keyPath: "key" });
      revisions.createIndex("base", "base", { unique: true });
      tx.objectStore("meta").put(crypto.randomUUID(), "databaseId");
      tx.objectStore("meta").put(empty, "manifest");
      if (event.oldVersion === 1) tx.objectStore("meta").put(true, "migration");
    };
    request.onblocked = () => {
      failure = new DurableStorageError("blocked", "Durable storage upgrade is blocked by another connection");
      reject(failure);
    };
    request.onerror = () => reject(failure ?? request.error ?? new Error("Cannot open Durable storage"));
    request.onsuccess = () => { if (failure) request.result.close(); else resolve(request.result); };
  });
}

export async function openIndexedDbStorage(name = "pi-for-excel-durable", options: {
  expectedDatabaseId?: string;
} = {}): Promise<DurableIndexedDbStorage> {
  if (name === "pi-for-excel") throw new Error("Durable storage cannot use the legacy database");
  if (typeof indexedDB === "undefined") throw new DurableStorageError("unsupported", "IndexedDB is unavailable");
  const release = await acquireStorageLock(name);
  let db: IDBDatabase | undefined;
  let storage: IndexedDbStorage | undefined;
  try {
    const persistence = await requestStoragePersistence();
    db = await openDatabase(name, options.expectedDatabaseId);
    if (STORES.some((store) => !db?.objectStoreNames.contains(store))) throw corrupt("Durable storage schema is incomplete");
    const tx = db.transaction("meta", "readonly");
    const done = completed(tx);
    const generation = tx.objectStore("meta").get("databaseId");
    const metadata = tx.objectStore("meta").get("manifest");
    await done;
    const databaseId: unknown = generation.result;
    if (typeof databaseId !== "string" || !databaseId) throw corrupt("Invalid Durable database generation");
    if (options.expectedDatabaseId !== undefined && databaseId !== options.expectedDatabaseId) throw new DurableStorageError("replaced", "Durable storage generation has changed");
    const rawMetadata: unknown = metadata.result;
    if (!Value.Check(envelopeSchema, rawMetadata)) throw corrupt("Invalid Durable manifest envelope");
    const manifest = await decode(manifestSchema, rawMetadata);
    if (manifest.checkpoint > manifest.sequence) throw corrupt("Invalid Durable checkpoint");
    storage = new IndexedDbStorage(db, release, databaseId, persistence, manifest);
    await storage.initialize();
    return storage;
  } catch (error) {
    if (storage) await storage.close();
    else { db?.close(); await release(); }
    throw error;
  }
}
