import { isJsonValue, type Context, type JsonValue } from "@earendil-works/chord";
import { BACKGROUND_CONTEXT } from "@earendil-works/chord/context";
import { assertValidOp, type Op } from "@earendil-works/chord/delta";
import type { Id, Seq, Storage, StorageWrite } from "@earendil-works/pi-durable";
import { MemoryStorage } from "@earendil-works/pi-durable/storage/memory";
import { Type } from "typebox";
import { Value } from "typebox/value";

const DEFAULT_DATABASE = "pi-for-excel-durable";
const JOURNAL = "journal";
const META = "meta";
const counterSchema = Type.Integer({ minimum: 1, maximum: Number.MAX_SAFE_INTEGER });
const jsonSchema = Type.Refine(Type.Unsafe<JsonValue>({}), isJsonValue);
const objectSchema = Type.Record(Type.String(), jsonSchema);
const opSchema = Type.Refine(Type.Unsafe<Op>({}), (op) => {
  try { assertValidOp(op); return true; } catch { return false; }
});
const optionalId = Type.Optional(counterSchema);
const absent = Type.Optional(Type.Never());
const textSchema = Type.Object({ type: Type.Literal("text"), text: Type.String(), textSignature: Type.Optional(Type.String()) });
const imageSchema = Type.Object({ type: Type.Literal("image"), data: Type.String(), mimeType: Type.String() });
const usageSchema = Type.Object({
  input: Type.Number(), output: Type.Number(), cacheRead: Type.Number(), cacheWrite: Type.Number(),
  cacheWrite1h: Type.Optional(Type.Number()), reasoning: Type.Optional(Type.Number()), totalTokens: Type.Number(),
  cost: Type.Object({ input: Type.Number(), output: Type.Number(), cacheRead: Type.Number(), cacheWrite: Type.Number(), total: Type.Number() }),
});
const errorSchema = Type.Object({ message: Type.String(), detail: Type.Optional(jsonSchema) });
const messageSchema = Type.Union([
  Type.Object({ role: Type.Literal("user"), timestamp: Type.Number(), content: Type.Union([Type.String(), Type.Array(Type.Union([textSchema, imageSchema]))]) }),
  Type.Object({
    role: Type.Literal("system"), timestamp: Type.Number(), content: Type.Union([Type.String(), Type.Array(textSchema)]),
    sections: Type.Optional(Type.Record(Type.String(), Type.Union([Type.String(), Type.Null()]))),
    toolsAdded: Type.Optional(Type.Array(Type.Object({
      name: Type.String(), description: Type.String(), parameters: objectSchema,
      constrainedSampling: Type.Optional(Type.Union([
        Type.Literal(false), Type.Object({ type: Type.Literal("json_schema"), strict: Type.Union([Type.Literal("prefer"), Type.Literal("require")]) }),
        Type.Object({ type: Type.Literal("grammar"), variants: Type.Object({ openai_lark: Type.Optional(Type.String()), openai_regex: Type.Optional(Type.String()) }) }),
      ])),
    }))),
    toolsRemoved: Type.Optional(Type.Array(Type.Object({ name: Type.String() }))),
  }),
  Type.Object({
    role: Type.Literal("assistant"), timestamp: Type.Number(), api: Type.String(), provider: Type.String(), model: Type.String(),
    content: Type.Array(Type.Union([textSchema,
      Type.Object({ type: Type.Literal("thinking"), thinking: Type.String(), thinkingSignature: Type.Optional(Type.String()), redacted: Type.Optional(Type.Boolean()) }),
      Type.Object({ type: Type.Literal("toolCall"), id: Type.String(), name: Type.String(), arguments: objectSchema,
        thoughtSignature: Type.Optional(Type.String()), namespace: Type.Optional(Type.String()), argumentsError: Type.Optional(Type.String()) }),
    ])),
    responseModel: Type.Optional(Type.String()), responseId: Type.Optional(Type.String()), providerThinkingLevel: Type.Optional(Type.String()),
    thinkingLevel: Type.Optional(Type.Union([Type.Literal("off"), Type.Literal("minimal"), Type.Literal("low"), Type.Literal("medium"), Type.Literal("high"), Type.Literal("xhigh"), Type.Literal("max")])),
    usage: usageSchema,
    stopReason: Type.Union([Type.Literal("pending"), Type.Literal("stop"), Type.Literal("length"), Type.Literal("toolUse"), Type.Literal("error"), Type.Literal("aborted"), Type.Literal("deferred")]),
    deferred: Type.Optional(Type.Object({ provider: Type.String(), modelId: Type.String(), api: Type.String(), id: Type.String(),
      expiresAt: Type.Optional(Type.Number()), pollAfterMs: Type.Optional(Type.Number()), data: Type.Optional(jsonSchema) })),
    diagnostics: Type.Optional(Type.Array(Type.Object({ type: Type.String(), timestamp: Type.Number(), details: Type.Optional(objectSchema),
      error: Type.Optional(Type.Object({ message: Type.String(), name: Type.Optional(Type.String()), stack: Type.Optional(Type.String()), code: Type.Optional(Type.Union([Type.String(), Type.Number()])) })) }))),
    errorMessage: Type.Optional(Type.String()), rawStopReason: Type.Optional(Type.String()), endTurn: Type.Optional(Type.Boolean()), durationMs: Type.Optional(Type.Number()),
  }),
  Type.Object({
    role: Type.Literal("toolResult"), timestamp: Type.Number(), toolCallId: Type.String(), toolName: Type.String(),
    content: Type.Array(Type.Union([textSchema, imageSchema])), isError: Type.Boolean(), details: Type.Optional(jsonSchema),
    usage: Type.Optional(usageSchema), durationMs: Type.Optional(Type.Number()),
    nestedCalls: Type.Optional(Type.Object({ complete: Type.Boolean(), calls: Type.Array(Type.Object({
      id: Type.String(), name: Type.String(), arguments: Type.Optional(objectSchema), argumentsBytes: Type.Optional(Type.Number()),
      status: Type.Union([Type.Literal("ok"), Type.Literal("error"), Type.Literal("unfinished")]), durationMs: Type.Optional(Type.Number()), error: Type.Optional(Type.String()),
    })) })),
  }),
]);
const conversationSchema = Type.Object({ id: counterSchema,
  parent: Type.Optional(Type.Object({ conversationId: counterSchema, at: counterSchema })),
  owner: Type.Optional(Type.Object({ conversationId: counterSchema, taskId: counterSchema })),
});
const entrySchema = Type.Object({
  id: counterSchema, conversationId: counterSchema, kind: Type.String(), model: Type.Optional(Type.Array(messageSchema)),
  data: Type.Optional(jsonSchema), head: optionalId, byTaskId: optionalId,
  edits: Type.Optional(Type.Array(Type.Union([
    Type.Object({ target: counterSchema, action: Type.Literal("omit"), messages: absent }),
    Type.Object({ target: counterSchema, action: Type.Literal("replace"), messages: Type.Array(messageSchema) }),
  ]))),
});
const outcomeSchema = Type.Union([
  Type.Object({ status: Type.Literal("completed"), result: jsonSchema, error: absent, reason: absent }),
  Type.Object({ status: Type.Literal("failed"), error: errorSchema, result: Type.Optional(jsonSchema), reason: absent }),
  Type.Object({ status: Type.Literal("aborted"), reason: Type.Optional(Type.String()), result: Type.Optional(jsonSchema), error: absent }),
  Type.Object({ status: Type.Literal("orphaned"), reason: Type.String(), error: absent, result: absent }),
  Type.Object({ status: Type.Literal("faulted"), error: errorSchema, result: absent, reason: absent }),
]);
const taskIdentity = {
  id: counterSchema, conversationId: counterSchema, kind: Type.String(), version: Type.Number(), input: jsonSchema,
  owner: optionalId, background: Type.Boolean(), abortRequested: Type.Boolean(), abortReason: Type.Optional(Type.Literal("restart")),
  abandonOnRestart: Type.Optional(Type.Literal(true)), startedAt: Type.Optional(Type.Number()), endedAt: Type.Optional(Type.Number()),
};
const taskSchema = Type.Union([
  Type.Object({ ...taskIdentity, memos: Type.Optional(objectSchema), state: Type.Union([
    Type.Object({ status: Type.Union([Type.Literal("pending"), Type.Literal("running")]), checkpoint: jsonSchema, outcome: absent }),
    Type.Object({ status: Type.Literal("waiting"), checkpoint: jsonSchema, on: Type.Array(counterSchema), policy: Type.Union([Type.Literal("failFast"), Type.Literal("allSettled")]), outcome: absent }),
  ]) }),
  Type.Object({ ...taskIdentity, memos: absent, state: Type.Object({ status: Type.Union([Type.Literal("completing"), Type.Literal("terminal")]), outcome: outcomeSchema, checkpoint: absent }) }),
]);
const submissionIdentity = { id: counterSchema, conversationId: counterSchema, requestId: Type.Optional(Type.String()) };
const submissionSchema = Type.Union([
  Type.Object({ ...submissionIdentity, type: Type.Union([Type.Literal("input"), Type.Literal("write")]), status: Type.Literal("queued"), entry: absent, answer: absent, reason: absent, detail: absent }),
  Type.Object({ ...submissionIdentity, type: Type.Literal("input"), status: Type.Literal("placed"), entry: counterSchema, answer: absent, reason: absent, detail: absent }),
  Type.Object({ ...submissionIdentity, type: Type.Literal("input"), status: Type.Literal("done"), entry: counterSchema, answer: counterSchema, reason: absent, detail: absent }),
  Type.Object({ ...submissionIdentity, type: Type.Literal("write"), status: Type.Literal("done"), entry: counterSchema, answer: absent, reason: absent, detail: absent }),
  Type.Object({ ...submissionIdentity, type: Type.Literal("input"), status: Type.Literal("unanswered"), entry: optionalId, answer: absent, reason: Type.String(), detail: Type.Optional(jsonSchema) }),
  Type.Object({ ...submissionIdentity, type: Type.Literal("write"), status: Type.Literal("unanswered"), entry: absent, answer: absent, reason: Type.String(), detail: Type.Optional(jsonSchema) }),
]);
const documentIdentity = { id: counterSchema, kind: Type.String(), key: Type.Optional(Type.String()) };
const documentSchema = Type.Union([
  Type.Object({ ...documentIdentity, scope: Type.Object({ kind: Type.Literal("session") }), history: absent, fork: absent }),
  Type.Object({ ...documentIdentity, scope: Type.Object({ kind: Type.Literal("task"), taskId: counterSchema }), history: absent, fork: absent }),
  Type.Object({ ...documentIdentity, scope: Type.Object({ kind: Type.Literal("conversation"), conversationId: counterSchema }), history: Type.Literal("latest"), fork: Type.Union([Type.Literal("current"), Type.Literal("initial")]) }),
  Type.Object({ ...documentIdentity, scope: Type.Object({ kind: Type.Literal("conversation"), conversationId: counterSchema }), history: Type.Literal("rewindable"), fork: Type.Union([Type.Literal("asOf"), Type.Literal("current"), Type.Literal("initial")]) }),
]);
const baseSchema = Type.Object({ kind: Type.Literal("base"), version: Type.Number(), value: objectSchema });
const contentSchema = Type.Union([baseSchema, Type.Object({ kind: Type.Literal("delta"), version: Type.Number(), ops: Type.Array(opSchema) })]);
// The schema checks the 1.3.0 wire contract. Unsafe supplies only erased ID brands
// and readonly TypeScript annotations; it does not bypass runtime validation.
const writesSchema = Type.Unsafe<readonly StorageWrite[]>(Type.Array(Type.Union([
  Type.Object({ type: Type.Literal("conversation"), value: conversationSchema }),
  Type.Object({ type: Type.Literal("entry"), value: entrySchema }),
  Type.Object({ type: Type.Literal("task"), value: taskSchema }),
  Type.Object({ type: Type.Literal("submission"), value: submissionSchema }),
  Type.Object({ type: Type.Literal("document.create"), record: documentSchema, content: baseSchema }),
  Type.Object({ type: Type.Literal("document.change"), id: counterSchema, content: contentSchema }),
  Type.Object({ type: Type.Literal("document.retire"), id: counterSchema }),
])));
const journalSchema = Type.Object({
  seq: counterSchema,
  writes: Type.String(),
  digest: Type.String(),
}, { additionalProperties: false });

function completed(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onabort = () => reject(transaction.error ?? new Error("Durable storage transaction aborted"));
    transaction.onerror = () => { /* onabort is the final transaction outcome. */ };
  });
}

async function digest(text: string): Promise<string> {
  const bytes = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(bytes), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function openDatabase(name: string): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(name, 1);
    let rejected = false;
    request.onupgradeneeded = () => {
      request.result.createObjectStore(JOURNAL, { keyPath: "seq" });
      request.result.createObjectStore(META);
    };
    request.onblocked = () => {
      rejected = true;
      reject(new Error("Durable storage database upgrade is blocked"));
    };
    request.onerror = () => reject(request.error ?? new Error("Cannot open durable storage"));
    request.onsuccess = () => {
      if (rejected) request.result.close();
      else resolve(request.result);
    };
  });
}

/** Internal subclass: callers receive Storage, never the in-memory preparation API. */
class IndexedDbStorage extends MemoryStorage {
  private tail: Promise<void> = Promise.resolve();
  private closing = false;
  private shutdown: Promise<void> | undefined;
  private highWatermark = 1;
  private readonly db: IDBDatabase;
  private readonly release: () => Promise<void>;

  constructor(db: IDBDatabase, release: () => Promise<void>) {
    super();
    this.db = db;
    this.release = release;
    db.onversionchange = () => { void this.close(); };
    window.addEventListener("pagehide", this.pagehide);
  }

  private readonly pagehide = (event: PageTransitionEvent): void => {
    // A restored page must not keep executing with a lock it released on hide.
    if (event.persisted) window.addEventListener("pageshow", () => window.location.reload(), { once: true });
    void this.close();
  };

  private serialize<T>(operation: () => Promise<T>): Promise<T> {
    if (this.closing) return Promise.reject(new Error("Durable storage is closed"));
    const result = this.tail.then(operation);
    this.tail = result.then(() => {}, () => {});
    return result;
  }

  async restore(): Promise<void> {
    const transaction = this.db.transaction([JOURNAL, META], "readonly");
    const done = completed(transaction);
    const rows = transaction.objectStore(JOURNAL).getAll();
    const watermark = transaction.objectStore(META).get("highWatermark");
    const sequence = transaction.objectStore(META).get("sequence");
    await done;
    const rawSeq: unknown = sequence.result;
    if (rawSeq !== undefined && !Value.Check(counterSchema, rawSeq)) throw new Error("Invalid durable sequence");
    const rawWatermark: unknown = watermark.result;
    if (rawWatermark !== undefined) {
      if (!Value.Check(counterSchema, rawWatermark)) throw new Error("Invalid durable ID watermark");
      this.highWatermark = rawWatermark;
    }
    const rawRows: unknown = rows.result;
    if (!Value.Check(Type.Array(journalSchema), rawRows)) throw new Error("Invalid durable journal");
    let lastSeq = 0;
    for (const row of rawRows) {
      if (await digest(`${row.seq}:${row.writes}`) !== row.digest) throw new Error("Corrupt durable journal");
      const rawWrites: unknown = JSON.parse(row.writes);
      if (!Value.Check(writesSchema, rawWrites)) throw new Error("Invalid durable writes");
      const writes = rawWrites;
      // The schema proves the erased sequence brand's safe-integer invariant.
      super.prepareCommit(writes, row.seq as Seq).apply();
      this.observeIds(writes);
      lastSeq = row.seq;
    }
    if (lastSeq !== (rawSeq ?? 0)) throw new Error("Incomplete durable journal");
  }

  private observeIds(writes: readonly StorageWrite[]): void {
    for (const write of writes) {
      const id = write.type === "document.create" || write.type === "document.copy"
        ? write.record.id
        : write.type === "document.change" || write.type === "document.retire"
          ? write.id : write.value.id;
      this.highWatermark = Math.max(this.highWatermark, id);
    }
  }

  override async commit(writes: readonly StorageWrite[], _context: Context): Promise<Seq> {
    // Detach at invocation, not after another pending mutation has finished.
    const detached = structuredClone(writes);
    return this.serialize(async () => {
      const prepared = super.prepareCommit(detached);
      if (!Value.Check(writesSchema, prepared.writes)) throw new Error("Invalid durable writes");
      const encoded = JSON.stringify(prepared.writes);
      const hash = await digest(`${prepared.seq}:${encoded}`);
      const transaction = this.db.transaction([JOURNAL, META], "readwrite");
      const done = completed(transaction);
      try {
        transaction.objectStore(JOURNAL).add({ seq: prepared.seq, writes: encoded, digest: hash });
        transaction.objectStore(META).put(prepared.seq, "sequence");
      } catch (error) {
        transaction.abort();
        await done.catch(() => {});
        throw error;
      }
      await done;
      const seq = prepared.apply();
      this.observeIds(prepared.writes);
      return seq;
    });
  }

  override mintId<I extends Id<string>>(): Promise<I> {
    return this.serialize(async () => {
      const id = this.highWatermark + 1;
      if (!Number.isSafeInteger(id)) throw new Error("Durable ID space is exhausted");
      const transaction = this.db.transaction(META, "readwrite");
      const done = completed(transaction);
      transaction.objectStore(META).put(id, "highWatermark");
      await done;
      this.highWatermark = id;
      // Erased brand applied only after validated allocation and durable reservation.
      return id as I;
    });
  }

  override close(context = BACKGROUND_CONTEXT): Promise<void> {
    if (this.shutdown) return this.shutdown;
    this.closing = true;
    this.shutdown = this.tail.then(async () => {
      await super.close(context);
      window.removeEventListener("pagehide", this.pagehide);
      this.db.close();
      await this.release();
    });
    return this.shutdown;
  }
}

/** Open a dedicated Durable 1.3 database with one exclusive pane owner. */
export async function openIndexedDbStorage(name = DEFAULT_DATABASE): Promise<Storage> {
  if (!navigator.locks?.request) throw new Error("Durable storage requires Web Locks");
  if (name === "pi-for-excel") throw new Error("Durable storage cannot use the legacy database");
  let unlock: () => void = () => {};
  const held = new Promise<void>((resolve) => { unlock = resolve; });
  let acquired: (value: boolean) => void = () => {};
  const acquisition = new Promise<boolean>((resolve) => { acquired = resolve; });
  const lifetime = navigator.locks.request(`pi-for-excel-durable:${name}`, { mode: "exclusive", ifAvailable: true }, async (lock) => {
    acquired(lock !== null);
    if (lock) await held;
  });
  // Propagate API failures as well as an unavailable lock, without leaving a waiter.
  const available = await Promise.race([acquisition, lifetime.then(() => false)]);
  if (!available) throw new Error("Durable storage is already owned by another pane");
  const release = async (): Promise<void> => { unlock(); await lifetime; };
  let db: IDBDatabase | undefined;
  let storage: IndexedDbStorage | undefined;
  try {
    db = await openDatabase(name);
    storage = new IndexedDbStorage(db, release);
    await storage.restore();
    return storage;
  } catch (error) {
    if (storage) await storage.close();
    else { db?.close(); await release(); }
    throw error;
  }
}
