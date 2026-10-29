import { Type, type Static } from "typebox";
import { Value } from "typebox/value";
import { IndexedDBStorageBackend } from "../../storage/local/indexeddb-storage-backend.js";
import type { StoreConfig } from "../../storage/local/types.js";
import type { WorkbookRecoverySnapshot } from "../recovery-log.js";
import { MAX_RECOVERY_BYTES, RECOVERY_PAGE_SIZE } from "./constants.js";
import {
  createPersistedWorkbookRecoveryPayload,
  parsePersistedSnapshots,
  PersistedWorkbookRecoveryPayloadSchema,
  PersistedWorkbookRecoverySnapshotSchema,
  WorkbookRecoveryToolNameSchema,
} from "./log-codec.js";

export const RECOVERY_SETTING_KEY = "workbook.recovery-snapshots.v1";
const META = "recovery-metadata";
const PAYLOAD = "recovery-payloads";
const STATE = "recovery-state";
const STORES = [META, PAYLOAD, STATE];

export const RecoveryCursorSchema = Type.Object({ at: Type.Number(), id: Type.String() });
export type RecoveryCursor = Static<typeof RecoveryCursorSchema>;
export const RecoveryMetadataSchema = Type.Object({
  id: Type.String(), at: Type.Number(), workbookId: Type.String(),
  workbookLabel: Type.Optional(Type.String()), toolName: WorkbookRecoveryToolNameSchema,
  toolCallId: Type.String(), address: Type.String(), changedCount: Type.Number(), cellCount: Type.Number(),
  snapshotKind: Type.Optional(Type.String()), restoredFromSnapshotId: Type.Optional(Type.String()),
  turnId: Type.Optional(Type.String()), sizeBytes: Type.Integer({ minimum: 0 }),
});
export type RecoveryMetadata = Static<typeof RecoveryMetadataSchema>;
export interface RecoveryPage {
  entries: RecoveryMetadata[];
  nextCursor?: RecoveryCursor;
}
export interface RecoveryPageOptions {
  cursor?: RecoveryCursor;
  limit?: number;
  turnId?: string;
  oldestFirst?: boolean;
}
const AccountingSchema = Type.Object({ bytes: Type.Integer({ minimum: 0 }) });
type Accounting = Static<typeof AccountingSchema>;

export function recoveryStoreConfigs(): StoreConfig[] {
  return [
    { name: META, keyPath: "id", indices: [
      { name: "time", keyPath: ["at", "id"] },
      { name: "workbook", keyPath: ["workbookId", "at", "id"] },
      { name: "turn", keyPath: ["workbookId", "turnId", "at", "id"] },
    ] },
    { name: PAYLOAD }, { name: STATE },
  ];
}

function request<T>(req: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error("Backup storage request failed."));
  });
}

function metadata(cursor: IDBCursorWithValue): RecoveryMetadata {
  const raw: unknown = cursor.value;
  if (!Value.Check(RecoveryMetadataSchema, raw)) throw new Error("Invalid backup metadata; nothing was deleted.");
  return raw;
}

async function accounting(tx: IDBTransaction): Promise<Accounting> {
  const raw: unknown = await request(tx.objectStore(STATE).get("accounting"));
  if (!Value.Check(AccountingSchema, raw)) throw new Error("Backup accounting is unavailable; nothing was deleted.");
  return raw;
}

function withSize(entry: RecoveryMetadata, payloadBytes: number): RecoveryMetadata {
  const next = { ...entry, sizeBytes: 0 };
  for (;;) {
    const sizeBytes = payloadBytes + new Blob([JSON.stringify(next)]).size;
    if (sizeBytes === next.sizeBytes) return next;
    next.sizeBytes = sizeBytes;
  }
}

function snapshotMetadata(snapshot: WorkbookRecoverySnapshot): RecoveryMetadata {
  const { id, at, workbookId, workbookLabel, toolName, toolCallId, address, changedCount, cellCount,
    snapshotKind, restoredFromSnapshotId, turnId } = snapshot;
  return { id, at, workbookId: workbookId ?? "", toolName, toolCallId, address,
    changedCount, cellCount, sizeBytes: 0,
    ...(workbookLabel === undefined ? {} : { workbookLabel }),
    ...(snapshotKind === undefined ? {} : { snapshotKind }),
    ...(restoredFromSnapshotId === undefined ? {} : { restoredFromSnapshotId }),
    ...(turnId === undefined ? {} : { turnId }),
  };
}

async function encode(snapshot: WorkbookRecoverySnapshot): Promise<{ entry: RecoveryMetadata; blob: Blob }> {
  const raw = new Blob([JSON.stringify(createPersistedWorkbookRecoveryPayload([snapshot]))]);
  const blob = await new Response(raw.stream().pipeThrough(new CompressionStream("gzip"))).blob();
  return { entry: withSize(snapshotMetadata(snapshot), blob.size), blob };
}

function walk(
  index: IDBIndex, range: IDBKeyRange | null, direction: IDBCursorDirection,
  visit: (entry: RecoveryMetadata, cursor: IDBCursorWithValue) => boolean,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const req = index.openCursor(range, direction);
    req.onerror = () => reject(req.error ?? new Error("Backup index read failed."));
    req.onsuccess = () => {
      try {
        const cursor = req.result;
        if (!cursor || !visit(metadata(cursor), cursor)) resolve();
        else cursor.continue();
      } catch (error) { reject(error instanceof Error ? error : new Error("Backup index read failed.")); }
    };
  });
}

export class RecoveryStore {
  private readonly backend: IndexedDBStorageBackend;
  private readonly budget: number;

  constructor(backend: IndexedDBStorageBackend, budget = MAX_RECOVERY_BYTES) {
    this.backend = backend;
    this.budget = budget;
  }

  async migrate(): Promise<void> {
    if (Value.Check(Type.Literal(true), await this.backend.get(STATE, "migrated"))) return;
    const raw: unknown = await this.backend.get("settings", RECOVERY_SETTING_KEY);
    if (raw !== null && (!Value.Check(PersistedWorkbookRecoveryPayloadSchema, raw)
      || raw.snapshots.some((item) => !Value.Check(PersistedWorkbookRecoverySnapshotSchema, item)))) {
      throw new Error("Legacy backups could not be migrated; original data is unchanged.");
    }
    const source = JSON.stringify(raw);
    const snapshots = parsePersistedSnapshots(raw, { maxEntries: Infinity });
    const encoded: Awaited<ReturnType<typeof encode>>[] = [];
    const ids = new Set<string>();
    let bytes = 0;
    for (const snapshot of snapshots) {
      if (ids.has(snapshot.id)) throw new Error("Duplicate legacy backup ids; original data is unchanged.");
      ids.add(snapshot.id);
      const record = await encode(snapshot);
      bytes += record.entry.sizeBytes;
      if (bytes > this.budget) throw new Error("Legacy backups exceed the storage ceiling; original data is unchanged.");
      encoded.push(record);
    }
    await this.backend.indexedTransaction([...STORES, "settings"], "readwrite", async (tx) => {
      if (Value.Check(Type.Literal(true), await request(tx.objectStore(STATE).get("migrated")))) return;
      const current: unknown = await request(tx.objectStore("settings").get(RECOVERY_SETTING_KEY));
      if (JSON.stringify(current ?? null) !== source) throw new Error("Legacy backups changed; retry migration.");
      for (const { entry, blob } of encoded) {
        tx.objectStore(META).add(entry);
        tx.objectStore(PAYLOAD).add(blob, entry.id);
      }
      tx.objectStore(STATE).put({ bytes }, "accounting");
      tx.objectStore(STATE).put(true, "migrated");
      tx.objectStore("settings").delete(RECOVERY_SETTING_KEY);
      tx.objectStore("settings").delete("workbook.recovery.retentionLimit.v1");
    });
  }

  private async legacy(): Promise<WorkbookRecoverySnapshot[] | null> {
    try { await this.migrate(); return null; }
    catch (error) {
      console.warn("[pi] Backup migration deferred; original data retained.", error);
      if (Value.Check(Type.Literal(true), await this.backend.get(STATE, "migrated"))) return null;
      return parsePersistedSnapshots(await this.backend.get("settings", RECOVERY_SETTING_KEY), { maxEntries: Infinity });
    }
  }

  async append(snapshot: WorkbookRecoverySnapshot): Promise<void> {
    await this.migrate();
    const { entry, blob } = await encode(snapshot);
    if (entry.sizeBytes > this.budget) throw new Error("Too large to back up: compressed backup exceeds the storage ceiling.");
    await this.backend.indexedTransaction(STORES, "readwrite", async (tx) => {
      const state = await accounting(tx);
      tx.objectStore(META).add(entry);
      tx.objectStore(PAYLOAD).add(blob, entry.id);
      let bytes = state.bytes + entry.sizeBytes;
      if (bytes > this.budget) {
        await walk(tx.objectStore(META).index("time"), null, "next", (old, cursor) => {
          if (bytes <= this.budget) return false;
          bytes -= old.sizeBytes;
          cursor.delete();
          tx.objectStore(PAYLOAD).delete(old.id);
          return true;
        });
      }
      tx.objectStore(STATE).put({ bytes }, "accounting");
    });
  }

  async page(workbookId?: string, options: RecoveryPageOptions = {}): Promise<RecoveryPage> {
    const limit = Math.min(RECOVERY_PAGE_SIZE, Math.max(1, Math.floor(options.limit ?? RECOVERY_PAGE_SIZE)));
    const legacy = await this.legacy();
    if (legacy) {
      const entries = legacy.filter((entry) => workbookId === undefined || entry.workbookId === workbookId)
        .filter((entry) => options.turnId === undefined || entry.turnId === options.turnId)
        .sort((a, b) => (a.at - b.at || a.id.localeCompare(b.id)) * (options.oldestFirst ? 1 : -1))
        .filter((entry) => !options.cursor || (options.oldestFirst
          ? entry.at > options.cursor.at || (entry.at === options.cursor.at && entry.id > options.cursor.id)
          : entry.at < options.cursor.at || (entry.at === options.cursor.at && entry.id < options.cursor.id)))
        .slice(0, limit + 1)
        .map(snapshotMetadata);
      return this.finishPage(entries, limit);
    }
    return this.backend.indexedTransaction([META], "readonly", async (tx) => {
      const prefix = workbookId === undefined ? [] : options.turnId === undefined ? [workbookId] : [workbookId, options.turnId];
      const lower = [...prefix, -Infinity, ""];
      const upper = [...prefix, Infinity, []];
      const cursor = options.cursor ? [...prefix, options.cursor.at, options.cursor.id] : null;
      const range = options.oldestFirst
        ? IDBKeyRange.bound(cursor ?? lower, upper, cursor !== null, false)
        : IDBKeyRange.bound(lower, cursor ?? upper, false, cursor !== null);
      const entries: RecoveryMetadata[] = [];
      await walk(tx.objectStore(META).index(workbookId === undefined ? "time" : options.turnId === undefined ? "workbook" : "turn"),
        range, options.oldestFirst ? "next" : "prev", (entry) => {
          entries.push(entry);
          return entries.length <= limit;
        });
      return this.finishPage(entries, limit);
    });
  }

  private finishPage(entries: RecoveryMetadata[], limit: number): RecoveryPage {
    const more = entries.length > limit;
    const page = entries.slice(0, limit);
    const last = page.at(-1);
    return { entries: page, ...(more && last ? { nextCursor: { at: last.at, id: last.id } } : {}) };
  }

  async read(id: string, workbookId?: string): Promise<WorkbookRecoverySnapshot | null> {
    const legacy = await this.legacy();
    if (legacy) {
      const snapshot = legacy.find((entry) => entry.id === id) ?? null;
      if (snapshot && workbookId !== undefined) this.checkIdentity(snapshot.workbookId, workbookId);
      return snapshot;
    }
    const record = await this.backend.indexedTransaction([META, PAYLOAD], "readonly", async (tx) => {
      const raw: unknown = await request(tx.objectStore(META).get(id));
      if (raw === undefined) return null;
      if (!Value.Check(RecoveryMetadataSchema, raw)) throw new Error("Invalid backup metadata.");
      if (workbookId !== undefined) this.checkIdentity(raw.workbookId, workbookId);
      const blob: unknown = await request(tx.objectStore(PAYLOAD).get(id));
      if (!(blob instanceof Blob)) throw new Error("Backup payload is unavailable.");
      return { entry: raw, blob };
    });
    if (!record) return null;
    const text = await new Response(record.blob.stream().pipeThrough(new DecompressionStream("gzip"))).text();
    const raw: unknown = JSON.parse(text);
    const snapshot = parsePersistedSnapshots(raw, { maxEntries: 1 })[0];
    if (!snapshot || snapshot.id !== id) throw new Error("Invalid backup payload.");
    return { ...snapshot, workbookId: record.entry.workbookId,
      ...(record.entry.workbookLabel === undefined ? {} : { workbookLabel: record.entry.workbookLabel }) };
  }

  private checkIdentity(stored: string | undefined, expected: string): void {
    if (!stored) throw new Error("Snapshot is missing workbook identity; cannot safely restore it.");
    if (stored !== expected) throw new Error("Snapshot belongs to a different workbook.");
  }

  async delete(id: string, workbookId: string): Promise<boolean> {
    await this.migrate();
    return this.backend.indexedTransaction(STORES, "readwrite", async (tx) => {
      const raw: unknown = await request(tx.objectStore(META).get(id));
      if (raw === undefined) return false;
      if (!Value.Check(RecoveryMetadataSchema, raw)) throw new Error("Invalid backup metadata; nothing was deleted.");
      if (raw.workbookId !== workbookId) return false;
      const state = await accounting(tx);
      tx.objectStore(META).delete(id);
      tx.objectStore(PAYLOAD).delete(id);
      tx.objectStore(STATE).put({ bytes: state.bytes - raw.sizeBytes }, "accounting");
      return true;
    });
  }

  async clear(workbookId: string): Promise<number> {
    await this.migrate();
    return this.backend.indexedTransaction(STORES, "readwrite", async (tx) => {
      const state = await accounting(tx);
      let removed = 0;
      let bytes = state.bytes;
      await walk(tx.objectStore(META).index("workbook"), IDBKeyRange.bound([workbookId, -Infinity], [workbookId, Infinity]), "next", (entry, cursor) => {
        bytes -= entry.sizeBytes;
        removed += 1;
        cursor.delete();
        tx.objectStore(PAYLOAD).delete(entry.id);
        return true;
      });
      tx.objectStore(STATE).put({ bytes }, "accounting");
      return removed;
    });
  }

  async move(from: string, to: string, label: string): Promise<number> {
    if (from === to) return 0;
    await this.migrate();
    return this.backend.indexedTransaction([META, PAYLOAD, STATE], "readwrite", async (tx) => {
      const state = await accounting(tx);
      let bytes = state.bytes;
      let moved = 0;
      await walk(tx.objectStore(META).index("workbook"), IDBKeyRange.bound([from, -Infinity], [from, Infinity]), "next", (entry, cursor) => {
        const payloadBytes = entry.sizeBytes - new Blob([JSON.stringify(entry)]).size;
        const next = withSize({ ...entry, workbookId: to, workbookLabel: label }, payloadBytes);
        bytes += next.sizeBytes - entry.sizeBytes;
        cursor.update(next);
        moved += 1;
        return true;
      });
      if (bytes > this.budget) {
        await walk(tx.objectStore(META).index("time"), null, "next", (entry, cursor) => {
          if (bytes <= this.budget) return false;
          bytes -= entry.sizeBytes;
          cursor.delete();
          tx.objectStore(PAYLOAD).delete(entry.id);
          return true;
        });
      }
      tx.objectStore(STATE).put({ bytes }, "accounting");
      return moved;
    });
  }
}

export async function defaultGetRecoveryStore(): Promise<RecoveryStore | null> {
  const { getAppStorage } = await import("../../storage/local/app-storage.js");
  const backend = getAppStorage().backend;
  return backend instanceof IndexedDBStorageBackend ? new RecoveryStore(backend) : null;
}
