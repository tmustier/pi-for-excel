/** Fail-closed lifecycle failures; callers must not substitute an empty store. */
export class DurableStorageError extends Error {
  readonly code: "missing" | "replaced" | "corrupt" | "owned" | "unsupported" | "blocked";

  constructor(code: DurableStorageError["code"], message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "DurableStorageError";
    this.code = code;
  }
}

/** Holds a non-stealable exclusive lock until the returned idempotent release finishes. */
export async function acquireStorageLock(name: string): Promise<() => Promise<void>> {
  if (typeof navigator === "undefined" || !navigator.locks) {
    throw new DurableStorageError("unsupported", "Web Locks are unavailable");
  }
  let release = (): void => {};
  const held = new Promise<void>((resolve) => { release = resolve; });
  let acquired = (): void => {};
  const ready = new Promise<void>((resolve) => { acquired = resolve; });
  const finished = navigator.locks.request(
    `pi-for-excel-durable:${name}`,
    { mode: "exclusive", ifAvailable: true },
    async (lock) => {
      if (!lock) throw new DurableStorageError("owned", "Durable storage is already owned");
      acquired();
      await held;
    },
  );
  // Observe request failures while waiting for the grant, without detaching a rejection.
  await Promise.race([ready, finished]);
  return async () => {
    release();
    await finished;
  };
}

export async function requestStoragePersistence(): Promise<"persistent" | "best-effort" | "unsupported"> {
  if (typeof navigator === "undefined" || !navigator.storage
    || typeof navigator.storage.persisted !== "function"
    || typeof navigator.storage.persist !== "function") {
    return "unsupported";
  }
  try {
    if (await navigator.storage.persisted()) return "persistent";
    return await navigator.storage.persist() ? "persistent" : "best-effort";
  } catch {
    return "best-effort";
  }
}

function openExistingDatabase(name: string): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(name);
    let missing = false;
    request.onupgradeneeded = () => {
      // Opening a missing name would otherwise create it. Abort the creation transaction.
      missing = true;
      request.transaction?.abort();
    };
    request.onerror = () => reject(new DurableStorageError(
      missing ? "missing" : "corrupt",
      missing ? "Durable storage is missing" : "Cannot inspect durable storage",
      { cause: request.error },
    ));
    request.onsuccess = () => resolve(request.result);
    // Do not settle early: an IDB open request cannot be cancelled, and could
    // otherwise leave an unobserved live connection after a reported failure.
    request.onblocked = () => {};
  });
}

function verifyDatabaseId(database: IDBDatabase, expected: string): Promise<void> {
  return new Promise((resolve, reject) => {
    if (!database.objectStoreNames.contains("meta")) {
      reject(new DurableStorageError("corrupt", "Durable storage metadata is missing"));
      return;
    }
    const transaction = database.transaction("meta", "readonly");
    const request = transaction.objectStore("meta").get("databaseId");
    let matches = false;
    let valid = false;
    request.onsuccess = () => {
      const value: unknown = request.result;
      valid = typeof value === "string" && value.length > 0;
      matches = valid && value === expected;
    };
    transaction.oncomplete = () => {
      if (!valid) reject(new DurableStorageError("corrupt", "Durable storage generation is invalid"));
      else if (!matches) reject(new DurableStorageError("replaced", "Durable storage has been replaced"));
      else resolve();
    };
    transaction.onabort = () => reject(new DurableStorageError(
      "corrupt", "Cannot read durable storage generation", { cause: transaction.error },
    ));
    transaction.onerror = () => { /* onabort supplies the final outcome. */ };
  });
}

/**
 * Deletes only an explicitly confirmed obsolete workbook generation. No recency
 * heuristic is used. All owners must cooperate with acquireStorageLock using
 * the same database name; the generation check cannot fence noncooperating code.
 *
 * A blocked deletion stays pending, holding the lock, until IDB reports success
 * or error. deleteDatabase cannot be cancelled: rejecting on blocked/timeout
 * would allow a queued deletion to run later after a reported failure. Callers
 * must not interpret a UI timeout as cancellation or release the lock themselves.
 */
export async function deleteDurableWorkbookStorage(options: {
  name: string;
  databaseId: string;
  authorization: "user-confirmed-obsolete-workbook";
}): Promise<void> {
  const { name, databaseId, authorization } = options;
  const prefix = "pi-for-excel-durable:";
  if (authorization !== "user-confirmed-obsolete-workbook"
    || typeof name !== "string" || !name.startsWith(prefix)
    || name.slice(prefix.length).trim().length === 0
    || typeof databaseId !== "string" || databaseId.trim().length === 0) {
    throw new DurableStorageError("corrupt", "Explicit workbook name, generation, and deletion authorization are required");
  }
  if (typeof indexedDB === "undefined") {
    throw new DurableStorageError("unsupported", "IndexedDB is unavailable");
  }
  const release = await acquireStorageLock(name);
  try {
    const database = await openExistingDatabase(name);
    try {
      await verifyDatabaseId(database, databaseId);
    } finally {
      database.close();
    }
    await new Promise<void>((resolve, reject) => {
      const request = indexedDB.deleteDatabase(name);
      let blocked = false;
      request.onblocked = () => { blocked = true; };
      request.onsuccess = () => resolve();
      request.onerror = () => reject(new DurableStorageError(
        blocked ? "blocked" : "corrupt", "Cannot delete durable storage", { cause: request.error },
      ));
    });
  } finally {
    await release();
  }
}
