# Durable IndexedDB storage

This backend implements the published Pi Durable **1.3.0** `Storage` contract.
It does not retain a `MemoryStorage` mirror. IndexedDB holds the complete history;
queries load pages or one document, rather than replaying the whole session into
JavaScript memory. The default taskpane integration is owned separately.

## Host integration

```ts
import {
  openIndexedDbStorage,
  DurableStorageError,
  deleteDurableWorkbookStorage,
} from "./durable/indexeddb-storage.js";

const storage = await openIndexedDbStorage(databaseName, {
  expectedDatabaseId, // omit only on first creation/adoption
});
// Before admitting new work, persist storage.databaseId in the host's workbook
// association. On subsequent opens supply that independently retained identity.
// storage satisfies Storage and can be passed directly to the Harness.
// Show storage.persistence: "persistent" | "best-effort" | "unsupported".
```

With `exactOptionalPropertyTypes`, omit the options field when there is no
expected ID; do not pass `{ expectedDatabaseId: undefined }`.

Use the existing workbook identity, never the raw workbook URL, in
`pi-for-excel-durable:<workbookId>`. Generation IDs are random opaque strings.
The host must preserve the association until the user explicitly chooses to
discard or replace it. Do not catch an open/read failure and create an empty
fallback store, remove the expectation, re-import over it, or automatically
start a replacement session.

**An independent expectation is necessary for eviction detection.** Opening
with an expectation aborts creation if the database is absent and rejects a
different generation. Without it, an absent database is indistinguishable from
first use. A marker in the same browser origin can itself disappear during
whole-origin eviction/clearing; a workbook or external backup is required to
detect that case. No browser backend can recover evicted bytes without a backup.
The API deliberately calls this `missing`, not a proven browser eviction:
manual clearing, a different storage partition, and eviction look alike.

`DurableStorageError.code` distinguishes `missing`, `replaced`, `corrupt`,
`owned`, `unsupported`, and `blocked`. Ordinary IndexedDB errors (including
quota/I/O failure and transaction abort) propagate without being relabelled as
missing data. Read-side `StorageRequestError` retains upstream semantics for
invalid cursors, unknown conversations, and unsupported historical content.

Persistence is requested on open using `navigator.storage.persisted()` and
`persist()`. Denial or API failure yields `best-effort`; missing APIs yield
`unsupported`. Neither prevents use of IndexedDB, but the host should disclose
the risk. A grant is not a backup or a promise against manual deletion, storage
partition changes, hardware failure, or an Office-specific lifecycle event.

`await storage.checkpoint()` returns `{ sequence, pruned }`. It verifies and
removes redundant journal batches, atomically advancing the checkpoint marker.
It is optional maintenance: successful commits already persist the materialized
records, document revisions, request lookup, allocator and sequence atomically.
The backend also prunes automatically every 64 commits
(agent default, not requested). Never treat a maintenance failure as permission
to resubmit an already committed mutation.

## Durable representation and bounds

Schema version 2 has five stores:

| Store | Contents |
| --- | --- |
| `meta` | `databaseId`; checksummed `manifest` with sequence, ID watermark, checkpoint sequence, and record/revision/request counts; migration marker during v1 conversion |
| `records` | Global numeric ID namespace: immutable conversations and entries, latest tasks/submissions, and document incarnation metadata |
| `revisions` | Document versions keyed by `[documentId, commitSeq]`; bases, deltas and internal physical checkpoints |
| `requests` | Conversation-scoped submission request ownership, preserving commit ordering even for overwritten keys |
| `journal` | Checksummed resolved write batches since the last storage checkpoint |

`records`, `revisions`, `requests`, and `meta.manifest` contain an envelope:
`{ text: JSON.stringify(payload), digest: SHA256(text) }`. Records/revisions also
carry IndexedDB key/index fields outside the envelope. `journal` retains the
prototype format `{ seq, writes, digest: SHA256(seq + ":" + writes) }`.
The Mac read-only WebKit inspector can decode these stores and then parse
`text`/`writes`. `meta.databaseId` is a plain string, not an envelope.

The manifest updates only after transaction completion. ID allocation persists
its watermark before returning, including unused IDs. Strict transaction
durability is requested. A failed write publishes neither data nor sequence.
All public operations serialize with writes; close drains admitted operations,
closes the connection, then releases the lifetime Web Lock. New operations
after close reject. A second pane never steals a lock or uses an expiry lease.
Pagehide closes storage; a restored back/forward-cache page reloads.

Open reads metadata, store counts, and at most 63 retained journal batches,
one batch at a time. It does not load records or historical document revisions.
Scans read cursor pages of 64 rows (agent default, not requested) and stop at the
requested result limit plus one. Filtering can visit many pages; those pages
are not retained. Fork reads retain the ancestry chain, not ancestor transcripts.
No pagination limit is
silently substituted for the caller's limit.

Every 64 document delta batches (agent default, not requested) the backend stores
a materialized physical checkpoint at that revision's original sequence.
Historical revisions before it remain intact. A historical/current read selects the latest physical base
at or before the requested sequence and replays at most 63 following delta
batches. **The stored logical `deltasSinceBase` does not reset at a physical
checkpoint.** Only a Session-supplied base resets it. This preserves upstream
`checkpointWhen`, migrations, stored versions, and root-replacement-delta
semantics. Ordinary delta admission reads only the latest revision metadata;
materialization is required when producing a physical checkpoint.

The constants are batching policies, not retention caps. Memory still scales
with the caller's write batch, requested output, document size, individual row
size, and fork depth. There is no bound on the number of stored conversations,
entries or historical revisions, and no truncation of their content. Disk use
therefore grows with retained history; physical checkpoints trade additional
disk space for bounded replay. Current-only document revisions are also retained
on disk, although their historical content remains inaccessible under the
upstream contract. This backend does not promise bounded disk usage.

## Migration, corruption and retention

Opening a prototype version-1 database upgrades its schema and converts one
journal batch at a time. Each conversion atomically persists that batch's
materialized state, preserves its sequence, advances migration progress, and
deletes only that converted journal batch. An abort/reload resumes from the
committed progress; it cannot publish half a batch. The original allocator
watermark is preserved before opening succeeds. This is a one-way data
migration: the prototype's version-1 opener will reject the newer schema.

Metadata, sequence gaps, journal checksums, and store counts are checked on
open. Payload checksums, schemas, identity and document replay consistency are
checked as records are accessed. Consequently a damaged historical payload
can be discovered on a later historical read rather than during open. Failed
validation never clears or repairs the database, and it never substitutes an
empty conversation. Checksums detect accidental damage; they are not an
authentication boundary against code able to rewrite the database and hashes.

No age, last-open time, or recently-used-workbook list triggers deletion.
Office does not establish that a workbook was deleted merely by not opening
it. Keep offline, moved, renamed, unmounted, or otherwise unknown workbooks.
The explicit cleanup API is:

```ts
await deleteDurableWorkbookStorage({
  name: "pi-for-excel-durable:<workbookId>",
  databaseId: reviewedGeneration,
  authorization: "user-confirmed-obsolete-workbook",
});
```

Only call this after the user has reviewed the exact workbook/generation and
confirmed permanent deletion (and backup/export needs). It accepts only the
workbook-scoped namespace, obtains the same exclusive lock, verifies the
generation, and deletes that one database. Missing databases are not created
by inspection. Active owners cause `owned`, not forced closure. A stale
generation causes `replaced` and leaves the current database untouched.

If an unrelated raw IndexedDB connection blocks deletion, the promise remains
pending **while retaining the lock** until IndexedDB reports completion/error.
`deleteDatabase()` cannot be cancelled; reporting an early timeout and releasing
the lock could delete a newly opened store later. A UI timeout is not deletion
cancellation. Web Locks coordinate only cooperating clients in the same origin
and storage partition; they cannot fence external/noncooperating database code.

## Validation and Excel acceptance

`tests/browser/durable-storage.browser-test.ts` uses actual Chromium IndexedDB:
all 24 upstream conformance cases, repeated with checkpoint/reopen after each
commit; aborted commits, checkpoints and ID reservations; v1 migration restart;
checksums and missing rows; generation/missing-store failures; two-pane ownership
and reload; full history across physical checkpoints; request deduplication;
explicit retention and blocked deletion. Persistence grant/denial branches use
a StorageManager test double, not simulated IndexedDB. Other cases invoke the
browser's real StorageManager.

After integration, independently inspect a **disposable scratch workbook** in
Excel with the coordinator's read-only Mac persistence inspector. Confirm:

1. V1 migration keeps entry sequences, original document values and unused IDs.
2. More than 64 commits reduce `journal` rows without reducing entry history or
   historical document readability; the manifest agrees with store counts.
3. Reload/quit/reopen preserves the generation, conversation and pending work.
4. A second pane fails closed; ownership transfers only after the first closes.
5. Persistence status is visible; a deliberately wrong/missing expected
   generation produces an error rather than a new session.

Do not delete real workbook databases or clear WebKit data for acceptance.
Real Office/WebKit storage eviction, disk-full behavior, power loss, and live
model-driven Excel acceptance are not established by the Chromium suite.
