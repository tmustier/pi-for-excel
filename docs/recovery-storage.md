# Local workbook backup storage

Approved policy, 2026-10-27. This supersedes the earlier proposal of approximately
256 MB with a newest-few floor per workbook. The goal is bounded latency and
resource use, not expiring useful history.

## Contract

- Each backup is an immutable gzip Blob. Its separate TypeBox-checked metadata
  carries id, time, workbook identity, tool/call, range, counts, kind, optional
  `turnId` and restored-from id, and logical stored bytes.
- Ceiling: **512 MiB (536,870,912 bytes)** per origin across all workbooks.
  Accounting includes compressed Blob bytes plus UTF-8 JSON metadata bytes;
  physical index/SQLite overhead is not included. Append and eviction commit
  atomically with the byte counter. At or below the ceiling nothing expires;
  above it the oldest `(time, id)` records go first. There is no count/age
  limit or protected workbook floor. A single oversized backup is refused
  before deleting anything. Host quota failures abort the whole transaction.
- Listing uses compound indexes and exclusive `(time, id)` cursors, 50 rows
  maximum, with one lookahead record. It never loads payloads. Backups and
  `/history` show one page at a time; search/action filters explicitly apply
  to that page. Oldest/newest sort selects the index direction. The tool
  `workbook_history` accepts/returns the same cursor. Turn-filtered pages use
  `(workbookId, turnId, time, id)` for CL-01.
- Restore and its later-edits preview load one selected payload. Identity is
  checked before decoding or touching the workbook; the existing live-state,
  destructive-restore and later-edits guards remain. `afterFormulas` survives
  migration. Restores still create inverse backups and warn if that fails.
- Delete and Clear all use metadata indexes and delete payloads by key. Clear
  all means every page of the current workbook, not the 50 visible rows.
- Save As, rename and move observed by the Office document-bound identity
  watcher move metadata keys/labels, never copy/wipe/recompress payloads. The
  destination's existing backups are retained. First-save document-token
  handling remains supported. A pane following the active WPS workbook must
  not mistake switching workbooks for Save As. Offline moves while no pane
  observes the old/new path cannot be inferred safely; old backups remain
  under the old identity and still count toward global oldest-first eviction.
- The separate bounded change-audit log is unchanged; it records restore
  operations, not a second copy of recovery state.

## Migration and compatibility

The existing `pi-for-excel` IndexedDB database advances from version 2 to 3.
The versionchange transaction creates empty stores and indexes only. Once open,
the recovery repository reads the legacy settings value
`workbook.recovery-snapshots.v1`, validates it, and compresses entries sequentially
outside an IDB transaction. A final readwrite transaction checks for a concurrent
winner and an unchanged source, inserts records, writes accounting and a marker,
then removes the old value and obsolete retention setting **in that same commit**.
No partial copy/marker/delete survives an abort. IDs and state are preserved.
Malformed, duplicate-id or over-budget archives are left untouched rather than
silently trimmed. A failed migration permits legacy read/restore preview, while
new persistence fails visibly and can retry later. Sessions/settings still work.
Migration is the one-time exception to bounded archive reads/memory: it must
read the old monolith, and temporarily needs room for old plus compressed data.

Old panes close on versionchange and show the existing reload-required message.
Older/cached builds cannot open v3 (`VersionError`): database downgrades are not
supported. On Mac, an older bundle displays the raw "Initialization failed:
An attempt was made to open a database using a lower version than the existing
version" error, not the newer pane's reload-required notice. **Developers should
test on an isolated origin**: upgrading their normal origin also locks every
older branch out of chats/settings/sign-ins there. Clearing WebKit data removes
the version lock by deleting all that data; it is not a safe rollback strategy.
Backups are captured before mutations but persisted **after** them;
this redesign does not turn quota/closed-pane failures into pre-write guarantees.
An unavailable backup is reported honestly in the tool result and toast. A
failed inverse save cannot undo an already completed restore.

## Measurements, not sizing estimates

Excel on the web, Chromium in an Amp orb, actual Office.js readbacks of 20,000
cells, including before values/formulas and WT-02 `afterFormulas`:

| Data | JSON/UTF-8 bytes | gzip bytes | Reduction | Mean gzip wall/Chrome-service CPU ms |
|---|---:|---:|---:|---:|
| Numbers | 568,261 | 185,761 | 3.06× | 23 / 25 |
| Formulas | 806,295 | 170,349 | 4.73× | 16 / 17 |
| Short text | 1,269,283 | 150,257 | 8.45× | 10 / 12 |
| Repeated 1,000-character text | 40,349,283 | ≈258,663 | 156× | 144 / 199 |

Twenty compression runs; service CPU includes background Chrome noise. Repeated
text is deliberately highly compressible, **not** a typical ratio for arbitrary
long strings. Deflate-raw saved only 18 bytes versus gzip. The long-text case
also spent about 450 ms encoding, with about 142 ms maximum timer lag during
gzip; cost still depends on the selected backup, not accumulated history.

Current single-record append at 124.77 MB: median 965 ms. Native metadata index
pages of 50 at 1k/10k/50k rows: about 3 ms; indexed append plus byte counter:
about 1.1 ms. Chromium rejected serialized object/string records above
133,169,152 bytes; 250 MiB Blob roundtrips succeeded.

Real Excel for Mac 16.114.2, macOS 26.3.1, WebKit 21623.2.7.11.7, idle pane:
the old object read-modify-write cost 566–586 ms at 150 MB and 1,227–1,407 ms
at 250 MB. Blob payloads of 70.1/164.1/273.5 MiB roundtripped, including after
pane close/reopen. The Blob files had exactly their UTF-8 size; indexed records
were only hundreds of bytes. Deleting one Blob freed its file immediately.
String payloads used UTF-16 pages and deletion did not shrink SQLite
(`auto_vacuum=0`), another reason not to keep payloads in metadata records.
Mixed Mac data compressed 4.41–4.47×: 70.1 MiB → 15.9 MiB in 627 ms;
273.5 MiB → 61.1 MiB in 2,496 ms. Decompression was about four times faster.
These were isolated disposable databases, not the user's real backups.

## Why 512 MiB, and what it cannot guarantee

Web-pane quota was approximately 10 GiB; persistent storage was not granted.
Mac `estimate().quota` reported approximately 19.2 GiB on a 98%-full disk with
about 20.8 GiB free. 1.5 GiB of incompressible records committed without error;
the real limit/disk-full behaviour was deliberately not tested. Mac
`estimate().usage` only increased, reaching 4.3 GB after records/databases were
deleted while disk usage had returned to 21 MB; **never use it for accounting**.
`persist()` returned false. Windows WebView2 was not directly measured; the
Chromium Blob test establishes transport behaviour, not its host quota.

512 MiB is 5% of the measured web quota and roughly 2.5% of the Mac's then-free
space. At the measured mixed-data ratio it holds about 2.2 GiB raw. It is a
fixed conservative safety ceiling, not a promise of indefinite retention or a
quota reservation. 1 GiB would not materially alter LRU selection but would
consume another 512 MiB on a nearly full disk. Keeping Blobs out of SQLite also
lets eviction release disk rather than leaving the archive's high-water mark.

WebKit's published [storage policy](https://webkit.org/blog/14403/updates-to-storage-policy/)
allows non-browser apps about 15% of disk per origin, 20% across origins, but
system disk pressure can trigger eviction sooner. Eviction is whole-origin and
LRU-based, not backup-aware. An active page is normally protected; persistence
is not available here. Quota/pressure/unused-origin behaviour depends on the
host; no ceiling can guarantee protection. **Origin eviction or clearing Excel's
WebKit cache also removes chats, sign-ins, settings and rules**, not just backups.
512 MiB adds to disk pressure but does not itself select Pi for eviction.
Durable recovery still requires downloading a full workbook backup.

Unmeasured: actual disk-full commit behaviour, peak heap at large payloads,
Windows desktop WebView2 quota, and real long-term host eviction. No exact
physical-size/eviction guarantee is claimed.

## Implementation acceptance

- Deterministic suites: 1,001 overall, 731 context, 69 recovery, 97 security,
  45 models. Check and production build pass. Native Chromium browser suite:
  118 tests, including two panes migrating a v2 archive concurrently, preserving
  `afterFormulas`, then appending without losses, paging 71 entries as 50 + 21,
  and clearing every page. Injected migration failure rolls back the copy,
  marker and source deletion; ordinary settings remain writable.
- Excel on the web: a real GPT-6.1 Sol turn called `read_range` and `write_cells`
  on a scratch 2×2 range. Independent Office.js readback confirmed the changed
  values; the stored gzip Blob was 385 bytes. `/revert` restored all four
  original values, including non-ASCII text. Original and inverse backups
  decoded after a taskpane reload. A separate later edit (909) triggered
  "Overwrite later edits?"; choosing "Keep my edits" preserved 909.
- Excel for Mac / WKWebView, isolated `https://localhost:3142` and separate
  add-in ID, production build of initial implementation `5c61b21d`: four real
  legacy backups from two workbooks migrated with all IDs/times/identities
  intact. Direct reads of WebKit's SQLite and Blob files independently confirmed
  four gzip payloads and 3,068 accounted bytes; the source key was gone. Real
  GPT-6 Luna writes, restore/inverse backups, restore of a migrated payload,
  and Excel relaunch persistence passed.
- Mac Save As to a new path moved six metadata rows, leaving the other
  workbook untouched. Save As over an existing destination retained its backup
  and merged all seven. Blob inodes, sizes and mtimes were unchanged, as was
  accounting. Deleting one backup freed its file immediately and subtracted
  exactly 771 bytes; Clear all left zero metadata/payload/files/accounted bytes.
  Thomas's normal `:3141` origin was never opened by this test build.
- Mac confirmed the older-bundle forward lock described above. Mac did not
  force failed migrations, quota exhaustion, old-pane upgrade handoff, external
  rename/move, or large payloads in this implementation. Failure/concurrency
  coverage comes from deterministic/native Chromium tests, not that Mac run.

An already-open Backups panel is a loaded page, not a live subscription. Use
Refresh or close/reopen after changes made outside it; rerunning `/history`
while the panel is open does not refresh it. This PR does not add cross-pane UI
subscriptions.
