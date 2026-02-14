# Manual full-workbook backups

Manual full-workbook backups are an **explicit user action** fallback for high-risk edits.

They are intentionally separate from automatic per-mutation range checkpoints (`workbook_history`).

## UI entry point

- Open **Backups (Beta)** (`/history` or Backups input action/menu).
- Click **Full backup** to capture and download the current workbook as `.xlsx`.

## Commands

- `/backup` or `/backup create`
  - Captures the current workbook as a compressed file.
  - Stores it in Files workspace under `manual-backups/full-workbook/v1/...`.
  - Immediately downloads the backup file.
- `/backup list [limit]`
  - Shows a short summary of recent manual full-workbook backups for the current workbook.
- `/backup restore [id]`
  - Downloads the latest backup (or a specific backup by id) for manual restore.
- `/backup clear`
  - Deletes all manual full-workbook backups for the current workbook.

## Restore semantics

Office.js does not expose a reliable in-place API to replace the currently open workbook from a captured snapshot across hosts.

So restore is intentionally:
1. Download chosen backup file.
2. Open it in Excel.
3. Continue from that restored workbook (or copy sheets/ranges as needed).

## Storage, naming, retention

- Storage backend: Files workspace (native directory when connected, otherwise OPFS/in-memory fallback).
- Path convention: `manual-backups/full-workbook/v1/<workbook-id>/<backup-id>.xlsx`.
- Scope guardrail: list/restore/clear are filtered to the active workbook identity.
- Retention: manual backups persist until explicitly cleared (`/backup clear`) or manually deleted from Files workspace.

## Host constraints

Manual full-workbook backup requires Office `document.getFileAsync("compressed", ...)` support.

If the host does not support compressed file capture, `/backup create` fails with an actionable error and no backup is created.
