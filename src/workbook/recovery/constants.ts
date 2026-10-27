/** Upper bound (and default) for "Keep at most N backups" per workbook. */
export const MAX_RECOVERY_ENTRIES = 120;
/**
 * Storage bound across all workbooks: the oldest backups, whichever workbook
 * they belong to, go first once the shared list holds more than this.
 */
export const MAX_STORED_RECOVERY_ENTRIES = 2 * MAX_RECOVERY_ENTRIES;
export const MAX_RECOVERY_CELLS = 20_000;

/** Settings key for user-configured retention limit (max snapshots to keep). */
export const RETENTION_LIMIT_SETTING_KEY = "workbook.recovery.retentionLimit.v1";

/** Absolute minimum for user-configured retention limit. */
export const MIN_RETENTION_LIMIT = 5;

/** Clamp a user-provided retention limit to valid bounds. */
export function clampRetentionLimit(raw: unknown): number {
  if (typeof raw !== "number" || !Number.isFinite(raw)) return MAX_RECOVERY_ENTRIES;
  const rounded = Math.floor(raw);
  if (rounded < MIN_RETENTION_LIMIT) return MIN_RETENTION_LIMIT;
  if (rounded > MAX_RECOVERY_ENTRIES) return MAX_RECOVERY_ENTRIES;
  return rounded;
}
