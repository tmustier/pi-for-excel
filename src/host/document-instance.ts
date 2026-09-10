/**
 * Add-in-scoped token stored inside the host document.
 *
 * Used only while a document has no path-derived identity (a new workbook that
 * has never been saved). The token lives in the document, so it survives
 * taskpane reloads and close/reopen of the pane. It is copied along with the
 * document by Save As / Save a Copy, which is why it must never be treated as
 * an alias for a path-derived `workbookId`: once a canonical identity exists,
 * the token is ignored.
 */
export interface DocumentInstanceIdentity {
  /** Read the stored token. Never creates one. */
  read(): string | null;

  /**
   * Return the stored token, creating and persisting one when absent.
   * Resolves `null` when the host could not persist it; callers must treat
   * that as "no identity" rather than retrying with an unpersisted value.
   */
  ensure(): Promise<string | null>;
}

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/u;

/** Tokens are opaque UUIDs; anything else in the slot is treated as absent. */
export function isDocumentInstanceToken(value: string | null | undefined): value is string {
  return typeof value === "string" && UUID_PATTERN.test(value);
}
