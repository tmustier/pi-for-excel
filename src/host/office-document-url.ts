/**
 * The current document's URL/path from Office.js.
 *
 * `Office.context.document.url` is captured when the add-in initialises and is
 * not updated by Save As (observed on Excel for Mac: it stays empty for a new
 * workbook saved mid-session, even across an add-in reload, until the file is
 * closed and reopened). `Office.context.document.getFilePropertiesAsync` asks
 * the host each time and reflects the current path, in the same format. It is
 * used first; the static value is the fallback.
 *
 * The last non-empty live answer is kept: a document's path never becomes
 * empty again within one add-in lifetime, so it is a safe answer when the host
 * is stalled (a modal sheet blocks every Office.js call) and prevents the
 * identity from flapping between "unsaved" and "saved".
 */

export interface OfficeDocumentUrlSource {
  /** `Office.context.document.url`, or `null` when unavailable. */
  readStaticUrl: () => string | null;
  /** Resolves the live URL; rejects when the host reports an error. */
  readLiveUrl: () => Promise<string | null>;
  timeoutMs: number;
}

export interface OfficeDocumentUrlReader {
  read(): Promise<string | null>;
}

export const DEFAULT_LIVE_URL_TIMEOUT_MS = 3000;

function nonEmpty(value: string | null | undefined): string | null {
  return typeof value === "string" && value.trim().length > 0 ? value : null;
}

function defaultReadStaticUrl(): string | null {
  try {
    if (typeof Office === "undefined") return null;
    return nonEmpty(Office.context?.document?.url);
  } catch {
    return null;
  }
}

function defaultReadLiveUrl(): Promise<string | null> {
  if (typeof Office === "undefined") return Promise.resolve(null);
  const document = Office.context?.document;
  if (!document || typeof document.getFilePropertiesAsync !== "function") return Promise.resolve(null);
  return new Promise((resolve, reject) => {
    document.getFilePropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(nonEmpty(result.value?.url));
      } else {
        reject(new Error(result.error?.message ?? "getFilePropertiesAsync failed"));
      }
    });
  });
}

export function createOfficeDocumentUrlReader(
  overrides: Partial<OfficeDocumentUrlSource> = {},
): OfficeDocumentUrlReader {
  const source: OfficeDocumentUrlSource = {
    readStaticUrl: overrides.readStaticUrl ?? defaultReadStaticUrl,
    readLiveUrl: overrides.readLiveUrl ?? defaultReadLiveUrl,
    timeoutMs: overrides.timeoutMs ?? DEFAULT_LIVE_URL_TIMEOUT_MS,
  };
  let lastLiveUrl: string | null = null;

  const readLiveWithTimeout = (): Promise<string | null> => new Promise((resolve) => {
    const timer = setTimeout(() => resolve(null), source.timeoutMs);
    source.readLiveUrl().then(
      (url) => {
        clearTimeout(timer);
        resolve(url);
      },
      () => {
        clearTimeout(timer);
        resolve(null);
      },
    );
  });

  return {
    async read(): Promise<string | null> {
      const live = nonEmpty(await readLiveWithTimeout());
      if (live) {
        lastLiveUrl = live;
        return live;
      }
      return lastLiveUrl ?? nonEmpty(source.readStaticUrl());
    },
  };
}
