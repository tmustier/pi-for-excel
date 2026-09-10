/** Office.js document-settings backed {@link DocumentInstanceIdentity}. */

import { isDocumentInstanceToken, type DocumentInstanceIdentity } from "./document-instance.js";

/** Per-add-in, per-document settings slot. Renaming orphans tokens in existing files. */
export const OFFICE_DOCUMENT_INSTANCE_SETTING = "pi.workbookInstanceId";

/** The subset of `Office.Settings` this module relies on; injectable for tests. */
export interface OfficeDocumentSettingsLike {
  get(name: string): string | undefined;
  set(name: string, value: string): void;
  remove(name: string): void;
  saveAsync(callback: (result: { succeeded: boolean }) => void): void;
}

export interface OfficeDocumentInstanceDependencies {
  /** Returns the live settings bag, or `null` when no Office document is attached. */
  getSettings: () => OfficeDocumentSettingsLike | null;
  /** Cryptographic UUID source; `null` when none is available. */
  createToken: () => string | null;
}

function defaultGetSettings(): OfficeDocumentSettingsLike | null {
  if (typeof Office === "undefined") return null;
  // Office.js populates `context.document` only when a document is attached;
  // the taskpane also boots in browser fallback where it is absent.
  const document: Office.Document | undefined = Office.context?.document;
  const settings: Office.Settings | undefined = document?.settings;
  if (!settings) return null;
  return {
    get: (name) => {
      // Settings values are untyped in Office.js; only string tokens are accepted here.
      const raw: unknown = settings.get(name);
      return typeof raw === "string" ? raw : undefined;
    },
    set: (name, value) => settings.set(name, value),
    remove: (name) => settings.remove(name),
    saveAsync: (callback) => settings.saveAsync((result) => {
      callback({ succeeded: result.status === Office.AsyncResultStatus.Succeeded });
    }),
  };
}

function defaultCreateToken(): string | null {
  const randomUuid = globalThis.crypto?.randomUUID;
  return typeof randomUuid === "function" ? randomUuid.call(globalThis.crypto) : null;
}

export function createOfficeDocumentInstanceIdentity(
  dependencies: Partial<OfficeDocumentInstanceDependencies> = {},
): DocumentInstanceIdentity {
  const getSettings = dependencies.getSettings ?? defaultGetSettings;
  const createToken = dependencies.createToken ?? defaultCreateToken;
  let inFlight: Promise<string | null> | null = null;

  const read = (): string | null => {
    const stored = getSettings()?.get(OFFICE_DOCUMENT_INSTANCE_SETTING);
    return isDocumentInstanceToken(stored) ? stored : null;
  };

  const persistNewToken = async (): Promise<string | null> => {
    const settings = getSettings();
    if (!settings) return null;

    const token = createToken();
    if (!isDocumentInstanceToken(token)) return null;

    settings.set(OFFICE_DOCUMENT_INSTANCE_SETTING, token);
    const saved = await new Promise<boolean>((resolve) => {
      settings.saveAsync((result) => resolve(result.succeeded));
    });
    if (saved) return token;

    // Do not leave an unpersisted token where `read()` would report it as durable.
    settings.remove(OFFICE_DOCUMENT_INSTANCE_SETTING);
    return null;
  };

  return {
    read,
    ensure(): Promise<string | null> {
      const existing = read();
      if (existing) return Promise.resolve(existing);
      if (!inFlight) {
        inFlight = persistNewToken().finally(() => {
          inFlight = null;
        });
      }
      return inFlight;
    },
  };
}
