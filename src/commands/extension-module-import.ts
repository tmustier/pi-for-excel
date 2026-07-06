import {
  ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY,
  isRemoteExtensionOptIn,
  type ExtensionSourceKind,
} from "./extension-source-policy.js";
function isCommandsExtensionModuleImportPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}


export type ExtensionModuleImporter = () => Promise<DynamicValue>;

function isExtensionModuleImporter(value: DynamicValue): value is ExtensionModuleImporter {
  return typeof value === "function";
}

function readBundledImportersFromGlob(): DynamicValue {
  try {
    return import.meta.glob("../extensions/*.{ts,js}");
  } catch {
    return null;
  }
}

export function resolveBundledLocalExtensionImporters(rawImporters: DynamicValue): Record<string, ExtensionModuleImporter> {
  if (!isCommandsExtensionModuleImportPayloadShape(rawImporters)) {
    return {};
  }

  const importers: Record<string, ExtensionModuleImporter> = {};

  for (const [path, importer] of Object.entries(rawImporters)) {
    if (!isExtensionModuleImporter(importer)) {
      continue;
    }

    importers[path] = importer;
  }

  return importers;
}

const BUNDLED_LOCAL_EXTENSION_IMPORTERS = resolveBundledLocalExtensionImporters(readBundledImportersFromGlob());

export function getLocalExtensionImportCandidates(specifier: string): string[] {
  const normalized = specifier.trim();
  const candidates = new Set<string>([normalized]);

  if (normalized.endsWith(".js")) {
    candidates.add(`${normalized.slice(0, -3)}.ts`);
  } else if (normalized.endsWith(".ts")) {
    candidates.add(`${normalized.slice(0, -3)}.js`);
  } else {
    candidates.add(`${normalized}.ts`);
    candidates.add(`${normalized}.js`);
  }

  return Array.from(candidates);
}

interface ImportExtensionModuleOptions {
  bundledImporters?: Record<string, ExtensionModuleImporter>;
  dynamicImport?: (specifier: string) => Promise<DynamicValue>;
  isDev?: boolean;
}

function getDefaultDynamicImport(): (specifier: string) => Promise<DynamicValue> {
  return (specifier: string) => import(/* @vite-ignore */ specifier);
}

function readIsDevDefault(): boolean {
  try {
    return import.meta.env.DEV === true;
  } catch {
    return false;
  }
}

export async function importExtensionModule(
  specifier: string,
  sourceKind: ExtensionSourceKind,
  options: ImportExtensionModuleOptions = {},
): Promise<DynamicValue> {
  const bundledImporters = options.bundledImporters ?? BUNDLED_LOCAL_EXTENSION_IMPORTERS;
  const dynamicImport = options.dynamicImport ?? getDefaultDynamicImport();
  const isDev = options.isDev ?? readIsDevDefault();

  if (sourceKind === "local-module") {
    for (const candidate of getLocalExtensionImportCandidates(specifier)) {
      const importer = bundledImporters[candidate];
      if (!importer) {
        continue;
      }

      return importer();
    }

    if (isDev) {
      return dynamicImport(specifier);
    }

    throw new Error(
      `Local extension module "${specifier}" was not bundled. `
      + "Use a bundled module under src/extensions, paste code, or a remote URL (with explicit opt-in).",
    );
  }

  return dynamicImport(specifier);
}

export function readRemoteExtensionOptInFromStorage(storage: Storage | null | undefined): boolean {
  if (!storage) {
    return false;
  }

  try {
    const raw = storage.getItem(ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY);
    return isRemoteExtensionOptIn(raw);
  } catch {
    return false;
  }
}
