/**
 * Test-only fallback for import.meta.glob raw markdown loading.
 *
 * Browser/runtime builds use Vite's import.meta.glob transform.
 * Node test runs execute source directly, so register-test-ts-loader.mjs
 * injects a glob loader onto globalThis.
 */

interface RawMarkdownGlobLoader {
  (pattern: string, importerUrl: string): Record<string, string>;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isRawMarkdownRecord(value: unknown): value is Record<string, string> {
  if (!isRecord(value)) {
    return false;
  }

  for (const entryValue of Object.values(value)) {
    if (typeof entryValue !== "string") {
      return false;
    }
  }

  return true;
}

declare global {
  var __PI_TEST_RAW_MARKDOWN_GLOB: RawMarkdownGlobLoader | undefined;
}

function readRawMarkdownGlobLoaderFromGlobalScope(): RawMarkdownGlobLoader | null {
  const candidate = globalThis.__PI_TEST_RAW_MARKDOWN_GLOB;
  if (typeof candidate !== "function") {
    return null;
  }

  return candidate;
}

export function loadRawMarkdownFromTestGlob(pattern: string, importerUrl: string): Record<string, string> {
  const loader = readRawMarkdownGlobLoaderFromGlobalScope();
  if (!loader) {
    return {};
  }

  const loaded = loader(pattern, importerUrl);
  if (!isRawMarkdownRecord(loaded)) {
    return {};
  }

  return loaded;
}
