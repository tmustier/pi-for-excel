/**
 * CORS error detection for provider connectivity checks.
 *
 * Vendored from @earendil-works/pi-web-ui 0.75.3 utils/proxy-utils
 * (MIT, © Mario Zechner, https://github.com/badlogic/pi-mono).
 * See docs/ui-ownership.md.
 */

/**
 * Check if an error is likely a CORS error.
 *
 * CORS errors in browsers typically manifest as:
 * - TypeError with message "Failed to fetch"
 * - NetworkError
 */
/** `import()` of a chunk or module that does not exist or failed to evaluate. */
export function isModuleLoadError(error: Error): boolean {
  return /dynamically imported module|importing a module script failed|error loading dynamically imported/i.test(error.message);
}

export function isCorsError(error: unknown): boolean {
  if (!(error instanceof Error)) {
    return false;
  }

  const message = error.message.toLowerCase();

  // A lazily loaded module that cannot load is not a network or CORS problem.
  if (isModuleLoadError(error)) {
    return false;
  }

  // "Failed to fetch" is the standard CORS error in most browsers
  if (error.name === "TypeError" && message.includes("failed to fetch")) {
    return true;
  }

  // Some browsers report "NetworkError"
  if (error.name === "NetworkError") {
    return true;
  }

  // CORS-specific messages
  if (message.includes("cors") || message.includes("cross-origin")) {
    return true;
  }

  return false;
}
