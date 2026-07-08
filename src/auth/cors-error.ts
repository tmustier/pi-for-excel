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
export function isCorsError(error: DynamicValue): boolean {
  if (!(error instanceof Error)) {
    return false;
  }

  const message = error.message.toLowerCase();

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
