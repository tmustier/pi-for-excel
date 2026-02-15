/**
 * Proxy URL validation for Office taskpanes.
 *
 * Office add-ins are served over HTTPS. Some Office webviews (notably WKWebView on macOS)
 * will block calls to an HTTP proxy from an HTTPS taskpane (mixed content), surfacing as
 * "Load failed" / "Connection error".
 */

export const DEFAULT_LOCAL_PROXY_URL = "https://localhost:3003";
export const PROXY_HELPER_DOCS_URL =
  "https://github.com/tmustier/pi-for-excel/blob/main/docs/install.md#oauth-logins-and-cors-proxy";

export function normalizeProxyUrl(url: string): string {
  return url.trim().replace(/\/+$/, "");
}

export function validateOfficeProxyUrl(url: string): string {
  const normalized = normalizeProxyUrl(url);

  if (!/^https?:\/\//i.test(normalized)) {
    throw new Error(
      `Invalid Proxy URL: "${url}". Expected a full URL like ${DEFAULT_LOCAL_PROXY_URL}`,
    );
  }

  let parsed: URL;
  try {
    parsed = new URL(normalized);
  } catch {
    throw new Error(
      `Invalid Proxy URL: "${url}". Expected a full URL like ${DEFAULT_LOCAL_PROXY_URL}`,
    );
  }

  // Mixed content guardrail: HTTPS taskpane -> HTTP proxy.
  // This tends to fail in Office webviews (macOS), so fail fast with guidance.
  if (typeof window !== "undefined" && window.location?.protocol === "https:" && parsed.protocol === "http:") {
    throw new Error(
      `Proxy URL is HTTP (${normalized}) but the add-in is served over HTTPS. Office webviews may block this as mixed content. ` +
        `Use ${DEFAULT_LOCAL_PROXY_URL} and run a local HTTPS proxy. See ${PROXY_HELPER_DOCS_URL}.`,
    );
  }

  return normalized;
}

function isLoopbackHostname(hostname: string): boolean {
  const h = hostname.toLowerCase();
  if (h === "localhost") return true;
  if (h === "::1" || h === "0:0:0:0:0:0:0:1") return true;
  if (h.startsWith("127.")) return true;
  if (h.startsWith("::ffff:127.")) return true;
  return false;
}

/**
 * Returns true if the proxy URL points at a loopback/localhost address.
 * Useful for warning users when they configure a remote proxy.
 */
export function isLoopbackProxyUrl(url: string): boolean {
  const normalized = normalizeProxyUrl(url);
  try {
    const parsed = new URL(normalized);
    return isLoopbackHostname(parsed.hostname);
  } catch {
    return false;
  }
}
