/**
 * Dev-only fetch rewrite rules for Vite proxy routes.
 *
 * Keep the list ordered from most specific to least specific prefixes.
 */

export const DEV_REWRITES: ReadonlyArray<readonly [prefix: string, proxy: string]> = [
  // OAuth token endpoints
  ["https://console.anthropic.com/", "/oauth-proxy/anthropic/"],
  ["https://github.com/", "/oauth-proxy/github/"],
  ["https://auth.openai.com/", "/api-proxy/openai-auth/"],
  ["https://oauth2.googleapis.com/", "/api-proxy/google-oauth/"],

  // API endpoints
  ["https://api.anthropic.com/", "/api-proxy/anthropic/"],
  ["https://api.openai.com/", "/api-proxy/openai/"],
  ["https://chatgpt.com/", "/api-proxy/chatgpt/"],

  // Google routes: keep Cloud Code Assist entries before generic Google route.
  ["https://daily-cloudcode-pa.sandbox.googleapis.com/", "/api-proxy/google-cloudcode-sandbox/"],
  ["https://cloudcode-pa.googleapis.com/", "/api-proxy/google-cloudcode/"],
  ["https://generativelanguage.googleapis.com/", "/api-proxy/google/"],
];

export function rewriteDevProxyUrl(url: string): string | null {
  for (const [prefix, proxy] of DEV_REWRITES) {
    if (url.startsWith(prefix)) {
      return url.replace(prefix, proxy);
    }
  }

  return null;
}
