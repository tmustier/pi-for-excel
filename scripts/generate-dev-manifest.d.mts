/**
 * Type declarations for generate-dev-manifest.mjs so vite.config.ts can share
 * its strict dev-proxy origin validation (single source of truth).
 */

export declare const DEV_BASE_URL: string;
export declare const DEFAULT_DEV_PROXY_HOST: string;

export interface ResolveDevOriginOptions {
  arg?: string;
  env?: { DEV_HOST?: string; PORTLESS_URL?: string };
}

export interface ResolvedDevOrigin {
  origin: string;
  source: "argument" | "DEV_HOST" | "PORTLESS_URL" | "default";
}

/**
 * Resolve and strictly validate the HTTPS origin for the dev proxy.
 * Throws on anything that is not a clean https origin (no credentials,
 * path, query, or hash) or on the default dev URL itself.
 */
export declare function resolveDevOrigin(options?: ResolveDevOriginOptions): ResolvedDevOrigin;

/** Replace every occurrence of the default dev base URL with the proxy origin. */
export declare function renderDevManifest(xml: string, origin: string): string;
