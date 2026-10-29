export declare const DEV_SERVER_PORT: number;
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

export declare function resolveDevOrigin(options?: ResolveDevOriginOptions): ResolvedDevOrigin;

export declare function renderDevManifest(xml: string, origin: string): string;
