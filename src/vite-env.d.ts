/// <reference types="vite/client" />

/**
 * Build-time configuration for org/central deployments (docs/central-proxy.md).
 * Both are optional; unset means stock behavior.
 */
interface ImportMetaEnv {
  /** Default proxy URL baked into org builds (https:// only), e.g. "https://pi-proxy.example.com:3003". */
  readonly VITE_PI_DEFAULT_PROXY_URL?: string;
  /** Comma-separated provider ids to show in the connect UI, e.g. "openai,deepseek". */
  readonly VITE_PI_ALLOWED_PROVIDERS?: string;
}
