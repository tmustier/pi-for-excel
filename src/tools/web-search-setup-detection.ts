import { DEFAULT_LOCAL_PROXY_URL, probeProxyReachability } from "../auth/proxy-validation.js";
import { getEnabledProxyBaseUrl } from "./external-fetch.js";
import {
  getApiKeyForProvider,
  isApiKeyRequired,
  loadWebSearchProviderConfig,
  WEB_SEARCH_PROVIDERS,
  type WebSearchConfigStore,
  type WebSearchProvider,
  type WebSearchProviderConfig,
} from "./web-search-config.js";
import type { WebSearchDetails } from "./tool-details.js";

export type WebSearchSetupMode =
  | { type: "needs_key" }
  | { type: "needs_proxy" }
  | { type: "needs_both" }
  | { type: "wrong_provider"; availableProvider: WebSearchProvider }
  | { type: "generic_error" };

export interface WebSearchSetupContext {
  mode: WebSearchSetupMode;
  provider: WebSearchProvider;
  providerConfig: WebSearchProviderConfig;
  proxyBaseUrl: string | undefined;
}

export interface WebSearchSetupDetectionDeps {
  isDev?: boolean;
  probeProxyReachability?: (proxyUrl: string, timeoutMs?: number) => Promise<boolean>;
}

function isWebSearchProvider(value: string): value is WebSearchProvider {
  return WEB_SEARCH_PROVIDERS.some((provider) => provider === value);
}

function findAlternativeProvider(
  config: WebSearchProviderConfig,
  currentProvider: WebSearchProvider,
): WebSearchProvider | undefined {
  for (const provider of WEB_SEARCH_PROVIDERS) {
    if (provider === currentProvider) {
      continue;
    }

    if (getApiKeyForProvider(config, provider)) {
      return provider;
    }
  }

  return undefined;
}

export async function detectWebSearchSetupContext(
  details: WebSearchDetails,
  settings: WebSearchConfigStore,
  deps?: WebSearchSetupDetectionDeps,
): Promise<WebSearchSetupContext> {
  const [providerConfig, proxyBaseUrl] = await Promise.all([
    loadWebSearchProviderConfig(settings),
    getEnabledProxyBaseUrl(settings),
  ]);

  const provider = isWebSearchProvider(details.provider)
    ? details.provider
    : providerConfig.provider;

  const hasKey = Boolean(getApiKeyForProvider(providerConfig, provider));
  const needsKey = !hasKey && isApiKeyRequired(provider);

  const isDev = deps?.isDev === true;
  const probeProxy = deps?.probeProxyReachability ?? probeProxyReachability;

  let needsProxy = details.proxyDown === true;
  const shouldProbeProxy = !needsProxy && needsKey && !isDev && proxyBaseUrl !== undefined;

  if (shouldProbeProxy) {
    const probeUrl = proxyBaseUrl ?? DEFAULT_LOCAL_PROXY_URL;
    const proxyReachable = await probeProxy(probeUrl, 1500);
    needsProxy = !proxyReachable;
  }

  if (needsKey && !needsProxy) {
    const alternativeProvider = findAlternativeProvider(providerConfig, provider);
    if (alternativeProvider) {
      return {
        mode: { type: "wrong_provider", availableProvider: alternativeProvider },
        provider,
        providerConfig,
        proxyBaseUrl,
      };
    }
  }

  let mode: WebSearchSetupMode;
  if (needsKey && needsProxy) {
    mode = { type: "needs_both" };
  } else if (needsKey) {
    mode = { type: "needs_key" };
  } else if (needsProxy) {
    mode = { type: "needs_proxy" };
  } else {
    mode = { type: "generic_error" };
  }

  return {
    mode,
    provider,
    providerConfig,
    proxyBaseUrl,
  };
}
