/**
 * The error banner text for a tab whose turn has settled.
 */

import type { AgentMessage } from "@earendil-works/pi-agent-core";

import { t } from "../language/index.js";
import { PROXY_HELPER_DOCS_URL } from "../auth/proxy-validation.js";
import { isLikelyCorsErrorMessage } from "./runtime-utils.js";

export interface TurnFailureContext {
  messages: readonly AgentMessage[];
  /** The run ended in a context overflow (for the active model). */
  isContextOverflow: boolean;
  /** Pi already compacted and retried once, and the retry overflowed too. */
  overflowAfterRecovery: boolean;
  /** The user dismissed the connection dialog for this provider during the turn. */
  missingCredentialsFor: string | null;
  online: boolean;
  /** The configured proxy URL when the proxy is enabled. */
  proxyUrl: string | null;
  defaultProxyUrl: string;
}

/**
 * Returns the banner for the turn's final model response, or null when the
 * turn did not fail. A stop is never an error: it is recognized by the
 * response's stop reason, not by words like "abort" in a provider message.
 */
export function describeTurnFailure(context: TurnFailureContext): string | null {
  const last = context.messages[context.messages.length - 1];
  if (!last || last.role !== "assistant" || last.stopReason !== "error") return null;
  const error = last.errorMessage ?? "";

  if (context.missingCredentialsFor) {
    return t("turn.error.apiKeyRequired", { provider: context.missingCredentialsFor });
  }
  if (context.isContextOverflow) {
    return context.overflowAfterRecovery
      ? t("turn.error.contextOverflowAfterRetry")
      : t("turn.error.contextOverflow");
  }
  if (isLikelyCorsErrorMessage(error)) {
    if (!context.online) return t("turn.error.offline");
    if (context.proxyUrl) return t("turn.error.proxyUnreachable", { url: context.proxyUrl });
    return t("turn.error.cors", { url: context.defaultProxyUrl, guide: PROXY_HELPER_DOCS_URL });
  }
  return t("init.llmError", { error });
}
