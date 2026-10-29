import type { AgentMessage } from "@earendil-works/pi-agent-core";

import { t } from "../language/index.js";
import { PROXY_HELPER_DOCS_URL } from "../auth/proxy-validation.js";
import { isLikelyCorsErrorMessage } from "./runtime-utils.js";

export interface TurnFailureContext {
  messages: readonly AgentMessage[];
  isContextOverflow: boolean;
  overflowAfterRecovery: boolean;
  missingCredentialsFor: string | null;
  expiredLoginFor?: string | null;
  online: boolean;
  proxyUrl: string | null;
  defaultProxyUrl: string;
  proxyBlockedByBrowser: boolean;
}

export function describeTurnFailure(context: TurnFailureContext): string | null {
  const last = context.messages[context.messages.length - 1];
  if (!last || last.role !== "assistant" || last.stopReason !== "error") return null;
  const error = last.errorMessage ?? "";

  if (/\b(401|403)\b|unauthorized|invalid.*(?:key|token)/iu.test(error)) return t("turn.error.credentials");
  if (/\b404\b|model_not_found/iu.test(error)) return t("turn.error.model");
  if (context.expiredLoginFor) {
    return t("turn.error.loginExpired", { provider: context.expiredLoginFor });
  }
  if (context.missingCredentialsFor) {
    return t("turn.error.apiKeyRequired", { provider: context.missingCredentialsFor });
  }
  if (context.isContextOverflow) {
    return context.overflowAfterRecovery
      ? t("turn.error.contextOverflowAfterRetry")
      : t("turn.error.contextOverflow");
  }
  if (last.provider === "github-copilot" && /\bmodel_not_supported\b/u.test(error)) {
    return t("turn.error.copilotModelUnavailable");
  }
  if (isLikelyCorsErrorMessage(error)) {
    if (!context.online) return t("turn.error.offline");
    if (context.proxyBlockedByBrowser) return t("proxy.blocked_by_browser");
    if (context.proxyUrl) return t("turn.error.proxyUnreachable", { url: context.proxyUrl });
    return t("turn.error.cors", { url: context.defaultProxyUrl, guide: PROXY_HELPER_DOCS_URL });
  }
  return t("turn.error.generic");
}

export function turnFailureFix(error: string): { command: "login" | "settings" | "model"; label: string } {
  if (/\b(401|403)\b|unauthorized|key|token|login|credential/iu.test(error)) {
    return { command: "login", label: t("turn.fix.credentials") };
  }
  if (isLikelyCorsErrorMessage(error) || /proxy|offline|network/iu.test(error)) {
    return { command: "settings", label: t("turn.fix.proxy") };
  }
  return { command: "model", label: t("turn.fix.model") };
}
