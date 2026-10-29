import type { Agent, ThinkingLevel } from "@earendil-works/pi-agent-core";

import { t } from "../language/index.js";
import { getActiveProviders } from "../models/active-providers.js";
import { showToast } from "../ui/toast.js";
import { escapeAttr, escapeHtml, setSafeInnerHTML } from "../utils/html.js";
import { formatUsageDebug, isDebugEnabled } from "../debug/debug.js";
import { estimateContextTokens } from "../utils/context-tokens.js";
import type { ExecutionMode } from "../execution/mode.js";
import {
  getStatusContextHealth,
  STATUS_CONTEXT_DESC_ATTR,
  STATUS_CONTEXT_TOKENS_ATTR,
  getStatusContextTooltipDescription,
  STATUS_CONTEXT_WARNING_ATTR,
  STATUS_CONTEXT_WARNING_SEVERITY_ATTR,
} from "./status-context.js";
import type { RuntimeLockState } from "./session-runtime-manager.js";
import { closeStatusPopover } from "./status-popovers.js";
import { getThinkingLevelLabel } from "./thinking-display.js";
import { formatThinkingShortcut } from "../utils/platform.js";

export type ActiveAgentProvider = () => Agent | null;
export type ActiveLockStateProvider = () => RuntimeLockState;
export type ActiveExecutionModeProvider = () => ExecutionMode;

function adjustContextTooltipAlignment(statusBar: HTMLElement): void {
  const trigger = statusBar.querySelector<HTMLElement>(".pi-status-ctx--trigger");
  const tooltip = trigger?.querySelector<HTMLElement>(".pi-tooltip");
  if (!trigger || !tooltip) return;

  tooltip.classList.remove("pi-tooltip--left", "pi-tooltip--right");

  const viewportWidth = document.documentElement.clientWidth;
  const triggerRect = trigger.getBoundingClientRect();
  const tooltipWidth = tooltip.offsetWidth;
  if (tooltipWidth <= 0) return;

  const centeredLeft = triggerRect.left + ((triggerRect.width - tooltipWidth) / 2);
  const centeredRight = centeredLeft + tooltipWidth;
  const edgePadding = 8;

  if (centeredRight > viewportWidth - edgePadding) {
    tooltip.classList.add("pi-tooltip--right");
    return;
  }

  if (centeredLeft < edgePadding) {
    tooltip.classList.add("pi-tooltip--left");
  }
}

const SHARED_MODEL_PROVIDER_TAGS: Record<string, { family: string; tag: string }> = {
  openai: { family: "openai", tag: "API" },
  "openai-codex": { family: "openai", tag: "ChatGPT" },
  google: { family: "google", tag: "Gemini API" },
  "google-gemini-cli": { family: "google", tag: "Code Assist" },
  "google-antigravity": { family: "google", tag: "Antigravity" },
};

function ambiguousProviderTag(provider: string): string | null {
  const entry = SHARED_MODEL_PROVIDER_TAGS[provider];
  const active = getActiveProviders();
  if (!entry || !active) return null;
  for (const other of active) {
    if (other !== provider && SHARED_MODEL_PROVIDER_TAGS[other]?.family === entry.family) return entry.tag;
  }
  return null;
}

function renderStatusBar(
  agent: Agent | null,
  lockState: RuntimeLockState,
  executionMode: ExecutionMode,
  autoCompactionEnabled: boolean,
): void {
  const el = document.getElementById("pi-status-bar");
  if (!el) return;

  if (!agent) {
    const emptyMarkup = `<span class="pi-status-ctx">${escapeHtml(t("status.no_session"))}</span>`;
    const emptySignature = "no-agent";
    if (el.getAttribute("data-status-signature") !== emptySignature) {
      setSafeInnerHTML(el, emptyMarkup, "status bar empty-state markup with escaped locale text");
      el.setAttribute("data-status-signature", emptySignature);
    }
    return;
  }

  const state = agent.state;

  const model = getActiveProviders()?.size === 0 ? null : state.model;
  const providerTag = model ? ambiguousProviderTag(model.provider) : null;
  const modelName = model ? (model.name || model.id) : t("status.connect_provider");
  const modelAlias = providerTag ? `${modelName} · ${providerTag}` : modelName;
  const modelAliasEscaped = escapeHtml(modelAlias);

  const { totalTokens, lastUsage } = estimateContextTokens(state);

  const contextWindow = state.model?.contextWindow || 200000;
  const pct = contextWindow > 0 ? Math.round((totalTokens / contextWindow) * 100) : 0;
  const ctxLabel = contextWindow >= 1_000_000
    ? `${(contextWindow / 1_000_000).toFixed(0)}M`
    : `${Math.round(contextWindow / 1000)}k`;

  const thinkingLevel = getThinkingLevelLabel(state.thinkingLevel);

  const ctxDescription = getStatusContextTooltipDescription();
  const ctxTokenDetail = t("status.context.tokens", { used: totalTokens.toLocaleString(), total: contextWindow.toLocaleString() });

  const contextHealth = getStatusContextHealth(totalTokens, contextWindow, autoCompactionEnabled);
  const ctxColor = contextHealth.colorClass;
  const ctxWarningText = contextHealth.warning?.text ?? "";
  const ctxWarningSeverity = contextHealth.warning?.severity ?? "";
  const ctxWarning = contextHealth.warning
    ? `<span class="pi-tooltip__warn pi-tooltip__warn--${contextHealth.warning.severity}">${escapeHtml(`${contextHealth.warning.text} ${contextHealth.warning.actionText}`)}</span>`
    : "";
  const chevronSvg = `<svg width="8" height="8" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>`;
  const affordanceChevronSvg = `<svg width="7" height="7" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>`;
  const brainSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 18V5"/><path d="M15 13a4.17 4.17 0 0 1-3-4 4.17 4.17 0 0 1-3 4"/><path d="M17.598 6.5A3 3 0 1 0 12 5a3 3 0 1 0-5.598 1.5"/><path d="M17.997 5.125a4 4 0 0 1 2.526 5.77"/><path d="M18 18a4 4 0 0 0 2-7.464"/><path d="M19.967 17.483A4 4 0 1 1 12 18a4 4 0 1 1-7.967-.517"/><path d="M6 18a4 4 0 0 1-2-7.464"/><path d="M6.003 5.125a4 4 0 0 0-2.526 5.77"/></svg>`;

  const debugOn = isDebugEnabled();

  const usageDebug = debugOn && lastUsage
    ? `<span class="pi-status-ctx__debug">${escapeHtml(formatUsageDebug(lastUsage))}</span>`
    : "";

  let lockBadge = "";
  if (lockState === "waiting_for_lock") {
    lockBadge = `<span class="pi-status-lock pi-status-lock--waiting" data-tooltip="${escapeAttr(t("status.lock.waiting.tooltip"))}">${escapeHtml(t("status.lock.waiting"))}</span>`;
  } else if (lockState === "holding_lock") {
    lockBadge = `<span class="pi-status-lock pi-status-lock--active" data-tooltip="${escapeAttr(t("status.lock.holding.tooltip"))}">${escapeHtml(t("status.lock.holding"))}</span>`;
  }

  const modeIsAuto = executionMode === "yolo";
  const modeBadgeClass = modeIsAuto ? " pi-status-mode--auto" : " pi-status-mode--confirm";
  const modeLabel = modeIsAuto ? t("status.mode.auto") : t("status.mode.confirm");
  const modeTooltip = modeIsAuto
    ? t("status.mode.auto.tooltip")
    : t("status.mode.confirm.tooltip");
  const modeBadge = `<button type="button" aria-haspopup="menu" class="pi-status-mode pi-status-clickable pi-status-tooltip--right${modeBadgeClass}" data-tooltip="${escapeAttr(modeTooltip)}"><span>${escapeHtml(modeLabel)}</span><span class="pi-status-affordance" aria-hidden="true">${affordanceChevronSvg}</span></button>`;

  const thinkingTooltip = escapeAttr(
    t("status.thinking.tooltip", { shortcut: formatThinkingShortcut() }),
  );

  const ctxPopoverDesc = escapeAttr(ctxDescription);
  const ctxPopoverTokens = escapeAttr(ctxTokenDetail);
  const ctxPopoverWarnText = ctxWarningText.length > 0 ? escapeAttr(ctxWarningText) : "";

  const nextMarkup = `
    <div class="pi-status-main">
      <button type="button" aria-haspopup="dialog" class="pi-status-model pi-status-clickable pi-status-tooltip--left" data-tooltip="${escapeAttr(t("status.model.tooltip"))}">
        <span class="pi-status-model__mark">π</span>
        <span class="pi-status-model__name">${modelAliasEscaped}</span>
        ${chevronSvg}
      </button>
      <button type="button" aria-haspopup="menu" class="pi-status-thinking pi-status-clickable" data-tooltip="${thinkingTooltip}" aria-label="${escapeAttr(t("status.thinking.aria", { level: thinkingLevel }))}">${brainSvg} ${escapeHtml(thinkingLevel)}<span class="pi-status-affordance" aria-hidden="true">${affordanceChevronSvg}</span></button>
      <button type="button" aria-haspopup="menu" class="pi-status-ctx pi-status-ctx--trigger pi-status-clickable has-tooltip" ${STATUS_CONTEXT_DESC_ATTR}="${ctxPopoverDesc}" ${STATUS_CONTEXT_TOKENS_ATTR}="${ctxPopoverTokens}" ${STATUS_CONTEXT_WARNING_ATTR}="${ctxPopoverWarnText}" ${STATUS_CONTEXT_WARNING_SEVERITY_ATTR}="${ctxWarningSeverity}" aria-label="${escapeAttr(t("status.context.aria", { pct, label: ctxLabel }))}"><span class="pi-status-ctx__pct ${ctxColor}">${pct}%</span><span class="pi-status-ctx__sep">/</span><span class="pi-status-ctx__limit">${ctxLabel}</span>${usageDebug}<span class="pi-status-affordance" aria-hidden="true">${affordanceChevronSvg}</span><span class="pi-tooltip"><span class="pi-tooltip__desc">${escapeHtml(ctxDescription)}</span><span class="pi-tooltip__tokens">${escapeHtml(ctxTokenDetail)}</span>${ctxWarning}</span></button>
      ${lockBadge}
    </div>
    <div class="pi-status-side">
      ${modeBadge}
    </div>
  `;

  const renderSignature = JSON.stringify({
    modelAlias,
    thinkingLevel,
    pct,
    ctxLabel,
    ctxColor,
    ctxTokenDetail,
    ctxWarningText,
    ctxWarningSeverity,
    usageDebug,
    lockState,
    executionMode,
  });

  if (el.getAttribute("data-status-signature") === renderSignature) {
    adjustContextTooltipAlignment(el);
    return;
  }

  const active = document.activeElement;
  const focusedControl = active instanceof Element && el.contains(active)
    ? STATUS_CONTROL_CLASSES.find((name) => active.classList.contains(name))
    : undefined;

  setSafeInnerHTML(el, nextMarkup, "status bar markup with escaped model and localized text");
  el.setAttribute("data-status-signature", renderSignature);
  if (focusedControl) {
    el.querySelector<HTMLElement>(`.${focusedControl}`)?.focus({ preventScroll: true });
  }
  adjustContextTooltipAlignment(el);
}

const STATUS_CONTROL_CLASSES = ["pi-status-model", "pi-status-thinking", "pi-status-ctx", "pi-status-mode"] as const;

export function updateStatusBar(
  getActiveAgent: ActiveAgentProvider,
  getLockState?: ActiveLockStateProvider,
  getExecutionMode?: ActiveExecutionModeProvider,
  autoCompactionEnabled = false,
): void {
  const activeAgent = getActiveAgent();
  const lockState = getLockState ? getLockState() : "idle";
  const executionMode = getExecutionMode ? getExecutionMode() : "yolo";
  renderStatusBar(activeAgent, lockState, executionMode, autoCompactionEnabled);
}

export function injectStatusBar(opts: {
  getActiveAgent: ActiveAgentProvider;
  getLockState?: ActiveLockStateProvider;
  getExecutionMode?: ActiveExecutionModeProvider;
  autoCompactionEnabled?: boolean;
}): () => void {
  const { getActiveAgent, getLockState, getExecutionMode, autoCompactionEnabled = false } = opts;

  let unsubscribeActiveAgent: (() => void) | undefined;
  let hasDeferredRender = false;
  let statusBarInteracting = false;
  let interactionLeaveTimer: number | null = null;

  const INTERACTION_SETTLE_MS = 180;

  const clearInteractionLeaveTimer = (): void => {
    if (interactionLeaveTimer === null) {
      return;
    }

    window.clearTimeout(interactionLeaveTimer);
    interactionLeaveTimer = null;
  };

  const flushDeferredRender = (): void => {
    if (!hasDeferredRender) {
      return;
    }

    hasDeferredRender = false;
    updateStatusBar(getActiveAgent, getLockState, getExecutionMode, autoCompactionEnabled);
  };

  const markStatusBarInteractionStart = (): void => {
    clearInteractionLeaveTimer();
    statusBarInteracting = true;
  };

  const markStatusBarInteractionEndSoon = (): void => {
    clearInteractionLeaveTimer();
    interactionLeaveTimer = window.setTimeout(() => {
      interactionLeaveTimer = null;
      statusBarInteracting = false;
      flushDeferredRender();
    }, INTERACTION_SETTLE_MS);
  };

  const isStatusBarFocused = (): boolean => {
    const active = document.activeElement;
    return active instanceof Element && active.closest("#pi-status-bar") !== null;
  };

  const shouldDeferRenderForInteraction = (): boolean => {
    return statusBarInteracting || isStatusBarFocused();
  };

  const requestRender = (): void => {
    if (shouldDeferRenderForInteraction()) {
      hasDeferredRender = true;
      return;
    }

    hasDeferredRender = false;
    updateStatusBar(getActiveAgent, getLockState, getExecutionMode, autoCompactionEnabled);
  };

  const bindActiveAgent = () => {
    unsubscribeActiveAgent?.();
    closeStatusPopover();

    const activeAgent = getActiveAgent();
    if (activeAgent) {
      unsubscribeActiveAgent = activeAgent.subscribe(requestRender);
    } else {
      unsubscribeActiveAgent = undefined;
    }

    requestRender();
  };

  const resolveStatusBarHost = (target: EventTarget | null): Element | null => {
    if (!(target instanceof Element)) {
      return null;
    }

    return target.closest("#pi-status-bar");
  };

  const onPointerOver = (event: PointerEvent): void => {
    const currentHost = resolveStatusBarHost(event.target);
    if (!currentHost) {
      return;
    }

    const relatedHost = resolveStatusBarHost(event.relatedTarget);
    if (relatedHost === currentHost) {
      return;
    }

    markStatusBarInteractionStart();
  };

  const onPointerOut = (event: PointerEvent): void => {
    const currentHost = resolveStatusBarHost(event.target);
    if (!currentHost) {
      return;
    }

    const relatedHost = resolveStatusBarHost(event.relatedTarget);
    if (relatedHost === currentHost) {
      return;
    }

    markStatusBarInteractionEndSoon();
  };

  const onFocusOut = (event: FocusEvent): void => {
    const currentHost = resolveStatusBarHost(event.target);
    if (!currentHost) {
      return;
    }

    const relatedHost = resolveStatusBarHost(event.relatedTarget);
    if (relatedHost === currentHost) {
      return;
    }

    markStatusBarInteractionEndSoon();
  };

  const onStatusUpdate = () => {
    hasDeferredRender = false;
    updateStatusBar(getActiveAgent, getLockState, getExecutionMode, autoCompactionEnabled);
  };

  document.addEventListener("pi:status-update", onStatusUpdate);
  document.addEventListener("pi:active-runtime-changed", bindActiveAgent);
  document.addEventListener("pointerover", onPointerOver, true);
  document.addEventListener("pointerout", onPointerOut, true);
  document.addEventListener("focusout", onFocusOut, true);

  requestAnimationFrame(bindActiveAgent);

  return () => {
    clearInteractionLeaveTimer();
    statusBarInteracting = false;
    hasDeferredRender = false;
    unsubscribeActiveAgent?.();
    document.removeEventListener("pi:status-update", onStatusUpdate);
    document.removeEventListener("pi:active-runtime-changed", bindActiveAgent);
    document.removeEventListener("pointerover", onPointerOver, true);
    document.removeEventListener("pointerout", onPointerOut, true);
    document.removeEventListener("focusout", onFocusOut, true);
  };
}

export function flashThinkingLevel(level: ThinkingLevel, color: string): void {
  showToast(t("status.thinking.toast", { level: getThinkingLevelLabel(level), shortcut: formatThinkingShortcut() }), 1500);

  const el = document.querySelector<HTMLElement>(".pi-status-thinking");
  if (!el) return;

  el.style.color = color;
  el.style.background = `${color}18`;
  el.style.boxShadow = `0 0 8px ${color}40`;
  el.style.transition = "none";

  let flashBar = document.getElementById("pi-thinking-flash");
  if (!flashBar) {
    flashBar = document.createElement("div");
    flashBar.id = "pi-thinking-flash";
    flashBar.style.cssText = `
      position: fixed; bottom: 0; left: 0; right: 0; height: 2px;
      pointer-events: none; z-index: 100; transition: opacity 0.6s ease-out;
    `;
    document.body.appendChild(flashBar);
  }
  flashBar.style.background = `linear-gradient(90deg, transparent, ${color}, transparent)`;
  flashBar.style.opacity = "1";

  const bar = flashBar;

  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      el.style.transition = "color 0.8s ease, background 0.8s ease, box-shadow 0.8s ease";
      el.style.color = "";
      el.style.background = "";
      el.style.boxShadow = "";
      bar.style.opacity = "0";
    });
  });
}
