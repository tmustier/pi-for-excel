/**
 * Status bar rendering + thinking level flash.
 */

import type { Agent } from "@mariozechner/pi-agent-core";

import { showToast } from "../ui/toast.js";
import { escapeAttr, escapeHtml } from "../utils/html.js";
import { formatUsageDebug, isDebugEnabled } from "../debug/debug.js";
import { estimateContextTokens } from "../utils/context-tokens.js";
import {
  ACTIVE_INTEGRATIONS_TOOLTIP_PREFIX,
  formatIntegrationCountLabel,
} from "../integrations/naming.js";
import type { RuntimeLockState } from "./session-runtime-manager.js";

export type ActiveAgentProvider = () => Agent | null;
export type ActiveLockStateProvider = () => RuntimeLockState;
export type ActiveRulesProvider = () => boolean;
export type ActiveIntegrationsProvider = () => string[];

function renderStatusBar(
  agent: Agent | null,
  lockState: RuntimeLockState,
  rulesActive: boolean,
  activeIntegrations: string[],
): void {
  const el = document.getElementById("pi-status-bar");
  if (!el) return;

  if (!agent) {
    el.innerHTML = `<span class="pi-status-ctx">No active session</span>`;
    return;
  }

  const state = agent.state;

  // Model alias
  const model = state.model;
  const modelAlias = model ? (model.name || model.id) : "Select model";
  const modelAliasEscaped = escapeHtml(modelAlias);

  // Context usage
  //
  // For providers with prompt caching (e.g. Anthropic), `usage.input` excludes cached
  // prompt tokens. Cached tokens still count towards the model's context window.
  //
  // The most reliable signal we have in the UI is the last successful assistant
  // turn's usage, which already reflects the prompt size.
  const { totalTokens, lastUsage } = estimateContextTokens(state);

  const contextWindow = state.model?.contextWindow || 200000;
  const pct = contextWindow > 0 ? Math.round((totalTokens / contextWindow) * 100) : 0;
  const ctxLabel = contextWindow >= 1_000_000
    ? `${(contextWindow / 1_000_000).toFixed(0)}M`
    : `${Math.round(contextWindow / 1000)}k`;

  // Thinking level
  const thinkingLabels: Record<string, string> = {
    off: "off", minimal: "min", low: "low", medium: "med", high: "high", xhigh: "max",
  };
  const thinkingLevel = thinkingLabels[state.thinkingLevel] || state.thinkingLevel;

  // Context health: color + tooltip based on usage
  let ctxColor = "";
  const ctxBaseTooltip = `How much of Pi's memory (context window) the conversation is using — ${totalTokens.toLocaleString()} / ${contextWindow.toLocaleString()} tokens. As this fills up, Pi may lose track of earlier details — start a new chat if quality drops.`;
  let ctxWarning = "";
  let ctxWarningText = "";
  if (pct > 100) {
    ctxColor = "pi-status-ctx--red";
    ctxWarningText = "Context is full — the next message will fail. Use /compact to summarize earlier messages, or /new for a fresh chat.";
    ctxWarning = `<span class="pi-tooltip__warn pi-tooltip__warn--red">${ctxWarningText}</span>`;
  } else if (pct > 60) {
    ctxColor = "pi-status-ctx--red";
    ctxWarningText = `Context ${pct}% full — responses will get less accurate. Use /compact to free space, or /new for a fresh chat.`;
    ctxWarning = `<span class="pi-tooltip__warn pi-tooltip__warn--red">${ctxWarningText}</span>`;
  } else if (pct > 40) {
    ctxColor = "pi-status-ctx--yellow";
    ctxWarningText = `Context ${pct}% full. Consider /compact to free space, or /new for a fresh chat.`;
    ctxWarning = `<span class="pi-tooltip__warn pi-tooltip__warn--yellow">${ctxWarningText}</span>`;
  }

  const ctxPopoverText = escapeAttr(
    ctxWarningText.length > 0 ? `${ctxBaseTooltip} ${ctxWarningText}` : ctxBaseTooltip,
  );

  const chevronSvg = `<svg width="8" height="8" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>`;
  const affordanceChevronSvg = `<svg width="7" height="7" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>`;
  const brainSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 18V5"/><path d="M15 13a4.17 4.17 0 0 1-3-4 4.17 4.17 0 0 1-3 4"/><path d="M17.598 6.5A3 3 0 1 0 12 5a3 3 0 1 0-5.598 1.5"/><path d="M17.997 5.125a4 4 0 0 1 2.526 5.77"/><path d="M18 18a4 4 0 0 0 2-7.464"/><path d="M19.967 17.483A4 4 0 1 1 12 18a4 4 0 1 1-7.967-.517"/><path d="M6 18a4 4 0 0 1-2-7.464"/><path d="M6.003 5.125a4 4 0 0 0-2.526 5.77"/></svg>`;
  // Lucide "file-text" — rules/instructions icon
  const rulesSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/></svg>`;
  // Lucide "plug" — integrations icon
  const integrationsSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22v-5"/><path d="M9 8V2"/><path d="M15 8V2"/><path d="M18 8v5a6 6 0 0 1-6 6a6 6 0 0 1-6-6V8Z"/></svg>`;

  const debugOn = isDebugEnabled();

  const usageDebug = debugOn && lastUsage
    ? `<span class="pi-status-ctx__debug">${escapeHtml(formatUsageDebug(lastUsage))}</span>`
    : "";

  let lockBadge = "";
  if (lockState === "waiting_for_lock") {
    lockBadge = `<span class="pi-status-lock pi-status-lock--waiting" data-tooltip="A workbook write is queued behind another session.">lock…</span>`;
  } else if (lockState === "holding_lock") {
    lockBadge = `<span class="pi-status-lock pi-status-lock--active" data-tooltip="This session currently holds the workbook write lock.">lock</span>`;
  }

  const rulesActiveClass = rulesActive ? " pi-status-rules--active" : "";
  const rulesTooltip = rulesActive
    ? "Custom rules are active. Click to edit rules &amp; number format."
    : "Click to edit rules &amp; number format preferences.";
  const rulesBadge = `<button type="button" class="pi-status-rules pi-status-clickable${rulesActiveClass}" data-tooltip="${rulesTooltip}">${rulesSvg}<span>rules</span><span class="pi-status-affordance" aria-hidden="true">${affordanceChevronSvg}</span></button>`;

  const integrationsBadge = activeIntegrations.length > 0
    ? `<button type="button" class="pi-status-integrations pi-status-clickable" data-tooltip="${ACTIVE_INTEGRATIONS_TOOLTIP_PREFIX}: ${escapeAttr(activeIntegrations.join(", "))}. Click to manage.">${integrationsSvg}<span>${formatIntegrationCountLabel(activeIntegrations.length)}</span><span class="pi-status-affordance" aria-hidden="true">${affordanceChevronSvg}</span></button>`
    : "";

  const thinkingTooltip = escapeAttr(
    "How deeply Pi reasons before answering — higher is slower but more thorough. Click to choose, or ⇧Tab to cycle.",
  );

  el.innerHTML = `
    <button type="button" class="pi-status-ctx pi-status-ctx--trigger pi-status-clickable has-tooltip" data-status-popover="${ctxPopoverText}" aria-label="Context usage ${pct}% of ${ctxLabel}"><span class="${ctxColor}">${pct}%</span> / ${ctxLabel}${usageDebug}<span class="pi-status-affordance" aria-hidden="true">${affordanceChevronSvg}</span><span class="pi-tooltip pi-tooltip--left">${ctxBaseTooltip}${ctxWarning}</span></button>
    ${lockBadge}
    ${rulesBadge}
    ${integrationsBadge}
    <button type="button" class="pi-status-model pi-status-clickable" data-tooltip="Switch the AI model powering this session">
      <span class="pi-status-model__mark">π</span>
      <span class="pi-status-model__name">${modelAliasEscaped}</span>
      ${chevronSvg}
    </button>
    <button type="button" class="pi-status-thinking pi-status-clickable" data-tooltip="${thinkingTooltip}" aria-label="Thinking level ${thinkingLevel}">${brainSvg} ${thinkingLevel}<span class="pi-status-affordance" aria-hidden="true">${affordanceChevronSvg}</span></button>
  `;
}

export function updateStatusBarForAgent(
  agent: Agent,
  lockState: RuntimeLockState = "idle",
  rulesActive = false,
  activeIntegrations: string[] = [],
): void {
  renderStatusBar(agent, lockState, rulesActive, activeIntegrations);
}

export function updateStatusBar(
  getActiveAgent: ActiveAgentProvider,
  getLockState?: ActiveLockStateProvider,
  getRulesActive?: ActiveRulesProvider,
  getActiveIntegrations?: ActiveIntegrationsProvider,
): void {
  const activeAgent = getActiveAgent();
  const lockState = getLockState ? getLockState() : "idle";
  const rulesActive = getRulesActive ? getRulesActive() : false;
  const activeIntegrations = getActiveIntegrations ? getActiveIntegrations() : [];
  renderStatusBar(activeAgent, lockState, rulesActive, activeIntegrations);
}

export function injectStatusBar(opts: {
  getActiveAgent: ActiveAgentProvider;
  getLockState?: ActiveLockStateProvider;
  getRulesActive?: ActiveRulesProvider;
  getActiveIntegrations?: ActiveIntegrationsProvider;
}): () => void {
  const { getActiveAgent, getLockState, getRulesActive, getActiveIntegrations } = opts;

  let unsubscribeActiveAgent: (() => void) | undefined;

  const bindActiveAgent = () => {
    unsubscribeActiveAgent?.();

    const activeAgent = getActiveAgent();
    if (activeAgent) {
      unsubscribeActiveAgent = activeAgent.subscribe(
        () => updateStatusBar(getActiveAgent, getLockState, getRulesActive, getActiveIntegrations),
      );
    } else {
      unsubscribeActiveAgent = undefined;
    }

    updateStatusBar(getActiveAgent, getLockState, getRulesActive, getActiveIntegrations);
  };

  const onStatusUpdate = () => updateStatusBar(getActiveAgent, getLockState, getRulesActive, getActiveIntegrations);

  document.addEventListener("pi:status-update", onStatusUpdate);
  document.addEventListener("pi:active-runtime-changed", bindActiveAgent);

  requestAnimationFrame(bindActiveAgent);

  return () => {
    unsubscribeActiveAgent?.();
    document.removeEventListener("pi:status-update", onStatusUpdate);
    document.removeEventListener("pi:active-runtime-changed", bindActiveAgent);
  };
}

export function flashThinkingLevel(level: string, color: string): void {
  const labels: Record<string, string> = {
    off: "Off",
    minimal: "Min",
    low: "Low",
    medium: "Medium",
    high: "High",
    xhigh: "Max",
  };
  showToast(`Thinking: ${labels[level] || level} (next turn)`, 1500);

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
