/**
 * Status bar rendering + thinking level flash.
 */

import type { Agent } from "@mariozechner/pi-agent-core";

import { showToast } from "../ui/toast.js";
import { escapeAttr, escapeHtml } from "../utils/html.js";
import { formatUsageDebug, isDebugEnabled } from "../debug/debug.js";
import { estimateContextTokens } from "../utils/context-tokens.js";
import type { RuntimeLockState } from "./session-runtime-manager.js";

export type ActiveAgentProvider = () => Agent | null;
export type ActiveLockStateProvider = () => RuntimeLockState;
export type ActiveInstructionsProvider = () => boolean;
export type ActiveSkillsProvider = () => string[];

function renderStatusBar(
  agent: Agent | null,
  lockState: RuntimeLockState,
  instructionsActive: boolean,
  activeSkills: string[],
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
  const ctxBaseTooltip = `How much of the model's context window has been used (${totalTokens.toLocaleString()} / ${contextWindow.toLocaleString()} tokens). As it fills up the model may lose track of earlier details — start a new chat if quality drops.`;
  let ctxWarning = "";
  let ctxWarningText = "";
  if (pct > 100) {
    ctxColor = "pi-status-ctx--red";
    ctxWarningText = "Context window exceeded — the next message will fail. Use /compact to free up some context, or /new to clear the chat.";
    ctxWarning = `<span class="pi-tooltip__warn pi-tooltip__warn--red">${ctxWarningText}</span>`;
  } else if (pct > 60) {
    ctxColor = "pi-status-ctx--red";
    ctxWarningText = `Context ${pct}% used up — quality will degrade. Use /compact to free up some context, or /new to clear the chat.`;
    ctxWarning = `<span class="pi-tooltip__warn pi-tooltip__warn--red">${ctxWarningText}</span>`;
  } else if (pct > 40) {
    ctxColor = "pi-status-ctx--yellow";
    ctxWarningText = `Context ${pct}% used up. Consider using /compact to free up some context, or /new to clear the chat.`;
    ctxWarning = `<span class="pi-tooltip__warn pi-tooltip__warn--yellow">${ctxWarningText}</span>`;
  }

  const ctxPopoverText = escapeAttr(
    ctxWarningText.length > 0 ? `${ctxBaseTooltip} ${ctxWarningText}` : ctxBaseTooltip,
  );

  const chevronSvg = `<svg width="8" height="8" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>`;
  const brainSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 18V5"/><path d="M15 13a4.17 4.17 0 0 1-3-4 4.17 4.17 0 0 1-3 4"/><path d="M17.598 6.5A3 3 0 1 0 12 5a3 3 0 1 0-5.598 1.5"/><path d="M17.997 5.125a4 4 0 0 1 2.526 5.77"/><path d="M18 18a4 4 0 0 0 2-7.464"/><path d="M19.967 17.483A4 4 0 1 1 12 18a4 4 0 1 1-7.967-.517"/><path d="M6 18a4 4 0 0 1-2-7.464"/><path d="M6.003 5.125a4 4 0 0 0-2.526 5.77"/></svg>`;

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

  const instructionsBadge = instructionsActive
    ? `<button class="pi-status-instructions" data-tooltip="Persistent instructions are active. Click to edit.">📋 instr</button>`
    : "";

  const skillsBadge = activeSkills.length > 0
    ? `<button class="pi-status-skills" data-tooltip="Active skills: ${escapeAttr(activeSkills.join(", "))}. Click to manage.">🧩 ${activeSkills.length} skill${activeSkills.length === 1 ? "" : "s"}</button>`
    : "";

  const thinkingTooltip = escapeAttr(
    "Controls how long the model \"thinks\" before answering — higher = slower but better reasoning. Click to choose a level, or use ⇧Tab to cycle.",
  );

  el.innerHTML = `
    <span class="pi-status-ctx pi-status-ctx--trigger has-tooltip" data-status-popover="${ctxPopoverText}"><span class="${ctxColor}">${pct}%</span> / ${ctxLabel}${usageDebug}<span class="pi-tooltip pi-tooltip--left">${ctxBaseTooltip}${ctxWarning}</span></span>
    ${lockBadge}
    ${instructionsBadge}
    ${skillsBadge}
    <button class="pi-status-model" data-tooltip="Switch the AI model powering this session">
      <span class="pi-status-model__mark">π</span>
      <span class="pi-status-model__name">${modelAliasEscaped}</span>
      ${chevronSvg}
    </button>
    <span class="pi-status-thinking" data-tooltip="${thinkingTooltip}">${brainSvg} ${thinkingLevel}</span>
  `;
}

export function updateStatusBarForAgent(
  agent: Agent,
  lockState: RuntimeLockState = "idle",
  instructionsActive = false,
  activeSkills: string[] = [],
): void {
  renderStatusBar(agent, lockState, instructionsActive, activeSkills);
}

export function updateStatusBar(
  getActiveAgent: ActiveAgentProvider,
  getLockState?: ActiveLockStateProvider,
  getInstructionsActive?: ActiveInstructionsProvider,
  getActiveSkills?: ActiveSkillsProvider,
): void {
  const activeAgent = getActiveAgent();
  const lockState = getLockState ? getLockState() : "idle";
  const instructionsActive = getInstructionsActive ? getInstructionsActive() : false;
  const activeSkills = getActiveSkills ? getActiveSkills() : [];
  renderStatusBar(activeAgent, lockState, instructionsActive, activeSkills);
}

export function injectStatusBar(opts: {
  getActiveAgent: ActiveAgentProvider;
  getLockState?: ActiveLockStateProvider;
  getInstructionsActive?: ActiveInstructionsProvider;
  getActiveSkills?: ActiveSkillsProvider;
}): () => void {
  const { getActiveAgent, getLockState, getInstructionsActive, getActiveSkills } = opts;

  let unsubscribeActiveAgent: (() => void) | undefined;

  const bindActiveAgent = () => {
    unsubscribeActiveAgent?.();

    const activeAgent = getActiveAgent();
    if (activeAgent) {
      unsubscribeActiveAgent = activeAgent.subscribe(
        () => updateStatusBar(getActiveAgent, getLockState, getInstructionsActive, getActiveSkills),
      );
    } else {
      unsubscribeActiveAgent = undefined;
    }

    updateStatusBar(getActiveAgent, getLockState, getInstructionsActive, getActiveSkills);
  };

  const onStatusUpdate = () => updateStatusBar(getActiveAgent, getLockState, getInstructionsActive, getActiveSkills);

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
