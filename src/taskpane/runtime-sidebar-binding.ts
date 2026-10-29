import type { PiSidebar } from "../ui/pi-sidebar.js";
import type { SessionRuntimeManager, SessionRuntime } from "./session-runtime-manager.js";
import { Type } from "typebox";
import { Value } from "typebox/value";

const removeAttachmentSchema = Type.Object({ path: Type.String() });

export function bindRuntimeSidebar(opts: {
  runtimeManager: SessionRuntimeManager;
  sidebar: PiSidebar;
  renderError: (message: string | null) => void;
}): void {
  const { runtimeManager, sidebar, renderError } = opts;
  let renderedRuntimeId: string | null = null;
  let unsubscribeTab: (() => void) | null = null;

  const showTabState = (runtime: SessionRuntime) => {
    const input = sidebar.getInput();
    if (input) input.attachments = runtime.tab.attachments;
    if (input && input.value !== runtime.tab.draft) {
      input.value = runtime.tab.draft;
    }
    const activity = runtime.tab.activity;
    sidebar.setBusyIndicator(activity?.label ?? null, activity?.hint ?? null);
    sidebar.tabBusy = activity !== null || runtime.tab.busy;
    renderError(runtime.tab.error);
  };

  sidebar.addEventListener("pi-attachment-remove", (event) => {
    if (!(event instanceof CustomEvent)) return;
    if (!Value.Check(removeAttachmentSchema, event.detail)) return;
    const path = event.detail.path;
    const tab = runtimeManager.getActiveRuntime()?.tab;
    if (tab) tab.setAttachments(tab.attachments.filter((entry) => entry !== path));
  });
  sidebar.addEventListener("input", (event) => {
    if (!(event.target instanceof HTMLElement) || event.target.tagName !== "PI-INPUT") return;
    const runtime = runtimeManager.getActiveRuntime();
    const input = sidebar.getInput();
    if (!runtime || !input || runtime.runtimeId !== renderedRuntimeId) return;
    runtime.tab.setDraft(input.value);
  });

  runtimeManager.subscribe((snapshot) => {
    sidebar.sessionTabs = snapshot.tabs;
    const activeTab = snapshot.tabs.find((tab) => tab.isActive);
    sidebar.tabBusy = activeTab?.isBusy ?? false;

    if (snapshot.activeRuntimeId === renderedRuntimeId) {
      sidebar.requestUpdate();
      return;
    }

    const previousRuntime = renderedRuntimeId
      ? runtimeManager.getRuntime(renderedRuntimeId)
      : null;
    previousRuntime?.queueDisplay.detach();
    unsubscribeTab?.();
    unsubscribeTab = null;

    renderedRuntimeId = snapshot.activeRuntimeId;
    const activeRuntime = runtimeManager.getActiveRuntime();
    if (activeRuntime) {
      sidebar.agent = activeRuntime.agent;
      sidebar.syncFromAgent();
      sidebar.getInput()?.loadDraft(activeRuntime.tab.draft);
      showTabState(activeRuntime);
      unsubscribeTab = activeRuntime.tab.subscribe(() => showTabState(activeRuntime));
    } else {
      renderError(null);
    }
    sidebar.requestUpdate();

    if (!activeRuntime) return;

    const activeRuntimeId = activeRuntime.runtimeId;
    requestAnimationFrame(() => {
      const activeNow = runtimeManager.getActiveRuntime();
      if (!activeNow || activeNow.runtimeId !== activeRuntimeId) return;
      activeNow.queueDisplay.attach(sidebar);
      showTabState(activeNow);
    });
  });
}
