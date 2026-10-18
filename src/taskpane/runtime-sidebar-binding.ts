/**
 * Binds UI-independent runtime lifecycle snapshots to the visible sidebar.
 */

import type { PiSidebar } from "../ui/pi-sidebar.js";
import type { SessionRuntimeManager } from "./session-runtime-manager.js";

export function bindRuntimeSidebar(opts: {
  runtimeManager: SessionRuntimeManager;
  sidebar: PiSidebar;
}): void {
  const { runtimeManager, sidebar } = opts;
  let renderedRuntimeId: string | null = null;

  runtimeManager.subscribe((snapshot) => {
    sidebar.sessionTabs = snapshot.tabs;

    if (snapshot.activeRuntimeId === renderedRuntimeId) {
      const activeRuntime = runtimeManager.getActiveRuntime();
      if (activeRuntime) sidebar.tools = activeRuntime.tools;
      sidebar.requestUpdate();
      return;
    }

    const previousRuntime = renderedRuntimeId
      ? runtimeManager.getRuntime(renderedRuntimeId)
      : null;
    previousRuntime?.queueDisplay.detach();

    renderedRuntimeId = snapshot.activeRuntimeId;
    const activeRuntime = runtimeManager.getActiveRuntime();
    if (activeRuntime) {
      sidebar.view = activeRuntime.view;
      sidebar.tools = activeRuntime.tools;
      sidebar.syncFromConversation();
    }
    sidebar.requestUpdate();

    if (!activeRuntime) return;

    const activeRuntimeId = activeRuntime.runtimeId;
    requestAnimationFrame(() => {
      const activeNow = runtimeManager.getActiveRuntime();
      if (!activeNow || activeNow.runtimeId !== activeRuntimeId) return;
      activeNow.queueDisplay.attach(sidebar);
    });
  });
}
