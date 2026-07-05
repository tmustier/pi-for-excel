/**
 * Session title helpers for tab labels and close/reopen toasts.
 */

import { t } from "../language/index.js";

export interface ResolveTabTitleArgs {
  hasExplicitTitle: boolean;
  sessionTitle: string;
  /** Stable, one-based number assigned when the tab/runtime is created. */
  defaultTabNumber: number;
}

function normalizeTabNumber(defaultTabNumber: number): number {
  if (!Number.isFinite(defaultTabNumber) || defaultTabNumber < 1) {
    return 1;
  }

  return Math.floor(defaultTabNumber);
}

export function resolveTabTitle(args: ResolveTabTitleArgs): string {
  if (args.hasExplicitTitle) {
    const trimmed = args.sessionTitle.trim();
    if (trimmed.length > 0) {
      return trimmed;
    }
  }

  return t("session.tab.default", { number: String(normalizeTabNumber(args.defaultTabNumber)) });
}
