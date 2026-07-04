/**
 * Shared relative date formatting for overlay lists.
 */

import { t } from "../../language/index.js";

export function formatRelativeDate(iso: string): string {
  const date = new Date(iso);
  const now = new Date();
  const diff = now.getTime() - date.getTime();

  if (diff < 60_000) return t("date.justNow");
  if (diff < 3_600_000) return t("date.minutesAgo", { n: Math.round(diff / 60_000) });
  if (diff < 86_400_000) return t("date.hoursAgo", { n: Math.round(diff / 3_600_000) });
  if (diff < 604_800_000) return t("date.daysAgo", { n: Math.round(diff / 86_400_000) });
  return date.toLocaleDateString();
}
