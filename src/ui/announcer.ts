/**
 * Screen-reader announcements for events that have no visible notification of
 * their own: a run finishing or failing (in this or another tab), a retry, an
 * approval waiting. See src/ui/README.md, "Notifications".
 *
 * Toasts and error banners already announce themselves; do not announce the
 * same event twice.
 */

type Politeness = "polite" | "assertive";

const regions = new Map<Politeness, HTMLElement>();
const timers = new Map<Politeness, ReturnType<typeof setTimeout>>();

function ensureRegion(politeness: Politeness): HTMLElement {
  const existing = regions.get(politeness);
  if (existing && existing.isConnected) return existing;

  const region = document.createElement("div");
  region.className = "pi-announcer";
  region.dataset.politeness = politeness;
  region.setAttribute("aria-live", politeness);
  region.setAttribute("aria-atomic", "true");
  document.body.appendChild(region);
  regions.set(politeness, region);
  return region;
}

/**
 * Announce `message`. Assertive interrupts the user (use only when they must
 * act, e.g. an approval is waiting); polite waits for a pause.
 * The returned cleanup withdraws this announcement without clearing a newer one.
 */
export function announce(message: string, opts: { assertive?: boolean } = {}): () => void {
  if (typeof document === "undefined") return () => {};
  const politeness: Politeness = opts.assertive ? "assertive" : "polite";
  const region = ensureRegion(politeness);

  // Clear first so the same text announced twice in a row is read again.
  region.textContent = "";
  const pending = timers.get(politeness);
  if (pending !== undefined) clearTimeout(pending);
  const timer = setTimeout(() => {
    region.textContent = message;
  }, 60);
  // Keep the handle after delivery: it identifies the announcement this cleanup owns.
  timers.set(politeness, timer);
  return () => {
    clearTimeout(timer);
    if (timers.get(politeness) !== timer) return;
    timers.delete(politeness);
    region.textContent = "";
  };
}
