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

export function announce(message: string, opts: { assertive?: boolean } = {}): () => void {
  if (typeof document === "undefined") return () => {};
  const politeness: Politeness = opts.assertive ? "assertive" : "polite";
  const region = ensureRegion(politeness);

  region.textContent = "";
  const pending = timers.get(politeness);
  if (pending !== undefined) clearTimeout(pending);
  const timer = setTimeout(() => {
    region.textContent = message;
  }, 60);
  timers.set(politeness, timer);
  return () => {
    clearTimeout(timer);
    if (timers.get(politeness) !== timer) return;
    timers.delete(politeness);
    region.textContent = "";
  };
}
