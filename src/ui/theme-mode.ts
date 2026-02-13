/**
 * Dark/light mode synchronization.
 *
 * Prefers the Office theme background when available, with a media-query fallback.
 */

interface RgbColor {
  r: number;
  g: number;
  b: number;
}

function parseHexColor(input: string): RgbColor | null {
  const raw = input.trim();
  const normalized = raw.startsWith("#") ? raw.slice(1) : raw;

  if (normalized.length === 3) {
    const r = Number.parseInt(normalized[0].repeat(2), 16);
    const g = Number.parseInt(normalized[1].repeat(2), 16);
    const b = Number.parseInt(normalized[2].repeat(2), 16);
    if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) {
      return null;
    }

    return { r, g, b };
  }

  if (normalized.length !== 6) {
    return null;
  }

  const r = Number.parseInt(normalized.slice(0, 2), 16);
  const g = Number.parseInt(normalized.slice(2, 4), 16);
  const b = Number.parseInt(normalized.slice(4, 6), 16);

  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) {
    return null;
  }

  return { r, g, b };
}

function toLinearSrgb(channel: number): number {
  const normalized = channel / 255;
  if (normalized <= 0.04045) {
    return normalized / 12.92;
  }

  return ((normalized + 0.055) / 1.055) ** 2.4;
}

function relativeLuminance(rgb: RgbColor): number {
  return (
    0.2126 * toLinearSrgb(rgb.r)
    + 0.7152 * toLinearSrgb(rgb.g)
    + 0.0722 * toLinearSrgb(rgb.b)
  );
}

function isDarkColor(rgb: RgbColor): boolean {
  return relativeLuminance(rgb) < 0.35;
}

function resolveOfficeThemeDark(): boolean | null {
  if (typeof Office === "undefined") {
    return null;
  }

  const officeTheme = Office.context?.officeTheme;
  const backgroundColor = officeTheme?.bodyBackgroundColor;

  if (typeof backgroundColor !== "string") {
    return null;
  }

  const parsed = parseHexColor(backgroundColor);
  if (!parsed) {
    return null;
  }

  return isDarkColor(parsed);
}

function resolvePreferredDark(mediaMatches: boolean): boolean {
  const officeDark = resolveOfficeThemeDark();
  if (officeDark !== null) {
    return officeDark;
  }

  return mediaMatches;
}

export function installThemeModeSync(): () => void {
  if (typeof document === "undefined" || typeof window === "undefined") {
    return () => {};
  }

  const root = document.documentElement;
  const media = window.matchMedia("(prefers-color-scheme: dark)");

  const apply = () => {
    root.classList.toggle("dark", resolvePreferredDark(media.matches));
  };

  apply();

  let disposed = false;
  let officeReadyHooked = false;
  let officeRetryTimer: number | null = null;
  let officeRetryStopTimer: number | null = null;

  const registerOfficeReadyHook = (): void => {
    if (officeReadyHooked || typeof Office === "undefined") {
      return;
    }

    officeReadyHooked = true;
    void Office.onReady(() => {
      if (disposed) return;
      apply();
    });
  };

  registerOfficeReadyHook();

  if (!officeReadyHooked) {
    officeRetryTimer = window.setInterval(() => {
      registerOfficeReadyHook();
      if (officeReadyHooked && officeRetryTimer !== null) {
        clearInterval(officeRetryTimer);
        officeRetryTimer = null;
      }
    }, 500);

    officeRetryStopTimer = window.setTimeout(() => {
      if (officeRetryTimer !== null) {
        clearInterval(officeRetryTimer);
        officeRetryTimer = null;
      }
      officeRetryStopTimer = null;
    }, 15_000);
  }

  const onMediaChange = () => {
    apply();
  };

  if (typeof media.addEventListener === "function") {
    media.addEventListener("change", onMediaChange);

    return () => {
      disposed = true;
      media.removeEventListener("change", onMediaChange);
      if (officeRetryTimer !== null) {
        clearInterval(officeRetryTimer);
      }
      if (officeRetryStopTimer !== null) {
        clearTimeout(officeRetryStopTimer);
      }
    };
  }

  media.addListener(onMediaChange);
  return () => {
    disposed = true;
    media.removeListener(onMediaChange);
    if (officeRetryTimer !== null) {
      clearInterval(officeRetryTimer);
    }
    if (officeRetryStopTimer !== null) {
      clearTimeout(officeRetryStopTimer);
    }
  };
}
