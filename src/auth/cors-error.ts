/**
 * Vendored from @earendil-works/pi-web-ui 0.75.3 utils/proxy-utils
 * (MIT, © Mario Zechner, https://github.com/badlogic/pi-mono).
 */

export function isModuleLoadError(error: Error): boolean {
  return /dynamically imported module|importing a module script failed|error loading dynamically imported/i.test(error.message);
}

export function isCorsError(error: unknown): boolean {
  if (!(error instanceof Error)) {
    return false;
  }

  const message = error.message.toLowerCase();

  if (isModuleLoadError(error)) {
    return false;
  }

  if (error.name === "TypeError" && message.includes("failed to fetch")) {
    return true;
  }

  if (error.name === "NetworkError") {
    return true;
  }

  if (message.includes("cors") || message.includes("cross-origin")) {
    return true;
  }

  return false;
}
