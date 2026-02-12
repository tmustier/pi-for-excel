/**
 * Safe path helpers for workspace file operations.
 *
 * Paths are always relative to the workspace root.
 */

export function normalizeWorkspacePath(rawPath: string): string {
  const trimmed = rawPath.trim();
  if (trimmed.length === 0) {
    throw new Error("Path is required.");
  }

  const normalizedSeparators = trimmed.replaceAll("\\", "/");
  if (normalizedSeparators.startsWith("/")) {
    throw new Error("Path must be relative to the workspace root.");
  }

  const parts = normalizedSeparators
    .split("/")
    .map((part) => part.trim())
    .filter((part) => part.length > 0);

  if (parts.length === 0) {
    throw new Error("Path is required.");
  }

  for (const part of parts) {
    if (part === "." || part === "..") {
      throw new Error("Path cannot contain '.' or '..' segments.");
    }

    if (part.includes("\0")) {
      throw new Error("Path contains invalid characters.");
    }
  }

  return parts.join("/");
}

export function splitWorkspacePath(path: string): string[] {
  return normalizeWorkspacePath(path).split("/");
}

export function getWorkspaceBaseName(path: string): string {
  const parts = splitWorkspacePath(path);
  return parts[parts.length - 1] ?? path;
}
