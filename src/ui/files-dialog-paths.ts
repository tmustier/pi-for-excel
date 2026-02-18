/**
 * Path helpers for Files dialog actions.
 */

interface PathParts {
  directory: string;
  baseName: string;
}

function splitPathParts(path: string): PathParts {
  const normalized = path.replaceAll("\\", "/");
  const lastSlash = normalized.lastIndexOf("/");

  if (lastSlash < 0) {
    return {
      directory: "",
      baseName: normalized,
    };
  }

  return {
    directory: normalized.slice(0, lastSlash),
    baseName: normalized.slice(lastSlash + 1),
  };
}

function getFileExtensionWithDot(fileName: string): string | null {
  const trimmed = fileName.trim();
  const lastDot = trimmed.lastIndexOf(".");

  if (lastDot <= 0 || lastDot >= trimmed.length - 1) {
    return null;
  }

  return trimmed.slice(lastDot);
}

function resolvePreservedExtension(args: {
  currentBaseName: string;
  targetBaseName: string;
}): string | null {
  const currentExtension = getFileExtensionWithDot(args.currentBaseName);
  if (!currentExtension) {
    return null;
  }

  const targetExtension = getFileExtensionWithDot(args.targetBaseName);
  if (targetExtension) {
    return null;
  }

  if (args.targetBaseName.startsWith(".") || args.targetBaseName.endsWith(".")) {
    return null;
  }

  return currentExtension;
}

/**
 * Resolves a rename destination path from user input.
 *
 * Rules:
 * - Empty input keeps the current path.
 * - Name-only input stays in the current folder.
 * - If the target basename omits an extension, preserve the current extension.
 */
export function resolveRenameDestinationPath(currentPath: string, inputPath: string): string {
  const normalizedInput = inputPath.trim().replaceAll("\\", "/");
  if (normalizedInput.length === 0) {
    return currentPath;
  }

  const currentParts = splitPathParts(currentPath);
  const targetParts = normalizedInput.includes("/")
    ? splitPathParts(normalizedInput)
    : {
      directory: currentParts.directory,
      baseName: normalizedInput,
    };

  const targetBaseName = targetParts.baseName.trim();
  if (targetBaseName.length === 0) {
    return currentPath;
  }

  const preservedExtension = resolvePreservedExtension({
    currentBaseName: currentParts.baseName,
    targetBaseName,
  });

  const resolvedBaseName = preservedExtension
    ? `${targetBaseName}${preservedExtension}`
    : targetBaseName;

  if (targetParts.directory.length === 0) {
    return resolvedBaseName;
  }

  return `${targetParts.directory}/${resolvedBaseName}`;
}
