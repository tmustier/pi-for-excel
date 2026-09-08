function getOwnerPrefix(ownerId: string): string {
  return `${ownerId.trim().toLowerCase()}.`;
}

export function qualifyExtensionConnectionId(ownerId: string, connectionId: string): string {
  const normalized = connectionId.trim().toLowerCase();
  if (normalized.length === 0) {
    throw new Error("Connection id cannot be empty.");
  }

  const ownerPrefix = getOwnerPrefix(ownerId);
  return normalized.startsWith(ownerPrefix) ? normalized : `${ownerPrefix}${normalized}`;
}

export function qualifyExtensionProviderId(ownerId: string, providerId: string): string {
  const normalized = providerId.trim().toLowerCase();
  if (normalized.length === 0) {
    throw new Error("Provider id cannot be empty.");
  }
  if (!/^[a-z0-9][a-z0-9._-]*$/u.test(normalized)) {
    throw new Error("Provider id may contain only letters, numbers, dots, underscores and hyphens.");
  }

  const ownerPrefix = getOwnerPrefix(ownerId);
  return normalized.startsWith(ownerPrefix) ? normalized : `${ownerPrefix}${normalized}`;
}

export function isExtensionOwnedId(ownerId: string, qualifiedId: string): boolean {
  return qualifiedId.startsWith(getOwnerPrefix(ownerId));
}
