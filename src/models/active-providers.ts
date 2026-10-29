let _activeProviders: Set<string> | null = null;

export function setActiveProviders(providers: Set<string>): void {
  _activeProviders = providers;
}

export function getActiveProviders(): ReadonlySet<string> | null {
  return _activeProviders;
}
