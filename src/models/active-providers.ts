/**
 * Active provider set — providers the user has credentials for.
 *
 * Written by the taskpane provider-refresh flows; read by the model selector
 * to hide models the user cannot run. Moved from the retired
 * src/compat/model-selector-patch.ts.
 */

let _activeProviders: Set<string> | null = null;

export function setActiveProviders(providers: Set<string>): void {
  _activeProviders = providers;
}

export function getActiveProviders(): ReadonlySet<string> | null {
  return _activeProviders;
}
