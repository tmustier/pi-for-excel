/**
 * Provider picker alias.
 *
 * Providers now live under Settings → Logins.
 */

export async function showProviderPicker(): Promise<void> {
  const { showSettingsDialog } = await import("./settings-overlay.js");
  await showSettingsDialog({ section: "providers" });
}
