/**
 * Provider picker alias.
 *
 * Providers now live under Settings → Model providers.
 */

export async function showProviderPicker(): Promise<void> {
  const { openSettings } = await import("./settings-pages/index.js");
  await openSettings("providers");
}
