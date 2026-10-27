import type { SettingsAccess } from "../storage/local/settings-store.js";

export const SUBSCRIPTION_WARNING_SETTING = "auth.anthropicSubscriptionWarning";

/** UI history only: never add this warning to Agent messages or prompts. */
export async function showSubscriptionWarning(
  settings: SettingsAccess,
  sessionId: string,
  provider: string,
  apiKey: string | undefined,
): Promise<boolean> {
  if (await settings.get(SUBSCRIPTION_WARNING_SETTING) === false) return false;
  const historyKey = `ui.subscriptionWarning.${sessionId}`;
  if (await settings.get(historyKey) === true) return true;
  if (provider !== "anthropic" || !apiKey?.startsWith("sk-ant-oat")) return false;
  await settings.set(historyKey, true);
  return true;
}
