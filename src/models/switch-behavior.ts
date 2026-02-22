/**
 * Model-switch behavior persistence.
 *
 * Controls what happens when the user changes model in a non-empty session:
 * - "inPlace": switch model in the current tab (pi-mono parity)
 * - "fork": clone into a new tab with the selected model
 */

export const MODEL_SWITCH_BEHAVIOR_SETTING_KEY = "model.switch.behavior.v1";

export type ModelSwitchBehavior = "inPlace" | "fork";

export interface ModelSwitchBehaviorStore {
  get: (key: string) => Promise<unknown>;
  set: (key: string, value: unknown) => Promise<void>;
}

export function isModelSwitchBehavior(value: unknown): value is ModelSwitchBehavior {
  return value === "inPlace" || value === "fork";
}

export function normalizeModelSwitchBehavior(value: unknown): ModelSwitchBehavior {
  return isModelSwitchBehavior(value) ? value : "inPlace";
}

export async function getStoredModelSwitchBehavior(
  store: ModelSwitchBehaviorStore,
): Promise<ModelSwitchBehavior> {
  const value = await store.get(MODEL_SWITCH_BEHAVIOR_SETTING_KEY);
  return normalizeModelSwitchBehavior(value);
}

export async function setStoredModelSwitchBehavior(
  store: ModelSwitchBehaviorStore,
  behavior: ModelSwitchBehavior,
): Promise<ModelSwitchBehavior> {
  await store.set(MODEL_SWITCH_BEHAVIOR_SETTING_KEY, behavior);
  return behavior;
}

export function shouldForkModelSwitch(args: {
  behavior: ModelSwitchBehavior;
  hasMessages: boolean;
}): boolean {
  return args.behavior === "fork" && args.hasMessages;
}
