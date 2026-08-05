import type { StreamFn } from "@earendil-works/pi-agent-core";

/** Fail fast if a test that only needs Agent state unexpectedly starts a provider request. */
export const unusedTestStreamFn: StreamFn = () => {
  throw new Error("Test agent stream function was called unexpectedly.");
};
