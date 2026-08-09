import type { StreamFn } from "@earendil-works/pi-agent-core";

export const failOnUnexpectedStream: StreamFn = () => {
  throw new Error("Unexpected test agent stream request.");
};
