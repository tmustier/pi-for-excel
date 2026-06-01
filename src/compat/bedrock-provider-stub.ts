/**
 * Browser-safe Amazon Bedrock provider shim.
 *
 * pi-ai 0.75 lazy-loads the Node-only Bedrock provider through a variable
 * dynamic import, which Vite cannot reliably rewrite to our local stub. Register
 * the stub explicitly so selecting a Bedrock model fails with a clear unsupported
 * message instead of trying to load AWS/Node transport code in the Excel WebView.
 */

import { setBedrockProviderModule } from "@earendil-works/pi-ai";

import * as bedrockProviderModule from "../stubs/amazon-bedrock.js";

let installed = false;

export function installBedrockProviderStub(): void {
  if (installed) return;
  installed = true;
  setBedrockProviderModule(bedrockProviderModule);
}
