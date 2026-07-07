/**
 * Boot — runs before any UI components mount.
 *
 * 1. Imports the first-party theme CSS (tokens, preflight, components)
 * 2. Installs markdown-safety and theme-mode patches
 * 3. Installs browser/runtime compatibility patches, including the Bedrock stub
 *
 * MUST be imported as the first module in taskpane.ts.
 */

import "./ui/theme.css";

import { installBedrockProviderStub } from "./compat/bedrock-provider-stub.js";
import { installCryptoRandomUuidPatch } from "./compat/crypto-random-uuid.js";
import { installMarkedSafetyPatch } from "./compat/marked-safety.js";
import { installThemeModeSync } from "./ui/theme-mode.js";

installBedrockProviderStub();
installCryptoRandomUuidPatch();
installMarkedSafetyPatch();
installThemeModeSync();
