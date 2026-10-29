import "./ui/theme.css";

import { installBedrockProviderStub } from "./compat/bedrock-provider-stub.js";
import { installCryptoRandomUuidPatch } from "./compat/crypto-random-uuid.js";
import { installMarkedSafetyPatch } from "./compat/marked-safety.js";
import { installThemeModeSync } from "./ui/theme-mode.js";

installBedrockProviderStub();
installCryptoRandomUuidPatch();
installMarkedSafetyPatch();
installThemeModeSync();
