import assert from "node:assert/strict";
import { test } from "node:test";

import {
  ALL_EXTENSION_CAPABILITIES,
  getDefaultPermissionsForTrust,
  isExtensionCapabilityAllowed,
  listGrantedExtensionCapabilities,
  setExtensionCapabilityAllowed,
} from "../src/extensions/permissions.ts";

void test("setExtensionCapabilityAllowed toggles capabilities consistently", () => {
  let permissions = getDefaultPermissionsForTrust("inline-code");

  for (const capability of ALL_EXTENSION_CAPABILITIES) {
    permissions = setExtensionCapabilityAllowed(permissions, capability, true);
    assert.equal(isExtensionCapabilityAllowed(permissions, capability), true);

    permissions = setExtensionCapabilityAllowed(permissions, capability, false);
    assert.equal(isExtensionCapabilityAllowed(permissions, capability), false);
  }
});

void test("listGrantedExtensionCapabilities follows capability checks", () => {
  let permissions = getDefaultPermissionsForTrust("inline-code");

  permissions = setExtensionCapabilityAllowed(permissions, "commands.register", true);
  permissions = setExtensionCapabilityAllowed(permissions, "ui.widget", true);
  permissions = setExtensionCapabilityAllowed(permissions, "http.fetch", true);

  const listed = listGrantedExtensionCapabilities(permissions);

  const expected = ALL_EXTENSION_CAPABILITIES.filter((capability) => {
    return isExtensionCapabilityAllowed(permissions, capability);
  });

  assert.deepEqual(listed, expected);
});
