import assert from "node:assert/strict";
import { test } from "node:test";

import {
  ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY,
  classifyExtensionSource,
  isRemoteExtensionOptIn,
} from "../src/commands/extension-source-policy.ts";

void test("classifyExtensionSource allows local specifiers and blob URLs by default", () => {
  assert.equal(classifyExtensionSource("./extensions/snake.js"), "local-module");
  assert.equal(classifyExtensionSource("../extensions/snake.js"), "local-module");
  assert.equal(classifyExtensionSource("/extensions/snake.js"), "local-module");

  assert.equal(classifyExtensionSource("https://example.com/ext.js"), "remote-url");
  assert.equal(classifyExtensionSource("HTTP://example.com/ext.js"), "remote-url");
  assert.equal(classifyExtensionSource("//attacker.example/ext.js"), "remote-url");

  assert.equal(classifyExtensionSource("blob:https://example.com/abc"), "blob-url");
  assert.equal(classifyExtensionSource("data:text/javascript,export%20default%20()=%3E{}"), "unsupported");
  assert.equal(classifyExtensionSource("my-extension"), "unsupported");
  assert.equal(classifyExtensionSource("   "), "unsupported");
});

void test("isRemoteExtensionOptIn only accepts explicit truthy values", () => {
  assert.equal(isRemoteExtensionOptIn("1"), true);
  assert.equal(isRemoteExtensionOptIn("true"), true);

  assert.equal(isRemoteExtensionOptIn("TRUE"), false);
  assert.equal(isRemoteExtensionOptIn("yes"), false);
  assert.equal(isRemoteExtensionOptIn("0"), false);
  assert.equal(isRemoteExtensionOptIn(""), false);
  assert.equal(isRemoteExtensionOptIn(null), false);
  assert.equal(isRemoteExtensionOptIn(undefined), false);

  assert.equal(ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY, "pi.allowRemoteExtensionUrls");
});
