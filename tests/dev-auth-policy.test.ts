import assert from "node:assert/strict";
import { test } from "node:test";

import {
  hostnameFromHostHeader,
  isLocalPiAuthHost,
  isPiAuthRequestAllowed,
  isLoopbackAddress,
} from "../src/dev-auth-policy.ts";

void test("isLoopbackAddress accepts IPv4 and IPv6 loopback forms", () => {
  assert.equal(isLoopbackAddress("127.0.0.1"), true);
  assert.equal(isLoopbackAddress("127.12.34.56"), true);
  assert.equal(isLoopbackAddress("::1"), true);
  assert.equal(isLoopbackAddress("0:0:0:0:0:0:0:1"), true);
  assert.equal(isLoopbackAddress("::ffff:127.0.0.1"), true);
  assert.equal(isLoopbackAddress("10.0.2.15"), false);
  assert.equal(isLoopbackAddress(undefined), false);
});

void test("hostnameFromHostHeader handles ports and bracketed IPv6", () => {
  assert.equal(hostnameFromHostHeader("localhost:3141"), "localhost");
  assert.equal(hostnameFromHostHeader("127.0.0.1:3141"), "127.0.0.1");
  assert.equal(hostnameFromHostHeader("[::1]:3141"), "::1");
  assert.equal(hostnameFromHostHeader("10.0.2.2:3141"), "10.0.2.2");
  assert.equal(hostnameFromHostHeader(undefined), null);
});

void test("isLocalPiAuthHost allows only loopback hostnames", () => {
  assert.equal(isLocalPiAuthHost("localhost:3141"), true);
  assert.equal(isLocalPiAuthHost("127.0.0.1:3141"), true);
  assert.equal(isLocalPiAuthHost("[::1]:3141"), true);
  assert.equal(isLocalPiAuthHost("10.0.2.2:3141"), false);
  assert.equal(isLocalPiAuthHost("example.com:3141"), false);
});

void test("isPiAuthRequestAllowed requires both loopback socket and local host header", () => {
  assert.equal(isPiAuthRequestAllowed({ remoteAddress: "127.0.0.1", hostHeader: "localhost:3141" }), true);
  assert.equal(isPiAuthRequestAllowed({ remoteAddress: "::1", hostHeader: "[::1]:3141" }), true);
  assert.equal(isPiAuthRequestAllowed({ remoteAddress: "10.0.2.15", hostHeader: "localhost:3141" }), false);
  assert.equal(isPiAuthRequestAllowed({ remoteAddress: "127.0.0.1", hostHeader: "10.0.2.2:3141" }), false);
});

void test("isPiAuthRequestAllowed has an explicit non-local-host opt-in for disposable dev runs", () => {
  assert.equal(isPiAuthRequestAllowed({ remoteAddress: "127.0.0.1", hostHeader: "10.0.2.2:3141", allowNonLocalHost: true }), true);
  assert.equal(isPiAuthRequestAllowed({ remoteAddress: "10.0.2.15", hostHeader: "10.0.2.2:3141", allowNonLocalHost: true }), false);
});
