import assert from "node:assert/strict";
import { test } from "node:test";

import { isPiAuthRequestAllowed, type PiAuthRequestPolicyInput } from "../src/dev-auth-policy.ts";

// Policy table for the dev credential endpoint. The HTTP wiring of this policy is
// covered by tests/dev-auth-server.test.ts against a real Vite server; this table
// pins the rows that HTTP/1.1 alone cannot reach (HTTP/2 `:authority`, opt-in).
const POLICY: ReadonlyArray<{ input: PiAuthRequestPolicyInput; allowed: boolean; why: string }> = [
  { input: { remoteAddress: "127.0.0.1", hostHeader: "localhost:3141" }, allowed: true, why: "loopback socket + local host" },
  { input: { remoteAddress: "::1", hostHeader: "[::1]:3141" }, allowed: true, why: "IPv6 loopback + bracketed host" },
  { input: { remoteAddress: "::ffff:127.0.0.1", hostHeader: "127.0.0.1:3141" }, allowed: true, why: "mapped IPv4 loopback" },
  { input: { remoteAddress: "10.0.2.15", hostHeader: "localhost:3141" }, allowed: false, why: "non-loopback socket" },
  { input: { remoteAddress: "127.0.0.1", hostHeader: "10.0.2.2:3141" }, allowed: false, why: "loopback socket but guest-facing host" },
  { input: { remoteAddress: "127.0.0.1", hostHeader: ["localhost:3141", "example.com"] }, allowed: true, why: "repeated host header: first value decides" },
  { input: { remoteAddress: "127.0.0.1" }, allowed: false, why: "no authority at all" },
  { input: { remoteAddress: undefined, hostHeader: "localhost:3141" }, allowed: false, why: "unknown socket address" },
  { input: { remoteAddress: "127.0.0.1", hostHeader: "10.0.2.2:3141", allowNonLocalHost: true }, allowed: true, why: "explicit opt-in relaxes host only" },
  { input: { remoteAddress: "10.0.2.15", hostHeader: "10.0.2.2:3141", allowNonLocalHost: true }, allowed: false, why: "opt-in never relaxes loopback" },
  { input: { remoteAddress: "::1", authorityHeader: "localhost:3141" }, allowed: true, why: "HTTP/2 authority alone" },
  { input: { remoteAddress: "::1", authorityHeader: "10.0.2.2:3141" }, allowed: false, why: "HTTP/2 non-local authority" },
  { input: { remoteAddress: "10.0.2.15", authorityHeader: "localhost:3141" }, allowed: false, why: "HTTP/2 needs loopback socket too" },
  { input: { remoteAddress: "::1", hostHeader: "example.com", authorityHeader: "localhost:3141" }, allowed: false, why: "host and authority must both be local" },
  { input: { remoteAddress: "::1", hostHeader: "localhost:3141", authorityHeader: "10.0.2.2:3141" }, allowed: false, why: "authority cannot be overridden by host" },
  { input: { remoteAddress: "::1", hostHeader: "localhost:3141", authorityHeader: "" }, allowed: false, why: "empty authority is not absent" },
  { input: { remoteAddress: "::1", hostHeader: "localhost:3141", authorityHeader: "localhost:3141" }, allowed: true, why: "both present and local" },
];

void test("dev auth request policy table", () => {
  const observed = POLICY.map((row) => ({ why: row.why, allowed: isPiAuthRequestAllowed(row.input) }));
  assert.deepEqual(observed, POLICY.map((row) => ({ why: row.why, allowed: row.allowed })));
});
