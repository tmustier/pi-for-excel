import test from "node:test";
import assert from "node:assert/strict";

import {
  isAllowedClientAddress,
  isLoopbackAddress,
  parseClientCidrAllowlist,
} from "../scripts/proxy-client-policy.mjs";

test("isLoopbackAddress accepts IPv4/IPv6 loopback forms", () => {
  assert.equal(isLoopbackAddress("127.0.0.1"), true);
  assert.equal(isLoopbackAddress("127.1.2.3"), true);
  assert.equal(isLoopbackAddress("::1"), true);
  assert.equal(isLoopbackAddress("0:0:0:0:0:0:0:1"), true);
  assert.equal(isLoopbackAddress("::ffff:127.0.0.1"), true);
});

test("isLoopbackAddress rejects non-loopback and empty values", () => {
  assert.equal(isLoopbackAddress("10.0.0.1"), false);
  assert.equal(isLoopbackAddress("192.168.1.1"), false);
  assert.equal(isLoopbackAddress(""), false);
  assert.equal(isLoopbackAddress(undefined), false);
});

test("parseClientCidrAllowlist parses CIDRs and bare IPs", () => {
  const { cidrs, invalid } = parseClientCidrAllowlist("10.96.0.0/13, 192.168.1.5");
  assert.equal(invalid.length, 0);
  assert.equal(cidrs.length, 2);
  assert.equal(cidrs[0].bits, 13);
  assert.equal(cidrs[1].bits, 32);
});

test("parseClientCidrAllowlist rejects invalid entries", () => {
  const cases = [
    "10.96.0.0/0", // allow-everyone must be explicit at the network layer
    "10.96.0.0/33",
    "10.96.0.0/x",
    "999.1.1.1/8",
    "10.0.0/8",
    "not-an-ip",
    "fe80::1/64", // IPv6 ranges unsupported
    "10.0.0.0/8/8",
  ];
  for (const entry of cases) {
    const { cidrs, invalid } = parseClientCidrAllowlist(entry);
    assert.deepEqual(invalid, [entry], `expected invalid: ${entry}`);
    assert.equal(cidrs.length, 0, `expected no cidrs for: ${entry}`);
  }
});

test("parseClientCidrAllowlist handles empty input", () => {
  assert.deepEqual(parseClientCidrAllowlist(""), { cidrs: [], invalid: [] });
  assert.deepEqual(parseClientCidrAllowlist(undefined), { cidrs: [], invalid: [] });
  assert.deepEqual(parseClientCidrAllowlist(" , , "), { cidrs: [], invalid: [] });
});

test("isAllowedClientAddress always allows loopback", () => {
  assert.equal(isAllowedClientAddress("127.0.0.1", []), true);
  assert.equal(isAllowedClientAddress("::1", []), true);
  const { cidrs } = parseClientCidrAllowlist("10.96.0.0/13");
  assert.equal(isAllowedClientAddress("::ffff:127.0.0.1", cidrs), true);
});

test("isAllowedClientAddress matches configured IPv4 ranges", () => {
  const { cidrs } = parseClientCidrAllowlist("10.96.0.0/13,192.168.1.5");

  // Inside 10.96.0.0/13 (10.96.0.0 - 10.103.255.255)
  assert.equal(isAllowedClientAddress("10.96.0.1", cidrs), true);
  assert.equal(isAllowedClientAddress("10.103.255.254", cidrs), true);
  assert.equal(isAllowedClientAddress("::ffff:10.97.193.77", cidrs), true);

  // Outside
  assert.equal(isAllowedClientAddress("10.104.0.1", cidrs), false);
  assert.equal(isAllowedClientAddress("10.95.255.255", cidrs), false);
  assert.equal(isAllowedClientAddress("11.96.0.1", cidrs), false);

  // Bare IP behaves as /32
  assert.equal(isAllowedClientAddress("192.168.1.5", cidrs), true);
  assert.equal(isAllowedClientAddress("192.168.1.6", cidrs), false);
});

test("isAllowedClientAddress rejects non-IPv4 clients when ranges are configured", () => {
  const { cidrs } = parseClientCidrAllowlist("10.0.0.0/8");
  assert.equal(isAllowedClientAddress("fe80::1", cidrs), false);
  assert.equal(isAllowedClientAddress("2001:db8::2", cidrs), false);
  assert.equal(isAllowedClientAddress(undefined, cidrs), false);
  assert.equal(isAllowedClientAddress("", cidrs), false);
});

test("isAllowedClientAddress denies everything non-loopback with no ranges", () => {
  assert.equal(isAllowedClientAddress("10.0.0.1", []), false);
  assert.equal(isAllowedClientAddress("8.8.8.8", []), false);
});

test("isLoopbackAddress strictly parses address forms", () => {
  // Malformed / decorated forms must not pass
  assert.equal(isLoopbackAddress("127.evil"), false);
  assert.equal(isLoopbackAddress("127.0.0.1:1234"), false);
  assert.equal(isLoopbackAddress("::ffff:127.0.0.1%lo0"), false);
  assert.equal(isLoopbackAddress("127.0.0"), false);
  assert.equal(isLoopbackAddress("0177.0.0.1"), false); // octal-looking form: octet > 255
  assert.equal(isLoopbackAddress("2130706433"), false); // single-integer IPv4 form

  // Case-insensitive mapped form is fine
  assert.equal(isLoopbackAddress("::FFFF:127.0.0.1"), true);
});

test("CIDR matching edge bits: /1 and /32", () => {
  const one = parseClientCidrAllowlist("128.0.0.0/1").cidrs;
  assert.equal(isAllowedClientAddress("128.0.0.1", one), true);
  assert.equal(isAllowedClientAddress("255.255.255.255", one), true);
  assert.equal(isAllowedClientAddress("100.0.0.1", one), false);

  const exact = parseClientCidrAllowlist("10.1.2.3/32").cidrs;
  assert.equal(isAllowedClientAddress("10.1.2.3", exact), true);
  assert.equal(isAllowedClientAddress("10.1.2.4", exact), false);
});
