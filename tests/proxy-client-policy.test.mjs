import test from "node:test";
import assert from "node:assert/strict";

import {
  isAllowedClientAddress,
  isLoopbackAddress,
  parseClientCidrAllowlist,
} from "../scripts/proxy-client-policy.mjs";

test("client-address policy classifies loopback forms", () => {
  const cases = [
    ["127.0.0.1", true], ["127.1.2.3", true], ["::1", true],
    ["0:0:0:0:0:0:0:1", true], ["::ffff:127.0.0.1", true], ["::FFFF:127.0.0.1", true],
    ["10.0.0.1", false], ["192.168.1.1", false], ["", false], [undefined, false],
    ["127.evil", false], ["127.0.0.1:1234", false], ["::ffff:127.0.0.1%lo0", false],
    ["127.0.0", false], ["0177.0.0.1", false], ["2130706433", false],
  ];
  for (const [address, expected] of cases) {
    assert.equal(isLoopbackAddress(address), expected, String(address));
  }
});

test("client CIDR policy parses valid, invalid, and empty allowlists", () => {
  const valid = parseClientCidrAllowlist("10.96.0.0/13, 192.168.1.5");
  assert.deepEqual(valid.invalid, []);
  assert.deepEqual(valid.cidrs.map(({ bits }) => bits), [13, 32]);

  const invalidCases = ["10.96.0.0/0", "10.96.0.0/33", "10.96.0.0/x", "999.1.1.1/8", "10.0.0/8", "not-an-ip", "fe80::1/64", "10.0.0.0/8/8"];
  for (const entry of invalidCases) {
    assert.deepEqual(parseClientCidrAllowlist(entry), { cidrs: [], invalid: [entry] }, entry);
  }
  for (const empty of ["", undefined, " , , "]) {
    assert.deepEqual(parseClientCidrAllowlist(empty), { cidrs: [], invalid: [] }, String(empty));
  }
});

test("client-address policy applies loopback, CIDR, and default-deny decisions", () => {
  const ranges = parseClientCidrAllowlist("10.96.0.0/13,192.168.1.5").cidrs;
  const cases = [
    ["127.0.0.1", [], true], ["::1", [], true], ["::ffff:127.0.0.1", ranges, true],
    ["10.96.0.1", ranges, true], ["10.103.255.254", ranges, true], ["::ffff:10.97.193.77", ranges, true],
    ["10.104.0.1", ranges, false], ["10.95.255.255", ranges, false], ["11.96.0.1", ranges, false],
    ["192.168.1.5", ranges, true], ["192.168.1.6", ranges, false],
    ["fe80::1", ranges, false], ["2001:db8::2", ranges, false], [undefined, ranges, false], ["", ranges, false],
    ["10.0.0.1", [], false], ["8.8.8.8", [], false],
  ];
  for (const [address, cidrs, expected] of cases) {
    assert.equal(isAllowedClientAddress(address, cidrs), expected, String(address));
  }

  const one = parseClientCidrAllowlist("128.0.0.0/1").cidrs;
  assert.equal(isAllowedClientAddress("128.0.0.1", one), true);
  assert.equal(isAllowedClientAddress("255.255.255.255", one), true);
  assert.equal(isAllowedClientAddress("100.0.0.1", one), false);
  const exact = parseClientCidrAllowlist("10.1.2.3/32").cidrs;
  assert.equal(isAllowedClientAddress("10.1.2.3", exact), true);
  assert.equal(isAllowedClientAddress("10.1.2.4", exact), false);
});
