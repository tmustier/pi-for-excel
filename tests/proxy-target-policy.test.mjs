import test from "node:test";
import assert from "node:assert/strict";

import {
  evaluateTargetHostPolicy,
  isAllowedTargetHost,
  isBlockedTargetByHostname,
  isIpLiteral,
  isLoopbackHostname,
  isPrivateOrLocalIp,
  parseAllowedTargetHosts,
} from "../scripts/proxy-target-policy.mjs";

test("target policy classifies literal, loopback, private, and public hosts", () => {
  const literalCases = [["127.0.0.1", true], ["[::1]", true], ["::1", true], ["localhost", false], ["api.openai.com", false]];
  for (const [host, expected] of literalCases) assert.equal(isIpLiteral(host), expected, host);

  const loopbackCases = [["localhost", true], ["127.0.0.1", true], ["::1", true], ["[::1]", true], ["::ffff:127.0.0.1", true], ["example.com", false]];
  for (const [host, expected] of loopbackCases) assert.equal(isLoopbackHostname(host), expected, host);

  const privateCases = [
    ["127.0.0.1", true], ["10.2.3.4", true], ["172.16.5.5", true], ["172.31.255.255", true],
    ["192.168.1.2", true], ["169.254.12.9", true], ["::1", true], ["fc00::1", true],
    ["fd12:3456::1", true], ["fe80::1", true], ["8.8.8.8", false], ["1.1.1.1", false],
    ["2001:4860:4860::8888", false],
  ];
  for (const [host, expected] of privateCases) assert.equal(isPrivateOrLocalIp(host), expected, host);
});

test("target allowlist parser accepts host and URL forms", () => {
  const allowed = parseAllowedTargetHosts("api.openai.com, https://oauth2.googleapis.com/token ,[::1]");
  for (const host of ["api.openai.com", "oauth2.googleapis.com", "::1"]) assert.equal(allowed.has(host), true, host);
  assert.equal(isAllowedTargetHost("api.openai.com", allowed), true);
  assert.equal(isAllowedTargetHost("example.com", allowed), false);
});

test("target-host policy applies deny, allowlist, and override precedence", () => {
  const strict = parseAllowedTargetHosts("api.openai.com");
  const cases = [
    [{ hostname: "127.0.0.1" }, { allowed: false, reason: "blocked_target_loopback" }],
    [{ hostname: "10.0.0.10" }, { allowed: false, reason: "blocked_target_private_ip" }],
    [{ hostname: "api.example.com", resolvedIps: ["192.168.1.22"] }, { allowed: false, reason: "blocked_target_private_ip" }],
    [{ hostname: "127.0.0.1", allowLoopbackTargets: true }, { allowed: true }],
    [{ hostname: "10.0.0.10", allowPrivateTargets: true }, { allowed: true }],
    [{ hostname: "api.openai.com", allowedHosts: strict }, { allowed: true }],
    [{ hostname: "example.com", allowedHosts: strict }, { allowed: false, reason: "blocked_target_not_allowlisted" }],
    [{ hostname: "127.0.0.1", allowedHosts: strict }, { allowed: false, reason: "blocked_target_loopback" }],
    [{ hostname: "127.0.0.1", allowLoopbackTargets: true, allowPrivateTargets: true, allowedHosts: strict }, { allowed: true }],
    [{ hostname: "10.0.0.5", allowPrivateTargets: true, allowedHosts: new Set(["api.openai.com", "10.97.193.77"]), requireAllowlistForOverriddenTargets: true }, { allowed: false, reason: "blocked_target_not_allowlisted" }],
    [{ hostname: "10.97.193.77", allowPrivateTargets: true, allowedHosts: new Set(["api.openai.com", "10.97.193.77"]), requireAllowlistForOverriddenTargets: true }, { allowed: true }],
    [{ hostname: "127.0.0.1", allowLoopbackTargets: true, allowedHosts: new Set(["api.openai.com", "10.97.193.77"]), requireAllowlistForOverriddenTargets: true }, { allowed: false, reason: "blocked_target_not_allowlisted" }],
    [{ hostname: "localhost", allowLoopbackTargets: true, allowedHosts: new Set(["localhost"]), requireAllowlistForOverriddenTargets: true }, { allowed: true }],
    [{ hostname: "127.0.0.1", allowLoopbackTargets: true, allowedHosts: strict }, { allowed: true }],
    [{ hostname: "10.0.0.5", allowPrivateTargets: true, allowedHosts: strict }, { allowed: true }],
  ];
  for (const [input, expected] of cases) assert.deepEqual(evaluateTargetHostPolicy(input), expected, JSON.stringify(input));

  assert.equal(isBlockedTargetByHostname("localhost"), true);
  assert.equal(isBlockedTargetByHostname("10.0.0.8"), true);
  assert.equal(isBlockedTargetByHostname("api.openai.com"), false);
});
