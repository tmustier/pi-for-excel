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

test("isIpLiteral detects IPv4/IPv6 literals", () => {
  assert.equal(isIpLiteral("127.0.0.1"), true);
  assert.equal(isIpLiteral("[::1]"), true);
  assert.equal(isIpLiteral("::1"), true);
  assert.equal(isIpLiteral("localhost"), false);
  assert.equal(isIpLiteral("api.openai.com"), false);
});

test("isLoopbackHostname covers localhost + mapped loopback", () => {
  assert.equal(isLoopbackHostname("localhost"), true);
  assert.equal(isLoopbackHostname("127.0.0.1"), true);
  assert.equal(isLoopbackHostname("::1"), true);
  assert.equal(isLoopbackHostname("[::1]"), true);
  assert.equal(isLoopbackHostname("::ffff:127.0.0.1"), true);
  assert.equal(isLoopbackHostname("example.com"), false);
});

test("isPrivateOrLocalIp classifies private/local ranges", () => {
  assert.equal(isPrivateOrLocalIp("127.0.0.1"), true);
  assert.equal(isPrivateOrLocalIp("10.2.3.4"), true);
  assert.equal(isPrivateOrLocalIp("172.16.5.5"), true);
  assert.equal(isPrivateOrLocalIp("172.31.255.255"), true);
  assert.equal(isPrivateOrLocalIp("192.168.1.2"), true);
  assert.equal(isPrivateOrLocalIp("169.254.12.9"), true);
  assert.equal(isPrivateOrLocalIp("::1"), true);
  assert.equal(isPrivateOrLocalIp("fc00::1"), true);
  assert.equal(isPrivateOrLocalIp("fd12:3456::1"), true);
  assert.equal(isPrivateOrLocalIp("fe80::1"), true);

  assert.equal(isPrivateOrLocalIp("8.8.8.8"), false);
  assert.equal(isPrivateOrLocalIp("1.1.1.1"), false);
  assert.equal(isPrivateOrLocalIp("2001:4860:4860::8888"), false);
});

test("parseAllowedTargetHosts accepts plain hosts and URLs", () => {
  const allowed = parseAllowedTargetHosts("api.openai.com, https://oauth2.googleapis.com/token ,[::1]");

  assert.equal(allowed.has("api.openai.com"), true);
  assert.equal(allowed.has("oauth2.googleapis.com"), true);
  assert.equal(allowed.has("::1"), true);

  assert.equal(isAllowedTargetHost("api.openai.com", allowed), true);
  assert.equal(isAllowedTargetHost("example.com", allowed), false);
});

test("evaluateTargetHostPolicy blocks loopback/private by default", () => {
  const loopback = evaluateTargetHostPolicy({ hostname: "127.0.0.1" });
  assert.deepEqual(loopback, { allowed: false, reason: "blocked_target_loopback" });

  const privateIp = evaluateTargetHostPolicy({ hostname: "10.0.0.10" });
  assert.deepEqual(privateIp, { allowed: false, reason: "blocked_target_private_ip" });

  const dnsPrivate = evaluateTargetHostPolicy({
    hostname: "api.example.com",
    resolvedIps: ["192.168.1.22"],
  });
  assert.deepEqual(dnsPrivate, { allowed: false, reason: "blocked_target_private_ip" });
});

test("evaluateTargetHostPolicy supports overrides", () => {
  const allowLoopback = evaluateTargetHostPolicy({
    hostname: "127.0.0.1",
    allowLoopbackTargets: true,
  });
  assert.deepEqual(allowLoopback, { allowed: true });

  const allowPrivate = evaluateTargetHostPolicy({
    hostname: "10.0.0.10",
    allowPrivateTargets: true,
  });
  assert.deepEqual(allowPrivate, { allowed: true });

  const allowlistedOnly = evaluateTargetHostPolicy({
    hostname: "api.openai.com",
    allowedHosts: parseAllowedTargetHosts("api.openai.com"),
  });
  assert.deepEqual(allowlistedOnly, { allowed: true });

  const blockedByAllowlist = evaluateTargetHostPolicy({
    hostname: "example.com",
    allowedHosts: parseAllowedTargetHosts("api.openai.com"),
  });
  assert.deepEqual(blockedByAllowlist, {
    allowed: false,
    reason: "blocked_target_not_allowlisted",
  });
});

test("loopback/private checks run before host allowlist", () => {
  const strictAllowlist = parseAllowedTargetHosts("api.openai.com");

  const blockedLoopback = evaluateTargetHostPolicy({
    hostname: "127.0.0.1",
    allowedHosts: strictAllowlist,
  });
  assert.deepEqual(blockedLoopback, {
    allowed: false,
    reason: "blocked_target_loopback",
  });

  const allowedLocalWithOverrides = evaluateTargetHostPolicy({
    hostname: "127.0.0.1",
    allowLoopbackTargets: true,
    allowPrivateTargets: true,
    allowedHosts: strictAllowlist,
  });
  assert.deepEqual(allowedLocalWithOverrides, { allowed: true });
});

test("isBlockedTargetByHostname reflects default deny policy", () => {
  assert.equal(isBlockedTargetByHostname("localhost"), true);
  assert.equal(isBlockedTargetByHostname("10.0.0.8"), true);
  assert.equal(isBlockedTargetByHostname("api.openai.com"), false);
});
