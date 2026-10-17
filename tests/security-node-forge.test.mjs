import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { generateKeyPairSync, sign, X509Certificate } from "node:crypto";
import { test } from "node:test";
import forge from "node-forge";

const require = createRequire(import.meta.url);
// Exercise the certificate generator actually resolved by Office's dev tooling.
const requireFromDevCerts = createRequire(require.resolve("office-addin-dev-certs"));
const { createCA, createCert } = requireFromDevCerts("mkcert");

void test("RSA verification rejects extra nested DigestAlgorithm elements (CVE-2026-85393)", () => {
  // Upstream regression fixture from digitalbazaar/forge#1158. Exponent 1
  // exposes the decoded signature block directly, isolating DER validation.
  const publicKey = forge.pki.rsa.setPublicKey(
    new forge.jsbn.BigInteger("f".repeat(256), 16),
    new forge.jsbn.BigInteger("1", 10),
  );
  const md = forge.md.sha256.create();
  md.update("nested-digest-algorithm");
  const digest = md.digest().getBytes();
  const digestInfo = forge.util.hexToBytes("30343010060960864801650304020105000401000420") + digest;
  const padding = forge.util.hexToBytes(`0001${"ff".repeat(128 - digestInfo.length - 3)}00`);
  assert.throws(() => publicKey.verify(digest, padding + digestInfo), /DigestInfo/);
});

void test("RSA verification accepts valid signatures and rejects changed messages", () => {
  for (const publicExponent of [3, 65537]) {
    const { publicKey, privateKey } = generateKeyPairSync("rsa", {
      modulusLength: 2048,
      publicExponent,
    });
    const verifier = forge.pki.publicKeyFromPem(publicKey.export({ type: "spki", format: "pem" }));
    const message = "security baseline signature control";
    const signature = sign("sha256", Buffer.from(message), privateKey).toString("binary");
    const md = forge.md.sha256.create();
    md.update(message);
    assert.equal(verifier.verify(md.digest().getBytes(), signature), true);
    md.update("changed");
    assert.equal(verifier.verify(md.digest().getBytes(), signature), false);
  }
});

void test("Office certificate generator creates a valid localhost certificate chain", async () => {
  // In-memory fixtures only. No trust-store installation or user certificate changes.
  const ca = await createCA({
    organization: "Pi security test CA",
    countryCode: "GB",
    state: "London",
    locality: "London",
    validity: 1,
  });
  const leaf = await createCert({ ca, domains: ["localhost", "127.0.0.1"], validity: 1 });
  const rootCertificate = new X509Certificate(ca.cert);
  const leafCertificate = new X509Certificate(leaf.cert);
  assert.equal(rootCertificate.ca, true);
  assert.equal(rootCertificate.verify(rootCertificate.publicKey), true);
  assert.equal(leafCertificate.verify(rootCertificate.publicKey), true);
  assert.equal(leafCertificate.checkHost("localhost"), "localhost");
  assert.equal(leafCertificate.checkIP("127.0.0.1"), "127.0.0.1");
});
