# Security dependency baseline

Last verified: 17 October 2026.

The October refresh updates vulnerable transitive dependencies within the existing direct dependency ranges and moves the Axios override to 1.20.0. Pi, TypeScript and the direct Office tooling dependencies stay on their existing major versions. It includes the shell-quote, ip-address, proxy-addr and Handlebars fixes proposed in PRs #758, #759, #762 and #763.

## Temporary node-forge patch

CVE-2026-85393 ([GHSA-86w9-cpqp-85rv](https://github.com/advisories/GHSA-86w9-cpqp-85rv)) affects every published node-forge version through 1.4.0. No patched npm release was available on the verification date. Downgrading Office tooling to satisfy npm audit would remove supported behavior.

With maintainer approval, the development dependency uses the immutable archive for [upstream PR #1158, commit f192c67ee5850e44296ef83d6a7b6df6552ba8bf](https://github.com/digitalbazaar/forge/pull/1158). This is an unmerged community patch, not an official release. Its package version is upstream's existing 1.4.1-0 development version. The archive has a SHA-512 integrity entry in `package-lock.json`.

The runtime change adds a nested DigestAlgorithm element-count check to RSA PKCS#1 v1.5 verification. The complete comparison against the 1.4.0 release also contains upstream's development-version increment, security-policy documentation, removal of an accidental test-only marker, and the new regression test. No other library implementation changes are included.

The root override uses `$node-forge` so both `mkcert` and the Microsoft Teams toolkit resolve the same patched implementation. Certificate generation, installation and desktop sideload capabilities remain present. `.npmrc` sets npm 12's `allow-remote=root` for this explicitly declared archive. It does not permit arbitrary transitive remote archives or enable Git dependencies. Older npm versions may warn about the new configuration key.

`tests/security-node-forge.test.mjs` rejects the malformed nested DigestAlgorithm, checks valid RSA signatures with exponents 3 and 65537, rejects a changed message, and independently verifies a localhost certificate chain generated through Office's resolved `mkcert` dependency. The malformed-signature test fails with published node-forge 1.4.0 and passes with the pinned patch. Certificates stay in memory; the tests do not change the system trust store.

A zero-finding audit alone is not proof that an unpublished patch works. The regression and compatibility tests provide the behavior evidence; the unchanged `npm audit --audit-level=high` gate continues to check the dependency graph.

## Remove the temporary pin

When an official release fixes CVE-2026-85393:

1. Verify the release includes the nested element-count fix.
2. Replace the archive dependency with that released version, keeping one patched copy for every consumer.
3. Remove `.npmrc` when no root archive is needed.
4. Regenerate the lockfile and run the regression tests, full checks, audit and real Excel acceptance again.
