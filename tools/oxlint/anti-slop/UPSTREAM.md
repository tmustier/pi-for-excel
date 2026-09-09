# Vendored anti-slop plugin

## Provenance

- Source: `https://github.com/dmmulroy/anti-slop`
- Upstream commit: `95a56e5d24fb3d849673c2d51eb0908b8bd2d33b`
- Vendored path: the complete upstream production `src/` snapshot copied to this directory
- License: MIT; the upstream license is preserved unchanged in `LICENSE`
- Dependencies: `oxlint` and `@oxlint/plugins`, both pinned to `1.82.0`

The vendored implementation is owned and reviewed as repository tooling. Update it by comparing against an immutable upstream revision; do not install the upstream skill globally or replace local policy with an upstream preset. Keep the production snapshot intact rather than deleting rule modules or shared code merely because a rule is not enabled.

## Local policy

This repository deliberately enables only the following reviewed rules at error level:

- native `oxc/no-accumulating-spread`
- `anti-slop/no-reduce-accumulator-copy`
- `anti-slop/no-module-mocking`
- `anti-slop/no-widen-then-assert`

This is not the full anti-slop preset. The other vendored rules are source for possible rule-by-rule review, not active repository policy. In particular, local policy does not globally ban raw boundary checks, conditional optional object spreads, names containing `shape`, or reflection in test harnesses. Production application code contains no `Reflect.get` or `Reflect.apply`. The syntax-based `unknown` rules remain disabled while `DynamicValue` can evade them, making their apparent coverage misleading. The optional Effect plugin is vendored for source fidelity but is not registered because this project does not directly depend on Effect.

Application lint covers root TypeScript, `src/**/*.ts`, and both TypeScript and MJS tests. The vendored plugin is ignored by application lint. `tests/anti-slop-lint.test.mjs` invokes the real pinned Oxlint CLI with positive and negative TS/MJS fixtures for every enabled policy; changing enabled rules requires updating that contract deliberately.

See `docs/anti-slop-policy.md` for rationale, exclusions and the update process.
