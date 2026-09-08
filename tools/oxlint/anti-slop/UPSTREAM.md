# Vendored anti-slop plugin

## Provenance

- Source: `https://github.com/dmmulroy/anti-slop`
- Upstream commit: `95a56e5d24fb3d849673c2d51eb0908b8bd2d33b`
- Vendored path: upstream production `src/` copied to this directory
- License: MIT; see `LICENSE` in this directory
- Dependencies: `oxlint` and `@oxlint/plugins`, both pinned to `1.82.0`

The vendored implementation is owned and reviewed as repository tooling. Update it by
comparing against an immutable upstream revision; do not install the upstream skill
globally or replace local policy with an upstream preset.

## Local policy

This first adoption deliberately enables only the following reviewed rules at error
level:

- native `oxc/no-accumulating-spread`
- `anti-slop/no-reduce-accumulator-copy`
- `anti-slop/no-module-mocking`
- `anti-slop/no-widen-then-assert`
- `anti-slop/no-unknown-parameters`
- `anti-slop/no-unknown-returns`

The remaining vendored generic rules are available for later, rule-by-rule rollout.
This is not the full anti-slop preset. In particular, local policy does not currently
ban runtime `typeof`, conditional optional object spreads, names containing `shape`,
or `Reflect` calls. The optional Effect plugin is vendored for source fidelity but is
not registered because this project does not directly depend on Effect.

Application lint covers root TypeScript, `src/**/*.ts`, and both TypeScript and MJS
tests. The vendored plugin is ignored by application lint and should be verified with
focused fixtures whenever its implementation changes.
