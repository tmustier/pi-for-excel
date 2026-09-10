# Anti-slop lint policy

**Status:** Active tooling policy

**Scope:** Oxlint rules configured in `.oxlintrc.json`

Anti-slop lint is a small, reviewed set of deterministic restrictions. It is not a proxy for code quality, a full preset, or a reason to rewrite working code. A rule belongs in the gate only when it identifies a concrete harmful pattern, has useful positive and negative cases, and fits this repository's architecture.

## Enabled rules

The repository enables four rules as errors:

- `oxc/no-accumulating-spread` prevents quadratic array or object copies inside reducers.
- `anti-slop/no-reduce-accumulator-copy` covers equivalent copies through `Object.assign`, `concat`, `slice` and related APIs.
- `anti-slop/no-module-mocking` keeps tests on real dependency seams. Pass a faithful fake through an interface or callback instead of replacing an imported module.
- `anti-slop/no-widen-then-assert` rejects values that are deliberately widened and then asserted back to a narrower type. Preserve inference or parse the value at its boundary.

`tests/anti-slop-lint.test.mjs` invokes the pinned Oxlint CLI against temporary TS and MJS fixtures. Its negative cases assert every enabled diagnostic, so disabling any enabled rule makes the contract fail. Its positive cases protect mutation-based accumulation and injected fakes. Temporary fixtures are always removed.

## Deliberate exclusions

This is not the full upstream anti-slop preset. The vendored plugin remains an intact snapshot of upstream production source so that provenance and future comparisons stay clear, but unlisted rules are not repository policy.

In particular:

- The `no-unknown-*` rules are not enabled while `DynamicValue` exists. The alias can conceal the syntax they inspect, so global enforcement would overstate boundary safety. The existing ESLint restriction on direct `unknown` syntax remains in force until a coordinated boundary migration replaces the alias with honest raw-input types and concrete decoders.
- Runtime `typeof` checks and small parser primitives remain valid boundary tools. An object check alone does not establish a domain contract; concrete parsers own domain shapes.
- `filter().map()`, object parameters, clear conditional object spreads, and descriptive symbol names are not banned by style alone.
- Production code uses typed property access rather than `Reflect.get` or `Reflect.apply`. Test harnesses still use reflection to install host globals and exercise malformed JavaScript calls, so the rules are not enabled across the repository.
- Return annotations are required where they make exported or public contracts clearer, not as a blanket anti-slop principle for every private helper.
- Assertions do not require boilerplate safety comments. Unavoidable interop assertions and lint suppressions need a concrete local invariant, while ordinary code should narrow or parse instead.
- The known-value-widening rule currently also rejects useful explicit public return types and dictionary contracts, and chained assertions remain in test-only DOM interop. Enable neither until the architecture can comply without concealing necessary contracts.

There are no blanket exemptions for existing practice. If an enabled rule finds existing code, either correct the harmful pattern or review a narrow, documented exception based on the rule's actual false positive. Do not add broad baselines, directory-wide disables, or repetitive safety comments.

## Vendor and update policy

The plugin production source is vendored under `tools/oxlint/anti-slop` from the immutable revision recorded in `UPSTREAM.md`, with the upstream MIT license unchanged. Do not edit vendored rule implementations as part of ordinary application work.

For an update:

1. compare the complete upstream production source with the recorded revision;
2. copy the new production snapshot without trimming files merely because their rules are disabled;
3. preserve the upstream MIT license and record the exact new commit;
4. run the real-CLI contract tests and add a positive and negative case before enabling another rule;
5. review any new diagnostic against the architecture rather than grandfathering all current code.
