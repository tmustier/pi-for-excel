# macOS adversarial extension-provider smoke — 2026-07-16

## Build and environment

- Commit: `d7bba2f` (`test/adversarial-extension-provider`)
- Host: Microsoft Excel for Mac, real Office.js taskpane
- Taskpane: `https://localhost:3141/src/taskpane.html`
- Background verification bridge: token-authenticated loopback HTTPS on `3157`
- CORS proxy: loopback HTTPS on `3003`, explicitly restricted to target host `localhost`
- Instrumented model gateways: HTTPS ports `3161` and `3162`
- Runtime: pasted inline extension in `sandbox-iframe` mode

## Result

Overall: **Pass**

Command:

```bash
npm run smoke:extension-provider
```

The runner completed with `"ok": true`. It installed pasted code through the real extension manager, granted capabilities through the persisted permission path, configured a host-owned connection, selected the discovered extension model in the real model selector and completed four exact streamed responses:

- `EXTENSION_PROVIDER_A_OK`
- `EXTENSION_PROVIDER_CACHE_OK`
- `EXTENSION_PROVIDER_B_OK`
- `EXTENSION_PROVIDER_DELAYED_OK`

## Focused checklist evidence

| Checklist | Result | Evidence |
|---|---|---|
| C-5 self-extension flow | Pass | Inline code activated in the sandbox iframe; owner-qualified provider appeared in the real selector and completed inference. |
| H-1 error paths | Pass | Permission denials, discovery outage, oversized catalogue, endpoint upgrade, in-flight unload and delayed discovery race all reached deterministic recovery states. |
| H-3 proxy/security boundaries | Pass | Both discovery and inference received host-injected test credentials; sandbox secret reads were denied; the old gateway never received the replacement credential; a non-allowlisted endpoint host was rejected. |
| H-4 storage/cache tolerance | Pass | Last safe catalogue survived an outage and oversized refresh; cached model metadata rebound to gateway B; an aborted late response did not resurrect deleted cache data. |

## Defects found by the adversarial run

The initial run exposed five real failures that narrower mocked tests had not covered:

1. Generated sandbox bootstrap JavaScript contained an unescaped newline and did not execute in Excel.
2. Dynamic discovery accepted unbounded response and model-list sizes.
3. Selecting the same provider/model ID after an endpoint upgrade retained stale session transport metadata.
4. Unloading a provider during an in-flight turn left that removed provider selected after the turn.
5. A delayed discovery response could write its catalogue after provider unregister/cache deletion.

The tested commit fixes each failure and adds focused regression coverage.

## Cleanup

- Probe and hostile-test extensions uninstalled.
- Permission-gate experiment reset.
- Active runtime reconciled to `openai-codex/gpt-5.6-sol`.
- No test credential values were printed by the runner.
