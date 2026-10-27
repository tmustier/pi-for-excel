# Deploy hosted build on Vercel (maintainers)

Pi for Excel’s taskpane is a static site built by Vite (`dist/`).

Vercel is a good default host because it’s free for OSS/hobby usage and handles HTTPS + caching well.

## One-time setup

1. Create a new Vercel project
2. Import `tmustier/pi-for-excel`
3. Framework preset: **Vite** (or “Other”)
4. Build settings:
   - **Build Command:** `npm run build`
   - **Output Directory:** `dist`

This repo includes `vercel.json` with:
- `outputDirectory: dist`
- `git.deploymentEnabled` limiting automatic deployments to `main`, plus an `ignoreCommand` fallback (`node scripts/vercel-ignore-command.mjs`)
- `/proxy` rewrite to `/proxy.sh` (bootstrap script for `npx pi-for-excel-proxy`)
- a header rule to disable caching for `/src/taskpane.html` to make updates propagate reliably
- an enforced `Content-Security-Policy` on `/src/taskpane.html` (Office.js + provider/auth endpoints + localhost proxy + Pyodide CDN host).

### Automatic deployment policy

`git.deploymentEnabled` creates automatic deployments only for `main`:
`"**": false` turns off every branch (`**` also matches names with slashes),
and `"main": true` turns `main` back on, because Vercel deploys when any
matching rule is `true`. PR branches, every layer of a PR stack and
dependency PRs get no preview deployment; CI's `build-check-test` covers the
build. Previews were turned off because each one counts toward Vercel's
daily deployment quota, and busy days of PR pushes used it up and blocked
production deploys of `main`. Previews also could not use the local proxy,
whose origin allowlist only includes the production URL.

The configuration must be present in the pushed branch's commit. Branches
cut before this change keep their old behaviour until they merge or rebase
onto `main`.

This is intentionally not a blanket Markdown/docs exclusion: `README.md` and
non-archive `docs/**/*.md` are bundled by `src/files/builtin-docs.ts`.

### `ignoreCommand` fallback

If a deployment is created anyway (for example a manual deploy), the ignored-build policy is:
- **build** for `main`
- **build** for pull requests (`VERCEL_GIT_PULL_REQUEST_ID` is set)
- **build** for manual deploys (`VERCEL_GIT_COMMIT_REF` is unset)
- **skip** non-PR feature branch pushes

An ignored build **still counts toward Vercel's daily deployment quota**;
it saves build work, not deployments. Use the branch gate above to avoid
creating known-unneeded deployments, not another ignored-build condition.
See [Vercel's Ignored Build Step documentation](https://vercel.com/docs/project-configuration/project-settings#ignored-build-step).

Regression coverage lives in `tests/vercel-ignore-command.test.mjs` (run via `npm run test:security`).

If a host-specific regression appears, temporary rollback is a single-header change:
`Content-Security-Policy` → `Content-Security-Policy-Report-Only`.

## Production URL

After deploy, you’ll have a production URL like:

- `https://<project>.vercel.app`

Keep this URL stable; it becomes the base URL used by `manifest.prod.xml`.

## Generate / update the production manifest

The dev manifest (`manifest.xml`) points at `https://localhost:3141`.

Generate the production manifest with the hosted base URL:

```bash
ADDIN_BASE_URL="https://<project>.vercel.app" npm run manifest:prod
```

This writes:
- `manifest.prod.xml` (repo root)
- `public/manifest.prod.xml` (so the hosted site can offer a one-click download at `/manifest.prod.xml`)

## Updates (automatic)

For most UI/behavior changes:
- deploy a new build to the same Vercel project
- users get the update automatically next time they open the taskpane

If a release requires a manifest change (rare):
- update and redistribute `manifest.prod.xml`
- users re-upload the manifest in Excel
