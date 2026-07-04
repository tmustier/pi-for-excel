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
- an `ignoreCommand` deploy policy (`node scripts/vercel-ignore-command.mjs`) for `main`, PR previews, and manual deploys
- `/proxy` rewrite to `/proxy.sh` (bootstrap script for `npx pi-for-excel-proxy`)
- a header rule to disable caching for `/src/taskpane.html` to make updates propagate reliably
- an enforced `Content-Security-Policy` on `/src/taskpane.html` (Office.js + provider/auth endpoints + localhost proxy + Pyodide CDN host).

### `ignoreCommand` policy

Automatic deploy behavior is:
- **build** for `main`
- **build** for pull requests (`VERCEL_GIT_PULL_REQUEST_ID` is set)
- **build** for manual deploys (`VERCEL_GIT_COMMIT_REF` is unset)
- **skip** non-PR feature branch pushes

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
