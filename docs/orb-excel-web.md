# Developing in Amp orbs with Excel on the web

**Status:** working recipe (verified 2026-10-26). Orbs are Linux x64 without
KVM, so desktop Excel cannot run there. Excel on the web can: the taskpane
loads from the orb's dev server into Excel in the orb's Desktop Chrome, and
Office.js reports `host=Excel`, `platform=OfficeOnline`, ExcelApi up to 1.19.

Excel on the web is a development loop, not release acceptance: the release
checklist still requires desktop Excel (see
[release-smoke-test-checklist.md](./release-smoke-test-checklist.md)).

## What the orb provides

| Piece | Where | What it does |
| --- | --- | --- |
| Dependencies, Playwright Chromium | `.agents/setup` | `npm ci` (skipped when the lockfile is unchanged) |
| Trusted `https://localhost` | `.agents/setup` | Office dev CA in Chrome's NSS store; `key.pem`/`cert.pem` symlinks switch Vite and the proxy to HTTPS |
| Dev server `:3141` | `.amp/services.yaml` (`dev`) | Taskpane + UI gallery portals |
| CORS proxy `:3003` | `.amp/services.yaml` (`proxy`) | Model traffic for providers that need it, and the OAuth callback listeners (OpenAI `:1455`, Google `:8085`, ...) |
| Excel Chrome | `scripts/orb/excel-web.sh` | Chrome with a persistent profile and CDP on `127.0.0.1:9223`, Local Network Access checks off; headed in the Desktop, headless without one |
| Manifest upload helper | `scripts/orb/set-file-input.mjs` | Fills the cross-origin "Upload Add-in" file input over CDP |

Start the services with `amp orb services ensure`.

## Steps

1. **Open Excel** on a workbook (see [Sign-in](#sign-in) for which URL):
   `amp orb service start excel-web --command 'scripts/orb/excel-web.sh "<workbook url>"'`.
   For a visible window, run `amp orb desktop ensure` first (a cold orb has
   no Desktop); without one the launcher runs headless. Users can also run
   the script from an `amp-desktop-run` link.
2. **Upload the add-in.** Home → Add-ins → More Add-ins → My Add-ins →
   Manage My Add-ins → Upload My Add-in. In the Upload Add-in dialog run
   `node scripts/orb/set-file-input.mjs manifest.xml`, then press Upload.
   Banners such as "Your administrator has not installed Web Add-ins",
   "The Add-ins store has been disabled" and "Cannot connect to catalog."
   are normal and do not block the upload. A Developer Mode prompt may or
   may not appear; if it does, tick its checkbox explicitly before OK.
3. **Open Pi.** Home ribbon → `⋯` (More Options) → Open Pi. The upload lives
   in the browser profile, not the Microsoft account: it survives reloads
   and Chrome restarts, and a fresh profile needs it again. Microsoft may
   also drop it after some days
   ([install.md](./install.md#excel-on-the-web-office-online)).
4. **Log in to a provider** from the taskpane. The `proxy` service must be
   running first; otherwise the OAuth redirect lands on
   `http://localhost:1455/...` and the browser shows "connection refused".

Agents drive the same window with `agent-browser --cdp 9223 ...`. The taskpane
is a separate CDP target (`https://localhost:3141/src/taskpane.html?et=`), so
`Office`/`Excel` can be evaluated there directly over CDP.

## Gotchas

- **Blank taskpane:** Chrome's Local Network Access checks block the public
  Office origin from framing `https://localhost`. The launcher disables them;
  a Chrome started any other way shows an empty "Pi for Excel" pane.
- **No query-parameter sideloading.** `?wdaddindevserverport=3141&...` makes
  Excel fetch `https://127.0.0.1:3141/manifest.xml` cross-origin, which Vite's
  default CORS policy refuses. Do not widen `server.cors` for Office origins:
  it would also let those origins read `/__pi-auth`. Use the upload route.
- **CDP ports:** the orb Desktop's own Chrome owns `127.0.0.1:9222`; this
  window uses 9223. In Node use `127.0.0.1`, not `localhost` (which resolves
  to `::1`).
- **Clicking Office UI:** ribbon and dialogs live in cross-origin iframes that
  `agent-browser snapshot` often does not descend into. Most reliable: CDP
  `Runtime.evaluate` on the `excel.officeapps.live.com` iframe target and call
  `.click()` on the button (e.g. by `aria-label`). Coordinates work but shift
  with window size; a narrow window pushes Add-ins into the ribbon overflow.
- **Screenshots can lag a frame** in this Chrome: wait or take two.
- **`git am` in orbs** needs `-c commit.gpgsign=false` (no signing key).

## Sign-in

Signing in to Microsoft in each new orb can mean MFA. Two ways avoid it,
both verified in fresh orbs on 2026-10-26:

- **Anonymous edit link (preferred).** Share a scratch workbook from OneDrive
  as "Anyone with the link can edit" and open that link signed out. The
  workbook opens in Editing mode, Upload My Add-in works, and Office.js reads
  and writes cells. The link grants edit access to that one workbook only, so
  keep it in an Amp secret rather than in the repository.
- **Restored session.** Export the Microsoft-domain cookies of a signed-in
  profile with CDP `Storage.getCookies`, and in the new orb inject them with
  `Storage.setCookies` on the browser target before navigating (set host-only
  and `__Host-` cookies by `url`, not `domain`). Copying the profile directory
  also works: Chrome here encrypts cookies with its fixed Linux key, so they
  decrypt in another orb. Either way the material is a live session for the
  whole Microsoft account; prefer the anonymous link.

A manual sign-in needs the visible Desktop window (`amp orb desktop ensure`).
