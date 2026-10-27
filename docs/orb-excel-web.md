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
| Session script | `scripts/orb/excel-web-session.mjs` | One command from a fresh orb to Pi open in its own workbook |
| Blank workbook | `scripts/orb/blank.xlsx` | Default workbook each orb uploads |
| Model login | Amp project secret `PI_EXCEL_WEB_AUTH_JSON` | pi `auth.json` contents; written to `~/.pi/agent/auth.json`, which the dev server serves to the taskpane at `/__pi-auth`, so Pi starts logged in |
| Shared folder link | Amp project secret `PI_EXCEL_WEB_FOLDER_LINK` | "Anyone with the link can edit" link to a OneDrive folder (see [Sign-in](#sign-in)) |

## One command

```bash
node scripts/orb/excel-web-session.mjs --check
```

It runs `amp orb services ensure`, starts the Excel Chrome if CDP is not
answering, opens the shared folder signed out, uploads this orb's own
workbook (`orb-<thread>-<timestamp>.xlsx`, or `--workbook <file.xlsx>`) as
the guest, opens it in Excel, uploads `manifest.xml`, and opens Pi. `--check`
then writes, reads back and clears Z1 through Office.js in the taskpane. It
prints the workbook name and URL as JSON. A cold orb takes about 50 s
(including starting the services); later new sessions about 40 s.

Add `--prompt "<message>"` to send Pi a real agent turn after the setup; the
command waits for the turn to finish and prints the model, the tools Pi
called and the end of its reply under `turn`. If Pi shows a login overlay
instead (missing or rejected credentials), it fails at once. Check the workbook yourself
afterwards (Office.js in the taskpane target) rather than trusting the reply.

`PI_EXCEL_WEB_AUTH_JSON` currently holds an `openai-codex` (ChatGPT) OAuth credential.
Pi only refreshes it once the access token expires; refreshing rotates the
refresh token and logs out every other copy, so when it nears expiry,
replace the secret with a freshly exported one instead of letting orbs
refresh it. Every orb shares that subscription's rate limits.

The secret has its own name because a personal secret named `PI_AUTH_JSON`
(used by pi elsewhere) takes precedence over a project secret of that name.

A rerun while Pi is open reuses that workbook (`"reused": true`) and only
repeats the check. Otherwise each run uploads a new workbook, so parallel
orbs never share one. Guests cannot delete files, so the folder's owner
clears out old `orb-*.xlsx` files now and then.

In a cold orb the Excel Chrome runs headless, so the user cannot see it.
For a visible window run `amp orb desktop ensure` before the first run; if
Chrome already started headless, run `amp orb service stop excel-web` and
rerun the command.

`.agents/setup` runs when an orb is created. In an orb created before these
files reached the default branch, run `.agents/setup` by hand first.

## Manual steps

What the script automates, for when Office's UI changes and it breaks:

1. **Open Excel** on a workbook:
   `amp orb service start excel-web --command 'scripts/orb/excel-web.sh "<workbook url>"'`.
   For a visible window, run `amp orb desktop ensure` first (a cold orb has
   no Desktop); without one the launcher runs headless. Users can also run
   the script from an `amp-desktop-run` link.
2. **Upload the add-in.** Home → Add-ins → More Add-ins → My Add-ins →
   Manage My Add-ins → Upload My Add-in. The file input is in a
   cross-origin iframe; set it with CDP `DOM.setFileInputFiles` on the
   `res.public.onecdn.static.microsoft` iframe target, then press Upload.
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

## Record PR evidence

Open Pi first with `node scripts/orb/excel-web-session.mjs --check`, then:

```bash
node scripts/orb/record-excel-web.mjs .amp/in/artifacts/excel-turn.mp4 -- node scripts/orb/excel-web-session.mjs --prompt "Put 1 and 2 in A1:A2 and their sum in A3, then tell me A3."
gh pr comment <number> --body "Real Excel taskpane turn" --attach .amp/in/artifacts/excel-turn.mp4
```

The wrapper records the single open workbook tab (Excel **and** Pi) over CDP
at `127.0.0.1:${PI_EXCEL_WEB_CDP_PORT:-9223}` until the command exits; any
command works, including `agent-browser --cdp 9223 ...` or a manual driver.
It preserves CDP frame timing, including idle time, and the command's exit
status. Requires system `ffmpeg` (already present in the orb); existing outputs
are not overwritten. Default: silent H.264 MP4, at most 1600×1000, 15 fps,
2 Mb/s ceiling (~15 MB/minute), CRF 28/JPEG 70 — **agent-default (not a user rule): 2026-10-27**.
Keep clips short, inspect for secrets and readability, independently read back
the workbook, and keep videos out of Git. `gh pr create/comment --attach`
requires GitHub CLI 2.99+ and credentials that support uploads. The orb's
GitHub App token cannot upload attachments, so attach from a machine with a
normal `gh` login instead: download the orb's file with
`amp files get <attachment-url> -o excel-turn.mp4`, then run the `gh pr comment
... --attach excel-turn.mp4` line there. `gh` cannot set alt text on videos, so
describe the recording in `--body`.

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

- **Anonymous folder link (preferred, what the script uses).** Share a
  OneDrive folder as "Anyone with the link can edit" and store the link as
  the project secret `PI_EXCEL_WEB_FOLDER_LINK`. Signed out, the browser gets
  a SharePoint REST session as "Guest Contributor" that can upload files;
  their workbooks open in Editing mode, Upload My Add-in works, and Office.js
  reads and writes cells. The link grants edit access to that folder only.
  A link to a single workbook works the same way but is shared by every orb.
- **Restored session.** Export the Microsoft-domain cookies of a signed-in
  profile with CDP `Storage.getCookies`, and in the new orb inject them with
  `Storage.setCookies` on the browser target before navigating (set host-only
  and `__Host-` cookies by `url`, not `domain`). Copying the profile directory
  also works: Chrome here encrypts cookies with its fixed Linux key, so they
  decrypt in another orb. Either way the material is a live session for the
  whole Microsoft account; prefer the anonymous link.

A manual sign-in needs the visible Desktop window (`amp orb desktop ensure`).
