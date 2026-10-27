# Integration verification: setup

This is a rerun of every item in [../setup.md](../setup.md) against the combined fix build: branch `integration-fixes` at `a2b3f2d`, which is `main` plus PRs #786–#799 (see `INTEGRATION.md` on that branch). It ran on 27 October 2026, by an agent in an Amp orb. Expected behavior is the checklist, as amended by the fix PRs. Where a PR deliberately changed behavior, the PR wins, and the note names the PR.

How it ran:

- **web**: Pi in Excel on the web through `node scripts/orb/excel-web-session.mjs --check`, with the dev server started from this checkout (host Excel, platform OfficeOnline, write and read-back ok). The taskpane was driven through CDP on its own target, using `Input.insertText`, key and mouse events, and DOM reads. Reloads used `location.reload()` in the taskpane target.
- **standalone**: Playwright Chromium on `https://localhost:3141/src/taskpane.html`, with a fresh context for every item and `/__pi-auth` stubbed (empty, a made-up API key, or a fake OAuth login). The key check (`/api-proxy/openai/v1/models`), the Anthropic token endpoint and model API, `chatgpt.com` and proxy health were stubbed where an item needed it.
- **production build**: `npx vite build --outDir /tmp/prod`, served statically over HTTPS on `:4443`. This was used for the proxy gate.
- **mock gateway**: a local HTTPS OpenAI-compatible stub on `:4010`. It streams twelve words, one a second, with CORS, and was used as a direct gateway for the busy and second-provider items.
- Real model turns on the shared ChatGPT login: four tiny turns on GPT-6 Luna, GPT-6.1 Sol and GPT-6 Sol. The shared credential was never logged out or refreshed.

**Result: 134 items. 132 pass, 2 fail, 0 blocked.** Every item that failed on `main` now passes or behaves as its owning PR describes. One item that passed on `main` now fails: PROXY-10, but its code is unchanged from `main` (see below). PROXY-07 was a "suspected bug confirmed" on `main`; #793's fix for it is only partial.

## Results

| ID | Result | Note |
| --- | --- | --- |
| MODEL-01 | pass | "+" from a GPT-6.1 Sol tab: the new "Chat 3" shows "GPT-6.1 Sol" and "High". |
| MODEL-02 | pass | After a reload, all four tabs came back on their models (GPT-6 Luna, GPT-6.1 Sol, GPT-6 Luna, GPT-6 Sol). Fork titles now survive too (#795): "Chat 1 (gpt-6.1-sol)", "Chat 1 (gpt-6.1-sol) (gpt-6-sol)". |
| MODEL-03 | pass (#795) | Saving `sk-fake-docs-test-key` now checks the key: "Checking key…", then "OpenAI (API) rejected this key (HTTP 401). Check it and try again." Nothing is saved. With the key check stubbed to 200, the key saves and "hello" fails with "Couldn’t get a response: OpenAI API error (401): …". |
| MODEL-04 | pass | Rows in the documented order. |
| MODEL-05 | pass | Proxy off: "Reply with just: OK" on GPT-6 Luna succeeded, with the request sent straight to `chatgpt.com/backend-api/codex/responses` (as the corrected document says). |
| MODEL-06 | pass | "Google Gemini (API) disconnected"; the row shows "set up →". |
| MODEL-07 | pass | After its gateway was deleted, the tab moved to "GPT-6.1 Sol" "High". |
| MODEL-08 | pass (#795) | Fork on, empty "Chat 3", picked gpt-6-luna: it switched in place, with no new tab and no toast. The suspected bug is fixed. |
| MODEL-09 | pass (#795) | While busy, the picker opens read-only: every row is `aria-disabled`, with the note "Pi is working in this tab. Change the model when it finishes, or press Stop." A click leaves it open and the model unchanged. There is no refusal toast any more. |
| MODEL-10 | pass (#795) | The thinking level is now clamped to the new model: High → gateway model shows "Off". (The checklist expected "High", from before #795.) |
| MODEL-11 | pass | Idle tab on Gemini 2.5 Flash-Lite, then disconnect: "GPT-6.1 Sol" "High". The only toast is the disconnect one. |
| MODEL-12 | pass | Mock gateway mid-reply, gateway deleted: the tab stayed on "slow-direct" until all twelve words arrived, then moved to "GPT-6.1 Sol". |
| MODEL-13 | pass | No provider: the status bar shows "GPT-6.1 Sol"; "hi" opens "Connect OpenAI (API)". |
| MODEL-14 | pass | Switching to a 16k gateway model moved the meter from 1% to 21% at once. |
| MODEL-15 | pass (#795) | Standalone, with a browser-stored Anthropic login (`oauth.anthropic`) expiring 140 s later and the token endpoint stubbed. Turn 1, outside the 2-minute margin, sent `sk-ant-oat01-OLD`. Turn 2, inside it, called the token endpoint first, then sent `sk-ant-oat01-REFRESHED`, and the new grant was stored. Desktop part not rerun (see below). |
| MODEL-16 | pass (#795) | With stubbed OpenAI (API) and ChatGPT credentials, the copies now differ: "GPT-6.1 Sol · ChatGPT" / "GPT-6.1 Sol · API" (P-89). The picker tags them "openai-codex" (500k) and "openai" (272k). |
| MODEL-17 | pass | "Max context: 16,384 tokens · Max output: 4,096 tokens". |
| MODEL-18 | pass (#795) | Product call reversed: with both connected, a fresh taskpane starts on "GPT-6.1 Sol · ChatGPT" (openai-codex). |
| MODEL-19 | pass | After the disconnect swap, the saved session still held `google/gemini-2.5-flash-lite`. |
| MODEL-20 | pass (#795) | Suspected bug fixed, see GATEWAY-14. |
| MODEL-21 | pass (#795) | Only an Anthropic login at startup: "Claude Opus 5.5" within 1 s and for the next 25 s; "hi" opened no connection dialog. |
| MODEL-22 | pass (#795) | After the disconnect and a reload, the tab is on "GPT-6.1 Sol" and "hi" opens no dialog. Same on the web (the saved Gemini tab came back on GPT-6.1 Sol). |
| LOGIN-01 | pass | Overlay at about 1 s (13 s on `main`). |
| LOGIN-02 | pass | Language buttons, title, subtitle, intro, rows, gateway link, proxy link, in that order. |
| LOGIN-03 | pass (#793) | Focus now starts in the welcome card, so **one** Escape closes the overlay; focus goes to the composer; nothing is stored. |
| LOGIN-04 | pass | A click outside the card closes it. |
| LOGIN-05 | pass | "Enabled" is ticked. |
| LOGIN-06 | pass | Disconnecting the only key showed no welcome overlay; it returned after a reload. |
| LOGIN-07 | pass | `/__pi-auth` held 18 s: welcome at about 7 s, closed by itself at 18 s; the tab moved to "Claude Opus 5.5". |
| LOGIN-08 | pass | "Login with Anthropic", "or enter API key", field, "Save"; focus in the key field. |
| LOGIN-09 | pass | "sk-ant-api… or sk-ant-oat…", "Enter API key". |
| LOGIN-10 | pass | Nothing happens; no error. |
| LOGIN-11 | pass | The OAuth-code message; the row stays open. |
| LOGIN-12 | pass (#795) | Suspected bug fixed: "Checking key…" lists models, and a 401 is rejected (see MODEL-03). |
| LOGIN-13 | pass | "That looks like an OAuth redirect URL/code. Use Google API key auth here, …". |
| LOGIN-14 | pass | `Bearer AIzaFAKE…` saved and connected. New with #795: a second toast, "Saved Google Gemini (API) key. It couldn't be checked now; Pi will use it on the next request.", because Google answers a bad key with 400. Stored value not inspected. |
| LOGIN-15 | pass | With the key check stubbed to 200: "OpenAI (API) connected — try "Explain this workbook""; overlay closed; focus in the composer. |
| LOGIN-16 | pass | "Google Gemini (API) connected"; "Model providers" stays open. |
| LOGIN-17 | pass | See MODEL-06. |
| LOGIN-18 | pass | The claude.ai authorize tab opened, the "Login with Anthropic" dialog has every part, and the row reads "Waiting for authorization code…". |
| LOGIN-19 | pass | "Waiting for the browser callback via the local proxy…". |
| LOGIN-20 | pass | The first Escape leaves the URL field; the second closes the prompt. The row is back to "Login with Anthropic" with no error; the claude.ai tab stays open. |
| LOGIN-21 | pass (#799) | "Login with GitHub Copilot" asks for an Enterprise domain (blank for github.com). Continue then shows "Finish sign-in in your browser", with a code and the verification URL. Escape closed it and reset the row. Note: the device-code request went through the dev server's `/oauth-proxy` to real GitHub, so GitHub issued a code that was never used. |
| LOGIN-22 | pass | Production build: the gate text and "Copied!". Retry without a helper keeps the gate; Retry with one goes on to the login prompt and the claude.ai tab. Cancel and Escape close it. The copy is now platform-neutral ("on this computer", "in a terminal", #795). |
| LOGIN-23 | pass (#793) | "Local HTTPS proxy", "Enabled", URL, "Save", hint, guide link. "Hide proxy settings" is gone: #793 removed the toggle. |
| LOGIN-24 | pass (#795) | Suspected bug fixed: `http://example.com:3003` is refused with "Proxy URL not saved: Proxy URL is HTTP …". |
| LOGIN-25 | pass (#793) | The gateway page opens with Back. Escape now closes Settings, because the gateway page is where the shell opened (#793's rule); it no longer goes to the root first. |
| LOGIN-26 | pass | "中文" reloaded into Chinese (`<html lang="zh-CN">`) with the welcome overlay; "English" reloaded back (`lang="en"`). |
| LOGIN-27 | pass | "Connect OpenAI (API) / This model needs a login or API key to continue."; "hi" in the transcript; Stop shown. |
| LOGIN-28 | pass | Focus in "Enter API key". |
| LOGIN-29 | pass (#795) | Suspected bug fixed: connecting in the dialog toasts "OpenAI (API) connected — try "Explain this workbook"", and the turn continues (401 for the made-up key). |
| LOGIN-30 | pass | See MODEL-13. |
| LOGIN-31 | pass (#795/#789) | The banner is now "API key required for openai." It uses the provider id, not the label "OpenAI (API)". The transcript still reads "Error: Provider is not configured: openai". |
| LOGIN-32 | pass | Reproduced by aborting Anthropic's token endpoint after a valid `code#state`: "Login couldn't connect — this provider needs a helper running on this computer. … Step-by-step guide →". (Copilot no longer fails this way.) |
| LOGIN-33 | pass | 3.17 s. |
| LOGIN-34 | pass (#793) | Moot: the "Hide proxy settings" string no longer exists. In Chinese the panel link reads "登录遇到问题？配置本地代理". |
| LOGIN-35 | pass (#795) | Fixed: "Helper not detected yet — make sure it says “Proxy listening” in a terminal, then try again." renders as text, with no entities. |
| PICKER-01 | pass | Picker open; focus in "Search models…". |
| PICKER-02 | pass | "gpt-6.1-sol ✓ openai-codex 500k ctx · 128k out". |
| PICKER-03 | pass | Current, then gpt-6.1-sol, gemini-3.1-pro-preview, then gateways and the rest. (gpt-5.3-codex-spark is not in this registry.) |
| PICKER-04 | pass | "luna": two rows, the first highlighted. "gemflash" (Gemini connected): gemini flash rows. |
| PICKER-05 | pass | "No matching models". |
| PICKER-06 | pass (#793) | Opened from the status bar: open after one Escape, closed after two, and focus returns to the opener (the model control). Opened with `/model`, focus returns to the composer. |
| PICKER-07 | pass | Down stops at the last row; Enter chose the second. |
| PICKER-08 | pass (#795) | Fixed: with the pointer resting where row 3 opens, Enter kept the current model. The picker starts in keyboard mode. |
| PICKER-09 | pass | Closes; no toast; no change. |
| PICKER-10 | pass | "GPT-6.1 Sol"; no toast; same tab. |
| PICKER-11 | pass | Closed; focus back on the opener. |
| PICKER-12 | pass (#795) | `/default-models` is removed: typing `/default` offers only `/model` and `/yolo`. |
| PICKER-13 | pass (#795) | See MODEL-08. |
| PICKER-14 | pass | "Chat 1 (gpt-6.1-sol)" opened and active; "Opened Chat 1 (gpt-6.1-sol) in a new tab". |
| PICKER-15 | pass (#795) | Read-only busy picker, see MODEL-09. |
| PICKER-16 | pass (#795) | `/model` is now allowed while busy and opens the read-only picker (no "Can't run /model while Pi is busy"); the draft is not kept as text. |
| PICKER-17 | pass | Full list, empty search. |
| PICKER-18 | pass | Gemini rows tagged "google". |
| PICKER-19 | pass (#795) | Suspected bug fixed: after deleting the stored Gemini key, the picker checks "gpt-6.1-sol", the model the refresh swapped in. |
| PICKER-20 | pass | "Chat 1 (gpt-6.1-sol) (gpt-6-sol)". |
| PICKER-21 | pass | Standalone: Ctrl+P and Cmd+P open nothing and are not prevented. Desktop not rerun. |
| SETTINGS-01 | pass | Five groups; "1 connected", "None", Proxy "On" (once enabled; a fresh profile starts "Off"), Plugins "1 installed". |
| SETTINGS-02 | pass | Focus on `.pi-set-shell__body`. |
| SETTINGS-03 | pass | Proxy → root → closed; focus in the composer. |
| SETTINGS-04 | pass | The utilities menu opens Settings while the mock gateway streams. |
| SETTINGS-05 | pass (#795) | `/settings` and `/login` now open while busy (no refusal toast); the turn continued. |
| SETTINGS-06 | pass | Back shown. |
| SETTINGS-07 | pass | × closes from "Model providers"; focus in the composer. |
| SETTINGS-08 | pass | "Discard changes?" on Escape and on ×; Keep editing stays; Discard closes. |
| SETTINGS-09 | pass (#795) | Missing guard added: Back or Escape with a half-filled gateway form asks "Discard gateway changes?" (Keep editing / Discard). Discard leaves the form empty. |
| SETTINGS-10 | pass | Both toasts. |
| SETTINGS-11 | pass | "Confirm mode." / "Auto mode."; the status bar follows. |
| SETTINGS-12 | pass (#796) | The toast is now in the language being left ("Language changed. Reloading..." going to Chinese, "语言已更改，正在重新加载…" going back to English); reload at 1.05 s; the UI switched both ways. |
| SETTINGS-13 | pass | "1 connected" and gateway row "4" (four gateways). |
| SETTINGS-14 | pass (#793) | Changed: F2 no longer reaches the composer behind Settings (focus stays in the shell; 30 Tabs stay inside it), so `/settings` cannot be typed while Settings is open. |
| SETTINGS-15 | pass (#790) | Changed: Save in Rules (opened from the root page) shows "Rules saved" and closes Settings; it no longer returns to the root. The row reads "On: All my files". Rule cleared afterwards. Desktop not rerun. |
| SETTINGS-16 | pass (#795) | "Proxy" now has a subtitle: "Some providers and logins can't be reached straight from Excel. The proxy relays those requests." |
| GATEWAY-01 | pass | All fields and placeholders. |
| GATEWAY-02 | pass | All seven messages, class `pi-overlay-text-warning`, form kept. |
| GATEWAY-03 | pass | "Direct connections need an https:// endpoint. …". |
| GATEWAY-04 | pass | "Endpoint: https://llm.example.com/v1". |
| GATEWAY-05 | pass | "Max output: 16,384 tokens". |
| GATEWAY-06 | pass | "Gateway · localhost:9090". |
| GATEWAY-07 | pass | "Gateway · warehouse (2)"; the toast still names it "warehouse". |
| GATEWAY-08 | pass | All six lines. |
| GATEWAY-09 | pass | Rows tagged "Gateway · {name}". |
| GATEWAY-10 | pass | One row per gateway. |
| GATEWAY-11 | pass | Every field prefilled (key in a password field); focus in Name; Cancel / Update gateway. Cancel empties the form. |
| GATEWAY-12 | pass | "Updated gateway Depot." |
| GATEWAY-13 | pass | "Delete gateway "Slow"? This removes its stored API key from this add-in." Escape keeps it; Delete gives "Deleted gateway Slow." |
| GATEWAY-14 | pass (#795) | Suspected bug fixed: after renaming "Warehouse" to "Depot", the idle tab stays on "gpt-oss-120b", checked under "Gateway · Depot". |
| GATEWAY-15 | pass | Deleting "Depot" moved its idle tab to "GPT-6.1 Sol". |
| GATEWAY-16 | pass | Transcript "Error: Connection error."; banner "Couldn’t get a response: Connection error." (the banner prefix changed with #789/#793). |
| GATEWAY-17 | pass (#795) | Now guarded, see SETTINGS-09. |
| GATEWAY-18 | pass | "Couldn’t get a response: 403 blocked_target_loopback: …". |
| GATEWAY-19 | pass | Enter in Model: no toast, nothing saved. |
| PROXY-01 | pass | "✓ Proxy connected at https://localhost:3003". |
| PROXY-02 | pass | "Proxy enabled." / "Proxy disabled."; status "Proxy disabled." |
| PROXY-03 | pass | "Proxy URL saved."; field without the slash. |
| PROXY-04 | pass | Status line and the "Proxy URL not saved: …" toast. |
| PROXY-05 | pass | "On" and "Off" both read. |
| PROXY-06 | pass (#793) | The banner now sits in the notice area and appeared at once: "The proxy is turned on but isn’t running, so some providers and web tools won’t work." / "Show how to start it". |
| PROXY-07 | **fail** | #793 should show the banner only while the proxy is enabled. A banner already on screen stays after the switch is turned off (70 s+, still saying the proxy "is turned on"); it goes only after a reload. See below. |
| PROXY-08 | pass (#793) | "Show how to start it" expands to "Hide steps", with the command, the copy button, "Open Terminal · paste · press Enter · keep open" and the install link. |
| PROXY-09 | pass | The banner disappeared as soon as the URL was set back to 3003 (the page re-check, #795). |
| PROXY-10 | **fail** | On Excel on the web the copy button never turns into a check, and nothing is copied: the command text is selected instead. The frame denies `clipboard-write`. The code is unchanged from `main`. See below. |
| PROXY-11 | pass (#795) | Suspected bug fixed: right after the change, "⚠ Proxy is enabled at https://localhost:3999 but is not currently reachable." |
| PROXY-12 | pass (#795) | Fixed in reverse too: after setting 3003 and turning the switch on, "✓ Proxy connected at https://localhost:3003" within 0.6 s. |
| PROXY-13 | pass | See LOGIN-05. |
| PROXY-14 | pass (#795) | With the switch on, saving `https://proxy.example.com` now toasts "Security warning: proxy URL is not localhost — it can see your tokens and prompts." The startup toast after a reload remains. Desktop not rerun. |
| PROXY-15 | pass | See LOGIN-22. |
| PROXY-16 | pass | Standalone, with a stubbed ChatGPT login and `chatgpt.com` blocked. Proxy off: "GPT-5.6 Luna currently requires the latest …". Proxy on without the bridge header: "GPT-5.6 Luna requires a newer …". |
| PROXY-17 | pass | `https://localhost:3003`. |
| PROXY-18 | pass | See MODEL-05. |
| PROXY-19 | pass | The next request went straight to chatgpt.com. |
| PROXY-20 | pass | See GATEWAY-16 (the corrected document). |
| PROXY-21 | pass | The pending edit saved on the first Escape; the root page on the second. |

## Regressions and new bugs

### 1. PROXY-07: turning the proxy off does not hide a proxy banner already showing (#793)

- **Repro (web):**
  1. Settings → Proxy, switch on, URL `https://localhost:3999` (nothing listening).
  2. Close Settings. The notice area shows "The proxy is turned on but isn’t running, …".
  3. Settings → Proxy, switch off, close Settings.
- **Expected (#793):** the banner shows only while the proxy is enabled.
- **Observed:** the banner stays, still saying the proxy "is turned on", for 70 s or more; it goes away only after a taskpane reload.
- **Cause:** `proxyBanner.update()` runs only on `pi:proxy-state-changed`, which `checkProxyOnce` dispatches only when the reachability state flips (`src/taskpane/proxy-status.ts:103-107`). Changing `proxy.enabled` never calls it (`src/taskpane/init.ts:2266-2279`, `src/ui/proxy-banner.ts:84-110`).
- **Owner:** #793. It is not specific to the combination. Sent to the #793 thread.

### 2. PROXY-10: the banner's "Copy command" does not copy on Excel on the web (#793; not a regression from the fixes)

- **Repro (web):** with the banner from bug 1 showing, click "Show how to start it", then the copy icon.
- **Observed:** the label stays "Copy command" and the command text is selected; nothing reaches the clipboard. In the add-in frame `document.featurePolicy.allowsFeature('clipboard-write')` is false, and `src/ui/command-copy.ts:18-25` tries only `navigator.clipboard.writeText`. `main` uses the same `writeText`, so the earlier pass on `main` was probably not on this frame policy.
- **Fix hint:** #795's `copyText()` (`src/ui/copy-text.ts`) has an `execCommand` fallback that works in this frame (#795's evidence).
- **Owner:** #793 owns the banner. Sent to the #793 thread as low priority.

No bug found here appears only in the combination, so the integration thread was not messaged.

## Desktop items to hand to the Mac

No item in this file is desktop-only, but these desktop halves were not rerun:

- **MODEL-15:** the earlier Mac pass loaded the expiring login through pi's `auth.json` (`/__pi-auth`). With #795, refresh-before-request applies only to logins stored by the taskpane's own OAuth flows (`oauth.<provider>` in settings). Logins imported from `/__pi-auth` are refreshed only when the pane opens (`src/auth/restore.ts:100-112`). So a repeat of the same method will still send the expired token. Test with a login made in the taskpane (stub the token endpoint), or treat the `/__pi-auth` path as dev-only.
- **PICKER-21:** Cmd+P in the composer (Excel's Print dialog expected; nothing in Pi).
- **PROXY-14:** the remote-URL warning at pane open, and now also on save (#795).
- **SETTINGS-15:** saving a workbook rule now closes Settings (#790).
- **LOGIN-22 / PROXY-15:** the proxy gate in a desktop production build, now with the "on this computer" / "in a terminal" copy.
- From "Not checkable by hand": where the login window opens on Excel desktop (system browser), and the device-code dialog's polling after Escape on desktop (LOGIN-21 now reaches the dialog).

## Notes on the run

- During one standalone run, a catch-all Playwright route let `/__pi-auth` through to the dev server, so that page got the shared ChatGPT login. Two tiny turns ran on it ("hi", "hi again", GPT-6.1 Sol). The login was not expired, so nothing was refreshed. Every later run stubbed `/__pi-auth` explicitly.
- Test data (four gateways, a workbook rule, a fake Gemini key) was removed. The web session ends with the proxy on at `https://localhost:3003`, fork off and Auto mode.
