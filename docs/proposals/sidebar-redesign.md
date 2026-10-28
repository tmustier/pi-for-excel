# Sidebar redesign

**Status:** Proposal, direction agreed with Thomas on 2026-10-28. Not implemented.
**Last reviewed:** 2026-10-28

This is the design specification for a new look and structure for the Pi for
Excel task pane. It covers every surface the pane shows: the shell, the
transcript, the tool activity group, menus, full-pane sheets, the first-run
screen and the ribbon icon. When the first implementation slice lands, the
"Tokens", "Surfaces" and "Components" sections below move to a root
`DESIGN.md` and this file becomes the record of why.

Mockups are static HTML under [`sidebar-redesign/`](./sidebar-redesign/).
Open them in a browser; each accepts `?w=320` (pane width), `?only=2,5`
(which states) and `?zoom=2`. On macOS use the `file://` form so `open` keeps
the query string:
`open "file://$PWD/docs/proposals/sidebar-redesign/02-turn-group.html?w=320&zoom=2"`.
Always review at 320 and 360, zoomed.

## 1. Why

The current theme was built piece by piece and reads as a chat app dropped
into Excel. Specific problems, from screenshots of the shipped build:

- **Colour is decoration.** Green tints on tool cards, thinking blocks, the
  focus ring, the mode chip and the send button. When everything is green,
  green no longer means "a change happened". The dark theme is blue-tinted.
- **Everything is a card.** 16px pills for tool cards, user messages and the
  input; 20px floating overlays over a blurred pane with nested cards inside.
  In a 360px pane each border eats width and overlays lose ~40px to margins.
- **Glass and texture.** Spreadsheet-grid background, specular highlights,
  inset shadows. None of it carries information.
- **Hierarchy by weight.** Bold plus green plus a pill, rather than size,
  weight and space.
- **The empty state leads with a slogan**, not with what Pi can see.
- **Tool calls are a log of cards**, not a live component: no group state,
  errors and approvals are separate cards or an overlay, the group never
  collapses itself.
- **The user bubble is right-aligned** in a pane too narrow for alignment to
  mean anything.

## 2. Principles

1. **The pane is an instrument next to the sheet, not a chat app.** Operate
   mode: scanability, consistency and native expectations outrank expression.
   Brand lives in precision, not in a mark.
2. **Colour means something.** One accent, a bottle green, with one rule: it
   marks cells Pi touched (done ticks, changed values, the receipt, a cell
   reference under the pointer, the live spinner) and the one action that asks
   Pi to touch them (Send). Amber for waiting on the user. Red for failed.
   Nothing else is coloured, and primary buttons stay ink.
3. **Rows, not cards.** Lists are hairline-separated rows at full width. A
   bordered surface exists only when it contains a unit the user acts on as a
   whole: the composer, a popover, the turn group, a code block.
4. **Drawn icons only.** One icon set, one stroke weight. No Unicode glyphs or
   emoji standing in for icons, no `·` separators standing in for structure.
5. **Mono is for data.** Cell refs, values, formulas, commands, elapsed time.
   Labels are sans.
6. **Controls belong to whoever acts.** If the model performs the action (retry
   a tool, run again), it is not a button. The user's controls are approval,
   undo, review, steer and stop.
7. **Copy is the product's own language**, past tense for what happened,
   present tense for what is happening.
8. **Verify at the sizes it ships at.** 320px is the floor.

## 3. Tokens

### Colour

Warm-tinted neutrals so the pane sits beside a white Excel canvas without
looking grey or blue. Every text colour meets 4.5:1 on its surface.

| Token | Light | Dark | Use |
|---|---|---|---|
| `bg` | `#faf9f7` | `#171614` | pane background |
| `surface` | `#ffffff` | `#1e1d1a` | composer, popovers, groups, code |
| `surface-2` | `#f1efeb` | `#262420` | user message, highlighted row, live row |
| `line` | `#e6e3de` | `#2d2b26` | hairlines |
| `line-strong` | `#d3cfc8` | `#3b3832` | input borders, outlined buttons |
| `fg` | `#1c1b19` | `#ecebe7` | primary text |
| `fg-2` | `#55524d` | `#b3b0a9` | secondary text |
| `fg-3` | `#6b6862` | `#99958d` | tertiary text, icons |
| `accent` | `#155e4b` | `#5fc294` | cells Pi touched; Send |
| `accent-strong` | `#0f4a3a` | `#7ad1a8` | hover on green fills |
| `amber` | `#9a6b12` | `#e0b24a` | waiting for approval, context warning |
| `amber-soft` | `#fbf3e3` | `#2e2718` | approval row surface |
| `red` | `#b4421f` | `#f08a66` | failed |
| `red-soft` | `#f9ebe5` | `#3a221a` | error banner surface |
| `selection` | `#d9e9e1` | `#244237` | `::selection` |

Removed: all `--glass-*`, `--pi-grid-line`, `--pill-green-*`, `--pill-red-*`,
the alpha ladders, the green-tinted tray.

### Type

- **Sans:** IBM Plex Sans 400/500/600. **Mono:** IBM Plex Mono 400/500.
  Chosen because the pair is a true sibling (refs don't look pasted in), the
  numerals are tabular, and it renders well at 12–13px. Geist is the
  alternative if Thomas prefers a cooler voice; it was the v1 choice. Confirm
  before the token slice.
- `font-variant-numeric: tabular-nums` on the root.
- Scale: 11 (table headers, keyboard hints), 11.5 (chip text, secondary meta),
  12 (mono refs, subtitles), 12.5 (rows, activity, secondary prose), 13 (UI
  default), 13.5 (assistant prose), 20–22 (one headline per screen: workbook
  name on the start screen, product name on first run). Nothing below 11px.
- Weights: 400 body, 500 titles and buttons, 600 only for `<strong>` in prose.
- No eyebrows or kicker labels above headings. Section headers in sheets are
  sentence-case 12px `fg-3` text, not uppercase.
- `text-wrap: balance` on the headline, `pretty` on prose.

### Spacing and radii

- Pane padding 14px. Row padding 10px 14px. Row min-height 44px.
- Radii: 10px composer and popovers; 8px turn group, user message, search
  input, toast; 6px buttons, code blocks, icon buttons, chips in the composer;
  5px chips in text; 3px inline code.
- Hairlines are 1px `line`. No coloured left borders.

### Elevation

Declared once per element, border or shadow:

- Border only: composer, turn group, code block, chips, outlined buttons.
- Shadow only: popovers (`0 6px 20px -6px rgba(0,0,0,.18)`), toasts (none, the
  dark surface is enough).
- Sheets are flat: they replace the pane, they don't float over it. No blur.

### Motion

One authored moment: the group header spinner. Row entrance is instant. The
sheet slides in from the right over 180ms ease-out; back slides out. Reduced
motion disables both.

## 4. Surfaces

### Shell

```
┌─ header ─────────────────────────────────┐  41px: chat title · + · menu
│  banner (when present)                   │  flat band, hairline below
│  transcript (scrolls)                    │  14px side padding
│                                          │
│  dock                                    │
│    queue tray                            │  rows of queued messages
│    working line                          │  "Filling 'Cash Flow'!D13:L13   Esc to stop"
│    composer                              │  attachments · textarea · chips · send
└──────────────────────────────────────────┘
```

- **No tab strip.** The header shows the current chat's title. `+` starts a
  chat. The menu lists recent chats, "All chats", then Files, Connections,
  Rules, Backups, then Settings and Keyboard shortcuts.
- **No mark in the pane.** Excel's task-pane chrome already says "Pi for
  Excel" above our header. The π leaves the header, the empty state, the
  status bar and the login screen.
- **No status bar.** Model, thinking level, workbook-change mode and send live
  in the composer's footer row. The context chip (`Context 71%`, amber)
  appears only once usage is worth knowing about; the model chip ellipsizes
  to make room. Nothing in the footer row wraps.

### Start screen (empty chat)

Workbook name as the headline (`Q3 Forecast` with `.xlsx` in `fg-3`), then a
two-column fact list in sans: Sheets, Selection (mono), Memory. A hairline,
then the four starters as rows with a drawn icon and a shortcut. No slogan, no
starter descriptions; the full prompt text still fills the composer on click.

### Transcript

- **User message:** `surface-2` block, 8px radius, full width, left-aligned.
  Attachments and the selection appear as chips inside it.
- **Assistant prose:** plain, 13.5px/1.55. Inline code is mono on `surface-2`.
- **Thinking:** `Thought for 3s` with a chevron in `fg-3`. Expanded text is
  `fg-2` 12.5px indented behind a 1px hairline. "Thinking…" while live.
- **Code and formulas:** `surface` block, 6px radius, header row with the
  language label and an icon-only copy button, 12px mono body.
- **Tables:** hairlines, sans, header in `fg-3`, numbers right-aligned in mono.
- **Answer actions:** a single text button, "Copy answer".
- **Turn group:** section 5.

### Composer

`surface`, 1px `line-strong`, 10px radius, no shadow. Attachment chips sit
above the textarea. Footer row: attach icon, model chip with chevron,
thinking chip, mode chip, optional context chip, send. The open chip gets
`surface-2`. Send is filled `fg` when there is text, `surface-2`/`fg-3` when
empty, and an outlined stop square while streaming. Placeholder reads
"Ask about this workbook", or "Steer, or queue a follow-up" while streaming.

### Banner

A flat band under the header: icon, text, text-link actions, dismiss. Info is
on `surface`; errors on `red-soft` with a red icon. Turn-level retry (provider
429, network) lives here, never in the transcript.

### Toast

Dark rectangle (`fg` on `bg`), 8px radius, bottom of the transcript, one
underlined action on the right.

### Popovers

`surface`, 1px `line-strong`, 8px radius, shadow, 4px inner padding, items 7px
10px with 5px radius, highlighted item on `surface-2`. Optional section label
in `fg-3` 11.5px. Used for: the header menu, slash commands (mono command left,
description right), model, thinking (six levels, one line each, note at the
bottom), workbook changes (two options with one-line descriptions).

### Sheets (full-pane overlays)

Replace the pane. Title row 41px: back (when nested), title, optional text
action, close. Body is sections of rows:

- Section header: sentence case, 12px `fg-3`, 16px above, 6px below.
- Row: optional leading icon, title (500) + optional subtitle (12px `fg-3`,
  ellipsized), trailing value in `fg-3` and/or chevron, or a toggle.
- Search: 34px, `surface`, `line-strong`, 8px radius, at the top of long lists.
- Segmented control for a small number of peer views (Rules).
- Forms end with a right-aligned action row: text Cancel, filled Save.
- Empty states are centred 12.5px text with a 500-weight first line.

Sheets in scope: Settings (a hub linking out to Providers, Rules, Files,
Backups), Model (search, Recent, then grouped by provider, "Providers" text
action), Chats (search, this workbook, other workbooks with the workbook
name as subtitle, "Edit" text action), Rules, Files, Connections, Skills,
Keyboard shortcuts, Backups. This supersedes the floating-card and tab
patterns in [settings-ux-redesign.md](../settings-ux-redesign.md); the
navigational model there (one hub, nested screens with back) stands.

### First run

Product name as a 22px headline, one sentence, then provider rows (Anthropic,
OpenAI, Google Gemini, Other providers, Custom gateway) with subtitles and
chevrons. Footer: "Having trouble signing in?" (the proxy path) and the
language switch, as text links.

## 5. The turn group

One group per turn. It is the record of what the agent did, and it carries its
own state. Mockup: [`02-turn-group.html`](./sidebar-redesign/02-turn-group.html).

```
┌───────────────────────────────────────────────┐
│ (icon stack)  Working                     8s ˄ │  header
├───────────────────────────────────────────────┤
│ ✓  Read Inputs!C4:K9                54 cells ˅ │  row
│ ✓  Filled 'Cash Flow'!D10:L10        9 cells ˅ │
│ ⚠  Could not fill 'Cash Flow'!D12:L12        ˅ │  failed row
│    Inputs!C6 no longer exists, so the formula  │  reason
│    returns #REF!.                              │
│ ◌  Filling 'Cash Flow'!D12:L12 from Inputs!C7  │  live row (surface-2)
├───────────────────────────────────────────────┤
│ Changed 45 cells                 Undo  Review  │  footer
└───────────────────────────────────────────────┘
```

**Header.** A stack of 20px circles, one per tool *kind* used so far (sheet,
Python, web, extension), overlapping by 6px. Then the state, 500 weight:

| State | Title | Right side |
|---|---|---|
| streaming | `Working` with a spinner | elapsed, mono |
| failed step so far | `Working` | red alert icon + `1 failed` |
| waiting | `Waiting for approval` in amber | |
| done | `Worked for 8s` | `1 failed` if any |

Chevron collapses or expands. The group collapses itself when the answer
starts; the footer stays visible when collapsed.

**Rows.** Grid: 16px status icon, sentence, meta, 14px chevron. Icons and
chevron align to the first line, so a wrapped sentence still reads. Refs are
`white-space: nowrap` so they move to the next line whole.

- Sentence: past-tense verb in sans + ref in mono. `Read Inputs!C4:K9`,
  `Filled 'Cash Flow'!D10:L10`, `Ran Python`, `Searched the web for UK CPI`.
  The live row is present tense on `surface-2`.
- Meta: one fact, 11.5px `fg-3`: `54 cells`, `6 results`, `3 s` for Python.
  No per-row timings for workbook operations.
- Status: accent check, red alert, amber hollow circle (pending), spinner.
- A failed row turns the sentence red (`Could not fill …`) and adds a reason
  line in `fg-2`: the error text the model received. **No Retry.** The model
  decides what to do; its recovery is the next row.
- Every row expands. A Fill row expands into the before/after table (cell,
  before, after; formulas as a second line; after-values in accent; "Show 7
  more" as a sans row).

**Approval row.** Amber surface. Sentence states the action; a consequence
line states scope and reversibility ("Removes 13 rows and 78 non-empty cells.
Covered by a backup for this session."). Actions right-aligned: outlined
**Decline**, filled verb (**Delete**). "Allow for the rest of this turn" is a
text link on its own line. Strings come from `approval.*` in `en.json`.

**Footer.** `Changed 45 cells` (500), then text **Undo**, outlined **Review**.
The sheet name is not repeated here; the rows carry it.

## 6. Components

### Buttons

Three weights, one component, 28px tall, 6px radius, 12px/500. Always
right-aligned in an action row, the action being asked of you rightmost.

| Weight | Look | Examples |
|---|---|---|
| primary | filled `fg` on `bg` | Delete, Save, Send |
| secondary | outlined `line-strong` | Decline, Review, Cancel in sheets |
| text | no border, `fg-2` | Undo, Copy answer, Providers, Edit |

Full-width 36px variant for the one primary action on a form screen.

### Chips

- In text (attachments, selection): 5px radius, `surface`, 1px `line`, drawn
  icon in `fg-3`, mono label 11.5px.
- In the composer footer: 26px, 6px radius, 12px/500 `fg-2`, `surface-2` when
  open, amber for the context warning. Never wrap.

### Rows, search, segmented control, toggle

As described under Sheets. Toggle: 30×18, `line-strong` off, `fg` on, white
knob.

### Icons

One set at 24-unit viewBox, 2px stroke (2.5px for the check and close),
rendered at 12–16px. Set so far: plus, menu, close, back, chevrons, check,
alert, info, spinner, pending circle, paperclip, up arrow, stop, document,
range, clock, copy, search, gear, folder, plug, rules, history, backup,
keyboard, sheet, terminal (Python), globe. Lucide covers all of these; the
range icon is custom (a 3×3 grid with the centre filled).

## 7. Brand

- **Inside the pane: nothing.** See Shell.
- **Ribbon and store icon:** "pi as a table", the letter reduced to a header
  rule over two column strokes. White on an accent-green tile for the 16/32/80
  PNGs, because Office ships one icon set for light and dark ribbons. Bare
  glyph anywhere we control the background. Mockup:
  [`04-ribbon-mark.html`](./sidebar-redesign/04-ribbon-mark.html).
- **Wordmark:** "Pi for Excel" set in Plex Sans 500. No glyph.

## 8. What is removed

Tab strip, status bar, π in the pane, spreadsheet-grid background behind the
transcript (an empty sheet of 64×20 cells remains behind the empty state, the
one place the motif appears), glass tokens, green-tinted tool and thinking cards, 16px pill radii, right-aligned
user bubble, floating overlays and their blur, nested cards inside overlays,
card-per-option popovers, starter descriptions, the slogan, the context ring,
per-row tool timings, "Retry" on tool rows, the approval overlay.

## 9. Open decisions

1. **Font.** Plex (this spec) or Geist. Decide before the token slice.
2. **Thinking block depth.** One-line summary with expand. If thinking is read
   often, give it more room.
3. **Settings hub.** Links out rather than inlining providers and rules. One
   more tap for power users, full width for every screen.
4. **Undo weight.** Text in the footer. Make it outlined if it reads as
   hidden.
5. **Consecutive same-verb rows.** Kept as separate rows (honest log). Could
   collapse into `Filled 'Cash Flow'!D10:L15, 5 ranges` with expand.
6. **Welcome warmth.** Deliberately plain. If one surface gets an illustration
   it is this one, and it should be a real asset.
7. **WPS and older WebViews.** Tokens are plain hex; no OKLCH or
   `color-mix`, so `legacy-webview.css` fallbacks shrink rather than grow.

## 10. Implementation plan

Each slice is a PR with its own screenshots at 320 and 360, light and dark,
and the gallery updated. Run the impeccable detector on changed CSS and HTML
in each PR (`npx impeccable install --providers=pi` is not yet in the repo;
add it with slice 1 and run `/impeccable init` to write `PRODUCT.md`).

1. **Tokens and type.** Replace `theme/tokens.css`; swap fonts; remove glass,
   grid and tint tokens; new radius and elevation rules; `::selection`,
   tabular numerals, focus rings. `check:css-theme` keeps every `var()` honest.
2. **Shell.** Header replaces tabs; menu popover; composer absorbs the status
   bar; banner, toast, queue tray, working line; start screen.
3. **Turn group.** New `<turn-group>` replacing grouped tool cards and the
   activity-block prototype; row renderers per tool kind; approval row replaces
   the approval overlay for workbook changes; footer replaces turn-changes
   messages; drill-in reuses the diff table.
4. **Transcript content.** User message, thinking, code, tables, answer
   actions.
5. **Sheets.** Overlay primitives become the sheet pattern; migrate Settings,
   Model, Chats, Rules, Files, Connections, Skills, Shortcuts, Backups.
6. **First run and icon.** Welcome screen; new ribbon PNGs; manifest.

Acceptance for behaviour-changing slices (2, 3, 5) is a real turn in local
Excel via `excel-background-verification`, not the gallery alone.

## 11. Mockup index

| File | States |
|---|---|
| `01-shell-and-start.html` | start screen, streaming shell, done, dark. Its flat activity log predates the turn group; `02` supersedes it |
| `02-turn-group.html` | streaming, error then recovery, waiting for approval, done collapsed, done expanded with drill-in |
| `03-menus-sheets-content.html` | header menu; banner, queue, attachments, toast; slash menu; thinking menu; mode menu; Settings; Model; Chats; Rules; first run; thinking, table, code; dark Settings |
| `04-ribbon-mark.html` | icon candidates at 80/32/16 on light and dark ribbons |

Inspiration, for provenance: Damir Sokolovsky's "unslopping a vibecoded
dashboard" (spacing, colour restraint, icon consistency, typography,
data-first content) and Elia Kuratli's tool-call group (live header state,
past-tense rows, inline error and approval, self-collapse).
