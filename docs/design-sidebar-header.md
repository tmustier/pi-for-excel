# Design: Sidebar Header Chrome

> **Related issues:** [#12](https://github.com/tmustier/pi-for-excel/issues/12)
> **Status:** Implemented (V1)
> **Last updated:** 2026-02-11

## Current state

The top of the sidebar currently renders three stacked elements:

```
┌──────────────────────────────────────┐  ← Office chrome (not ours)
│  Pi for Excel                    ⓘ  │     title bar + native info button
├──────────────────────────────────────┤
│  Pi                              ⋯  │  ← Our header (from branch work)
├──────────────────────────────────────┤
│  [Chat 1]                        +  │  ← Our tab strip
├──────────────────────────────────────┤
│  WORKBOOK                           │  ← Our workbook label
│  GIP Digital Case Study Maria …     │
├──────────────────────────────────────┤
│                                      │
│  (messages)                          │
```

**Problems:**

1. **"Pi" header is redundant** — Office chrome already shows "Pi for Excel" in the taskpane title bar. Our header wastes ~30px of vertical space just to repeat the name.
2. **Tab titles from first message are noisy** — when auto-generated from the user's first prompt, tabs become unreadable ("Analyze the revenue data on…" truncated to 12 chars).
3. **Workbook label row is low-value** — the user already knows which file they have open. A full row for it is wasteful.
4. **Settings / Instructions are hard to find** — reachable only via `/settings`, `/instructions`, or a tiny status-bar badge. No persistent, obvious entry point.
5. **Resume / session management is hidden** — only via `/resume` slash command.

---

## Goals

1. Minimize vertical space consumed by header chrome.
2. Provide a single, always-visible entry point for utilities (settings, instructions, resume, shortcuts).
3. Give tabs stable, predictable names by default.
4. Remove or compress the workbook label.
5. Keep the design clean in a ~350px sidebar.

---

## Constraints

- **Office taskpane chrome is fixed.** We cannot remove or modify the "Pi for Excel" title bar or its ⓘ button. Every pixel of our own header stacks *below* it.
- **Tab strip is needed** for multi-session (#31 Phase 1).
- **Footer status bar already holds** context %, model picker, thinking level, instructions badge, lock state. It should not get more crowded.

---

## Design

### `⋯` utilities menu in the tab strip (single row)

No separate header bar. The `⋯` menu button sits at the end of the tab strip row, after `+`.

```
┌──────────────────────────────────────┐  ← Office chrome (fixed)
│  Pi for Excel                    ⓘ  │
├──────────────────────────────────────┤
│  [Chat 1] [Chat 2]      [+]  [⋯]   │  ← single combined row
├──────────────────────────────────────┤
│                                      │
│  (messages / empty state)            │
```

**Rationale:** Office already brands the pane. A second header saying "Pi" adds nothing. The `⋯` button is compact (20×20px) and sits naturally at the row's right edge, visually distinct from `+`.

### Tab strip layout

```
┌─────────────────────────────────────────────────┐
│ [Agent 1]  [Agent 2·]      [+]    [⋯]          │
│  ↑ active    ↑ busy dot     ↑ new   ↑ menu     │
└─────────────────────────────────────────────────┘
```

- **Tabs scroll** horizontally when > 3 tabs.
- **Active tab** has a tinted border/background (current style).
- **Busy dot** (green) on tabs with active streaming.
- **Lock indicator** ("lock…") on tabs waiting for workbook write lock.
- **Close ×** on each tab when > 1 tab exists.
- **`+` button**: always visible, creates a new blank tab.
- **`⋯` button**: always visible, opens utilities menu.

### Tab naming

| Situation | Tab title |
|-----------|-----------|
| New tab, no explicit name | `Agent 1`, `Agent 2`, … (monotonic counter per session) |
| User runs `/name My Analysis` | `My Analysis` |
| Resumed session with existing explicit name | Restored name |
| Resumed session with no explicit name | `Agent N` (next counter value) |

**Key rule:** the auto-generated session title (derived from first user message) is stored in session metadata for the resume picker and history search. It is *never* used as the tab label. Only `/name` changes the tab label.

This keeps tabs clean and scannable. Session titles remain useful in the resume overlay.

### `⋯` Utilities menu

The menu provides a single discoverable entry point for things that don't need to be always-visible.

```
┌────────────────────────┐
│  Instructions…         │
│  Settings…             │
│  ─────────────────     │
│  Resume session…       │
│  Keyboard shortcuts    │
└────────────────────────┘
```

**Item ordering rationale:**
- Instructions + Settings first: most common non-chat actions (tier 1).
- Resume: session management.
- Shortcuts last: reference, not frequent.

Everything else stays as `/` slash commands (export, compact, copy, name, debug, extensions, experimental).

### Workbook label

**Remove the dedicated row.** The workbook name is already known to the user (it's in Excel's title bar) and is injected into the agent's context automatically.

If needed in the future (e.g., cross-workbook resume disambiguation), surface it as:
- A line in the `⋯` menu: `Workbook: <name>` (non-interactive, muted text).
- Or a tooltip on the tab strip area.

Do not give it its own row.

### Lock notice

When a tab is waiting for the workbook write lock, the tab itself shows the "lock…" badge. The separate `pi-lock-notice` banner row can be removed; the tab badge plus the status bar are sufficient.

---

## Interaction details

### `⋯` menu behavior
- Click `⋯` → toggle menu open/closed.
- Click outside menu → close.
- Escape → close.
- Menu items dispatch their action and close the menu.

### `+` button behavior
- **Click**: new blank tab (instant, no menu).
- **Right-click** (optional, Phase 2): context menu with `New tab`, `Resume session…`, recent sessions.

### Tab close behavior
- Click `×` on a tab → close tab.
- If actively streaming → confirm dialog ("Stop and close" / "Cancel").
- If holding write lock → `×` disabled until write completes.
- On close: force-save session, push to recently-closed stack, show undo toast.

### Keyboard shortcuts
- `Cmd/Ctrl+Shift+T`: reopen last closed tab.
- Other existing shortcuts unchanged.

---

## Visual spec (approximate)

### Dimensions
- Tab strip row height: ~32px (same as current).
- `⋯` button: 24×24px, same border-radius (6px) as `+`.
- No separate header row → saves ~30px.
- No workbook label row → saves ~26px.
- **Net vertical space saved: ~56px** returned to messages.

### Spacing
- Tab strip: `padding: 7px 8px 6px` (unchanged).
- `+` and `⋯` grouped at right edge with `gap: 4px`.
- Menu dropdown: `min-width: 190px`, anchored to `⋯` button top-right, offset `4px` below.

### Colors / treatment
- `⋯` button: same muted style as `+` (transparent bg, muted-foreground, hover highlight).
- Menu: glass background (`oklch(1 0 0 / 0.92)`, blur, shadow) matching existing overlay style.
- Menu items: `12.5px` font-sans, full-width hover highlight.

---

## Technical changes

### `src/ui/pi-sidebar.ts`

- Remove `_renderHeader()` method entirely.
- Remove `_headerMenuOpen` state.
- Move `⋯` button and menu into `_renderSessionTabs()`, after the `+` button.
- Remove `_renderWorkbookLabel()`.
- Remove `workbookLabel` property.
- Remove `lockNotice` property and its banner rendering (tab badge is sufficient).

### `src/ui/theme/components.css`

- Remove `.pi-header`, `.pi-header__title`, `.pi-header__menu-anchor`, `.pi-header__menu-btn` styles.
- Remove `.pi-workbook-label`, `.pi-workbook-label__hint`, `.pi-workbook-label__name` styles.
- Remove `.pi-lock-notice` styles.
- Keep `.pi-header-menu` / `.pi-header-menu__item` / `.pi-header-menu__divider` styles (rename to `.pi-utilities-menu` etc. for clarity since they're no longer in a header).
- Add `.pi-tab-strip__actions` container for `+` and `⋯` grouping.

### `src/ui/header.ts`

- Delete this file (already a stub; was kept for import compatibility).
- Remove import from `src/ui/index.ts`.
- Remove `headerStyles` injection from `src/taskpane/bootstrap.ts`.

### `src/taskpane/init.ts`

- Stop setting `sidebar.workbookLabel`.
- Wire `⋯` menu callbacks: `onOpenInstructions`, `onOpenSettings`, `onOpenResumePicker`, `onReopenLastClosed`, `onOpenShortcuts`.
- Remove `refreshWorkbookState` calls that only updated `sidebar.workbookLabel` (keep the ones that refresh system prompt).

### `src/taskpane/session-runtime-manager.ts`

- `snapshotTabs()`: use `Agent N` counter naming instead of session title. Only use `persistence.getSessionTitle()` when an explicit `/name` has been set (requires `hasExplicitTitle` flag on `SessionPersistenceController`).

### `src/taskpane/sessions.ts`

- Add `hasExplicitTitle()` method to `SessionPersistenceController`.
- `renameSession()` sets `explicitTitle = true`.
- `startNewSession()` resets `explicitTitle = false`.
- `applyLoadedSession()` restores `explicitTitle` from session data (requires minor metadata addition).

### `src/taskpane.html`

- Remove `<div id="header-root"></div>` (unused; our header renders inside `#app` via `pi-sidebar`).

---

## Acceptance criteria

1. No separate header row renders — tab strip is the first element below Office chrome.
2. No workbook label row renders.
3. No lock-notice banner row renders (tab badge only).
4. `⋯` button is always visible in the tab strip, after `+`.
5. `⋯` menu opens with: Instructions, Settings, Resume session, Reopen last closed (conditional), Keyboard shortcuts.
6. Tab titles are `Agent 1`, `Agent 2`, … by default; only change when user runs `/name`.
7. Net vertical space saved: ≥50px compared to current layout.

---

## Open questions

1. Should `⋯` menu include a non-interactive workbook name line for orientation, or is Excel's own title bar sufficient?
2. Should `+` get a right-click context menu (Phase 2) or keep it single-action?
3. Should the `Agent N` counter reset per page load or be globally monotonic (per workbook)?
