# Thinking

Behaviour as of [PR #803](https://github.com/tmustier/pi-for-excel/pull/803), implementing CL-09 / CC-07 and the thinking-only part of CL-05. This reverses #793's composer Shift+Tab decision.

## Summary

The thinking level trades speed for reasoning depth. Each tab keeps its own level, shown beside the model in the status bar and saved with the session. Click that control to choose a level, or press **Ctrl+Alt+T** in the composer (**Ctrl+Option+T** on Mac) to cycle. **Shift+Tab always moves focus backwards**; Pi does not use it for thinking anywhere.

## Choosing a level

The “Thinking level” popover lists the model's supported levels with the current one checked:

| Level | Label | Hint |
| --- | --- | --- |
| off | Off | Fastest — no reasoning step |
| minimal | Min | Quick — light reasoning |
| low | Low | Fast — moderate reasoning |
| medium | Medium | Balanced — solid reasoning |
| high | High | Slow — thorough reasoning |
| xhigh | Extra high | Very slow — extra-high reasoning |
| max | Max | Slowest — maximum reasoning |

The registry decides which levels are available. Non-reasoning models offer only Off; some reasoning models do not offer Off.

### Ends at once

Choosing the already selected level just closes the popover. Escape or clicking outside closes it without changing anything. Picking a row changes the control label without a toast.

The dedicated shortcut advances through the supported levels, wrapping to the first, and shows the colour flash plus “Thinking: {level} (next turn) · {shortcut}”. It works only with focus in the composer and no modal open. It is not a global shortcut for dialogs, Settings, search fields or Excel's grid. The tooltip and Shortcuts page show the same platform-specific chord.

### The level is set

Changing thinking is allowed while Pi works, but the running turn retains its starting level for every model response, including after a steer or follow-up. The next turn uses the new level. Cycling preserves the actual execution-mode and lock badges. A level unavailable after a model change is clamped to one that model supports.

## While the model reasons

Readable reasoning appears in a button-headed thinking block labelled “Thinking…” with a spinner. That label ends when reasoning ends, even if text or tools continue. Clicking the header expands the reasoning; Enter and Space do the same, and the button exposes `aria-expanded`.

Initial disclosure follows the remembered taskpane-wide details mode. An individual toggle overrides that block until the next global Ctrl/Cmd+O toggle. Disclosure survives the streaming-to-transcript handoff and tab switches.

The working line uses playful phrases only during readable thinking. Writing, tools, compaction, retry waits and stopping instead show their real activity. The working hint stays “Press Esc to stop”; it no longer rotates through thinking shortcuts.

Some providers return only a reasoning summary or no readable reasoning. No empty block is invented for them.

## The block after the model response

A block whose start was observed reads “Thought for {duration}”; otherwise it reads “Thought”. The timing is shared transcript state, not a measurement invented when an old session is restored. Multiple responses in a turn can each have their own thinking block. Stop or failure keeps whatever reasoning was already returned.

## Interactions with other systems

The level is saved with its session when persistence runs. A higher level uses more time and tokens; providers may map it to their nearest supported setting. Compaction independently asks for high reasoning.

## Host, language and verification

All UI labels, hints, timing and toasts are translated. Mac uses Ctrl+Option+T, not Cmd; Windows/Linux use Ctrl+Alt+T. Shift+Tab follows native focus navigation on all hosts, and modal focus stays in the modal.

## Open questions and verification

PR #803's browser test verifies that Shift+Tab moves focus without changing the level, Ctrl+Alt+T cycles only in the composer, and unrelated text fields do not trigger it. Excel-web acceptance with GPT-6 Luna also exercised the dedicated chord and recorded keyboard behaviour. Desktop key delivery was not repeated for this PR.
