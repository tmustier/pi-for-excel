# Discoverable Hints — Working Indicator + Status Bar

## Goals
Rotate educational hints in the working indicator while the agent is streaming, combining whimsical "Working…" text with feature discovery tips. Also add non-working-state discoverability for status bar items.

## Checklist

### 1. Working indicator: rotating hints
- [x] Replace static "escape to interrupt" with rotating tips array
- [x] Always show "escape to interrupt" first, then random rotation
- [x] Tips: escape to interrupt, ⇧Tab thinking, type / for commands, ⌃O collapse internals, enter to steer
- [x] Timer-based rotation (~4.5s) while streaming
- [x] Smooth crossfade transition between tips (0.25s opacity via .pi-working--fading)

### 2. Working indicator: whimsical "Working…" text (ties into #9, now closed)
- [x] Replace static "Working…" with rotating whimsical messages
- [x] Spreadsheet-themed/playful messages (10 variants)
- [x] Separate rotation timer from hints (6s vs 4.5s, staggered by 2s)
- [x] Smooth crossfade on the whimsical text too

### 3. Non-working-state discoverability
- [x] Model picker: tooltip "Click to change model"
- [x] Thinking level: tooltip "Click or ⇧Tab to cycle thinking depth"
- [x] Context usage: tooltip "Context window usage"
- [x] Input placeholder: rotates through "Ask about your spreadsheet…" / "Type / for commands…" every 8s

### 4. Polish
- [x] Both rotating texts staggered so they don't change simultaneously
- [x] Crossfade transitions smooth (0.25s opacity)
- [x] All timers cleaned up on disconnect / active=false
- [x] Clean compile, no regressions
