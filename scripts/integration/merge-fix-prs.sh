#!/usr/bin/env bash
# Rebuild the integration-fixes branch: origin/main + every open fix PR, merged
# (real merges, --no-ff) in a fixed order, then the integration follow-up fixes.
#
# Usage:
#   cp scripts/integration/{merge-fix-prs.sh,json3way.py} /tmp/integ/
#   /tmp/integ/merge-fix-prs.sh          # start fresh from origin/main
#   /tmp/integ/merge-fix-prs.sh --resume # continue after resolving a conflict
#
# Conflict resolutions are recorded with git rerere (rerere.enabled + autoupdate),
# so a rebuild replays earlier resolutions; it stops only on new conflicts.
# After resolving one by hand: `git add -A && git commit --no-edit --no-verify`, then --resume.
# (--no-verify: the pre-commit check may fail mid-build; the follow-up fixes repair it.)
#
# Follow-up fixes ("Integration fix: ..." commits) are cherry-picked from the
# list in scripts/integration/fixes.txt on the previous integration branch
# (kept as $BRANCH-previous). Each line: "<sha> <subject>". The tooling and
# INTEGRATION.md are re-committed last; edit INTEGRATION.md for new conflicts.
set -euo pipefail

BRANCH=integration-fixes
# Order matters: stacked PRs follow their bases.
PRS=(
  "fix-transcript 786"
  "agent-reads 787"
  "agent-writes 791"
  "agent-python 797"
  "fix-guidance-files 788"
  "fix-turn-composer 789"
  "agent-context 792"
  "compaction-summary 798"
  "fix-backups-approvals 794"
  "fix-sessions-models 795"
  "copilot-login 799"
  "fix-host-web 796"
  "fix-extensions-integrations 790"
  "ux-system 793"
)
# Add later PRs to PRS above.
PINS=/tmp/integration-pins.txt
# Run from a copy outside the worktree (the branch is reset under it).
TOOLS="$(cd "$(dirname "$0")" && pwd)"

git config rerere.enabled true
git config rerere.autoupdate true

pin() { awk -v b="$1" '$1 == b { print $2 }' "$PINS"; }
branches=()
for entry in "${PRS[@]}"; do branches+=("${entry%% *}"); done

if [[ "${1:-}" != "--resume" ]]; then
  # Fetch once and pin every head, so --resume merges the same commits even
  # if a PR pushes mid-build.
  git fetch --quiet origin main "${branches[@]}"
  : > "$PINS"
  for b in main "${branches[@]}"; do echo "$b $(git rev-parse "origin/$b")" >> "$PINS"; done
  # Keep the previous branch's fix list before resetting.
  if git rev-parse --verify --quiet "$BRANCH" >/dev/null; then
    git show "$BRANCH:scripts/integration/fixes.txt" > /tmp/integration-fixes.txt 2>/dev/null || : > /tmp/integration-fixes.txt
    git branch -f "$BRANCH-previous" "$BRANCH"
  else
    : > /tmp/integration-fixes.txt
  fi
  git checkout --quiet -B "$BRANCH" "$(pin main)"
fi

for entry in "${PRS[@]}"; do
  b="${entry%% *}"; n="${entry##* }"; sha="$(pin "$b")"
  if git merge-base --is-ancestor "$sha" HEAD; then
    echo "already merged: $b (#$n)"; continue
  fi
  echo "merging $b (#$n) @ ${sha:0:7}"
  if ! git merge --no-ff --no-edit -m "Merge $b (#$n) @ ${sha:0:7} into $BRANCH" "$sha"; then
    # Agreed: #794 deletes save-boundary-monitor.ts, so #792's change to it is dropped
    # (rerere cannot record a modify/delete resolution).
    if [[ "$b" == fix-backups-approvals ]] && git diff --name-only --diff-filter=U | grep -qx src/workbook/save-boundary-monitor.ts; then
      git rm -q src/workbook/save-boundary-monitor.ts
    fi
    # Locale JSON: merge per key when no key changed on both sides.
    locales="$(git diff --name-only --diff-filter=U -- 'src/language/locales/*.json')"
    if [[ -n "$locales" ]]; then
      both="$(python3 "$TOOLS/json3way.py" $locales)"
      if [[ -z "$both" ]]; then git add $locales; else echo "$both"; fi
    fi
    if [[ -z "$(git diff --name-only --diff-filter=U)" ]]; then
      echo "rerere resolved all conflicts for $b; review with: git diff --cached"
      git commit --no-edit --no-verify
    else
      echo "CONFLICT merging $b (#$n). Resolve, record in INTEGRATION.md, commit, then --resume."
      git diff --name-only --diff-filter=U
      exit 1
    fi
  fi
done

# Replay follow-up integration fixes (listed in fixes.txt on the previous
# integration branch), then re-commit the tooling and INTEGRATION.md last.
if [[ -s /tmp/integration-fixes.txt ]]; then
  while read -r sha subject; do
    [[ -z "$sha" || "$sha" == \#* ]] && continue
    if git log --format=%s "$(pin main)"..HEAD | grep -qxF "$subject"; then
      echo "fix already present: $subject"; continue
    fi
    echo "cherry-picking fix: $subject"
    git cherry-pick --empty=drop "$sha" >/dev/null || {
      echo "Fix $sha conflicts; resolve, git cherry-pick --continue, then --resume"; exit 1; }
  done < /tmp/integration-fixes.txt
fi

if [[ ! -f INTEGRATION.md ]] && git rev-parse --verify --quiet "$BRANCH-previous" >/dev/null; then
  git checkout "$BRANCH-previous" -- INTEGRATION.md
fi
mkdir -p scripts/integration
cp "$TOOLS/merge-fix-prs.sh" "$TOOLS/json3way.py" scripts/integration/
{
  last_merge="$(git log --first-parent --merges -1 --format=%H)"
  git log --reverse --format='%H %s' --grep='^Integration fix:' "$last_merge"..HEAD > scripts/integration/fixes.txt
  cp /tmp/integration-pins.txt scripts/integration/pins.txt
  git add scripts/integration INTEGRATION.md 2>/dev/null || git add scripts/integration
  git commit --no-verify -q -m "Integration: merge script, pins and conflict log (INTEGRATION.md)" || true
}
echo "done: $(git rev-parse HEAD)"
