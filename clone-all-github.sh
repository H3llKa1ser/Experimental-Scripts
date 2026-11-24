#!/usr/bin/env bash
# clone-all-github.sh — clone every repo owned by a GitHub user as normal working copies
# (real files on disk, submodules + LFS included). Re-running updates existing clones.
#
# Prereqs (once):  gh auth login && gh auth setup-git
# Usage:           bash clone-all-github.sh
#                  ALL_BRANCHES=1 bash clone-all-github.sh     # also check out every branch
#                  SKIP_FORKS=1 SKIP_ARCHIVED=1 bash clone-all-github.sh
set -euo pipefail

GH_USER="${GH_USER:-H3llKa1ser}"
DEST="${DEST:-$HOME/github-repos}"
ALL_BRANCHES="${ALL_BRANCHES:-0}"
SKIP_FORKS="${SKIP_FORKS:-0}"
SKIP_ARCHIVED="${SKIP_ARCHIVED:-0}"

export GIT_TERMINAL_PROMPT=0          # never hang on a credential prompt
command -v gh  >/dev/null || { echo "gh CLI not found"; exit 1; }
command -v git >/dev/null || { echo "git not found"; exit 1; }
HAS_LFS=0; git lfs version &>/dev/null && HAS_LFS=1

mkdir -p "$DEST"
ok=0; updated=0; failed=0; empty=0
failures=()

filter='.[]'
[ "$SKIP_FORKS" = 1 ]    && filter+=' | select(.isFork | not)'
[ "$SKIP_ARCHIVED" = 1 ] && filter+=' | select(.isArchived | not)'
filter+=' | .name'

checkout_all_branches() {
  local repo_dir="$1" wt_root="$2" current b path
  current=$(git -C "$repo_dir" symbolic-ref --short -q HEAD || true)
  git -C "$repo_dir" worktree prune
  git -C "$repo_dir" for-each-ref --format='%(refname:strip=3)' refs/remotes/origin |
  while read -r b; do
    [ "$b" = "HEAD" ] || [ "$b" = "$current" ] && continue
    path="$wt_root/$b"
    if [ -d "$path" ]; then
      git -C "$path" pull --ff-only --quiet </dev/null || echo "      [!] branch $b diverged, skipped"
    elif git -C "$repo_dir" show-ref --verify -q "refs/heads/$b"; then
      git -C "$repo_dir" worktree add --quiet "$path" "$b" </dev/null || echo "      [!] worktree $b failed"
    else
      git -C "$repo_dir" worktree add --quiet --track -b "$b" "$path" "origin/$b" </dev/null || echo "      [!] worktree $b failed"
    fi
  done
}

while read -r repo; do
  dir="$DEST/$repo"
  url="https://github.com/$GH_USER/$repo.git"

  if [ -d "$dir/.git" ]; then
    echo "[~] $repo"
    if git -C "$dir" fetch --all --prune --tags --quiet </dev/null &&
       { git -C "$dir" rev-parse -q --verify HEAD >/dev/null || continue; } &&
       git -C "$dir" pull --ff-only --quiet </dev/null; then
      git -C "$dir" submodule update --init --recursive --quiet </dev/null || true
      updated=$((updated+1))
    else
      echo "    [!] update failed (local changes or diverged)"; failures+=("$repo"); failed=$((failed+1)); continue
    fi
  else
    echo "[+] $repo"
    if ! git clone --quiet --recurse-submodules "$url" "$dir" </dev/null; then
      echo "    [!] clone failed"; failures+=("$repo"); failed=$((failed+1)); continue
    fi
    if ! git -C "$dir" rev-parse -q --verify HEAD >/dev/null; then
      echo "    [-] empty repo"; empty=$((empty+1)); continue
    fi
    ok=$((ok+1))
  fi

  # Pull real LFS binaries instead of pointer files
  if [ "$HAS_LFS" = 1 ] && [ -f "$dir/.gitattributes" ] && grep -q 'filter=lfs' "$dir/.gitattributes"; then
    git -C "$dir" lfs pull </dev/null || echo "    [!] LFS pull failed"
  fi

  [ "$ALL_BRANCHES" = 1 ] && checkout_all_branches "$dir" "$DEST/${repo}@branches" || true

done < <(gh repo list "$GH_USER" --limit 1000 --json name,isFork,isArchived -q "$filter")

echo
echo "Done → $DEST"
echo "  cloned: $ok  updated: $updated  empty: $empty  failed: $failed"
[ "$failed" -gt 0 ] && printf '  failed: %s\n' "${failures[@]}"
exit 0
