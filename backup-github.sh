#!/usr/bin/env bash
# backup-github.sh — mirror all repos owned by H3llKa1ser
set -euo pipefail
USER=H3llKa1ser
DEST=~/github-backup
mkdir -p "$DEST" && cd "$DEST"

gh repo list "$USER" --limit 1000 --json name -q '.[].name' | while read -r repo; do
  if [ -d "$repo.git" ]; then
    echo "Updating $repo"
    git -C "$repo.git" remote update --prune
  else
    echo "Cloning $repo"
    git clone --mirror "https://github.com/$USER/$repo.git" "$repo.git"
  fi
  # wikis live in a separate repo; skip silently if none exists
  if [ ! -d "$repo.wiki.git" ]; then
    git clone --mirror "https://github.com/$USER/$repo.wiki.git" "$repo.wiki.git" 2>/dev/null || true
  fi
done
