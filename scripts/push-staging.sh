#!/usr/bin/env bash
set -euo pipefail
BRANCH="staging"
REMOTE_NORMAL="https://github.com/admin-ops-stellaris/integrity.git"
REMOTE_AUTH="https://x-access-token:${GITHUB_PAT}@github.com/admin-ops-stellaris/integrity.git"
# Ensure token exists
if [[ -z "${GITHUB_PAT:-}" ]]; then
  echo "ERROR: GITHUB_PAT secret is not set in Replit Secrets."
  exit 1
fi
# Ensure we're on staging
current_branch="$(git rev-parse --abbrev-ref HEAD)"
if [[ "$current_branch" != "$BRANCH" ]]; then
  echo "ERROR: You're on '$current_branch' but this script pushes '$BRANCH'."
  exit 1
fi
# Commit message required
if [[ $# -lt 1 ]]; then
  echo "Usage: scripts/push-staging.sh \"commit message\""
  exit 1
fi
msg="$1"
echo "==> Status before commit:"
git status
# Stage everything
git add -A
# Commit if there are staged changes
if git diff --cached --quiet; then
  echo "No changes to commit."
else
  git commit -m "$msg"
fi

# Push using PAT, then restore remote URL
git remote set-url origin "$REMOTE_AUTH"
trap 'git remote set-url origin "$REMOTE_NORMAL" >/dev/null 2>&1 || true' EXIT

echo "==> Pushing to origin/$BRANCH ..."
git push origin "$BRANCH"

echo "✅ Pushed to $BRANCH. GitHub Actions should deploy Fly staging."
