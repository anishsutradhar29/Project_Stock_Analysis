#!/usr/bin/env bash
# Regenerates the notebook from the .py scripts, commits + pushes to GitHub,
# and refreshes the shared Google Drive copy — all in one go.
#
# Usage:
#   ./update_and_share.sh "commit message here"
#   ./update_and_share.sh              (prompts for a commit message instead)

set -euo pipefail

# Always run from this script's own directory, regardless of where it's called from
cd "$(dirname "$0")"

COMMIT_MSG="${1:-}"
if [ -z "$COMMIT_MSG" ]; then
    read -rp "Commit message: " COMMIT_MSG
fi

echo "==> Syncing notebook from the .py scripts (update_notebook.py)"
python3 update_notebook.py

echo "==> Sanity-checking both scripts still compile"
python3 -m py_compile NSE_Stock_Analysis.py SGX_Stock_Analysis.py

echo "==> Staging changes"
git add -A

if git diff --cached --quiet; then
    echo "==> Nothing new to commit (working tree already matches last commit)"
else
    echo "==> Committing: $COMMIT_MSG"
    git commit -m "$COMMIT_MSG"
fi

echo "==> Pushing to GitHub"
git push

echo "==> Refreshing shared Google Drive copy"
rclone copyto Stock_Market_Analysis_SGX_NSE.ipynb \
    "gdrive:Colab Notebooks/Project_Stock_Analysis/Stock_Market_Analysis_SGX_NSE.ipynb"

echo "==> Done."
