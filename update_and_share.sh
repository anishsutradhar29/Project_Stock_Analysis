#!/usr/bin/env bash
# Regenerates both scripts' stock lists from stock list.xlsx, then the notebook
# from the .py scripts, commits + pushes to GitHub,
# and refreshes the shared Google Drive copies of both the notebook and
# stock list.xlsx — all in one go.
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

# Use the project's venv Python (has openpyxl, needed to read stock list.xlsx);
# fall back to the system python3 on machines without a venv/ folder.
PY="venv/bin/python"
[ -x "$PY" ] || PY="python3"

echo "==> Syncing stock lists from stock list.xlsx + notebook from the .py scripts (update_notebook.py)"
"$PY" update_notebook.py

echo "==> Sanity-checking all scripts still compile"
"$PY" -m py_compile NSE_Stock_Analysis.py SGX_Stock_Analysis.py sync_stock_lists.py update_notebook.py

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

echo "==> Refreshing shared Google Drive copy of the notebook"
rclone copyto Stock_Market_Analysis_SGX_NSE.ipynb \
    "gdrive:Colab Notebooks/Project_Stock_Analysis/Stock_Market_Analysis_SGX_NSE.ipynb"

echo "==> Refreshing shared Google Drive copy of stock list.xlsx"
rclone copyto "stock list.xlsx" \
    "gdrive:Colab Notebooks/Project_Stock_Analysis/stock list.xlsx"

echo "==> Done."
