#!/bin/zsh
set -euo pipefail

PROJECT_DIR="/Users/tom/Desktop/Elave Dash Sales Report dash board"
LOG_FILE="$HOME/Library/Logs/elave-dashboard-sync.log"
STAMP="$(date -u '+%Y-%m-%dT%H:%M:%SZ')"

mkdir -p "$(dirname "$LOG_FILE")"
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin:$PATH"

{
  echo "[$STAMP] sync start"
  cd "$PROJECT_DIR"
  npm run sync:supabase
  echo "[$(date -u '+%Y-%m-%dT%H:%M:%SZ')] sync end"
} >>"$LOG_FILE" 2>&1
