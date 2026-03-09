#!/bin/bash
# Oryx extracts app to /tmp/<hash>/ — use the script's own directory as app root
APP_DIR="$(cd "$(dirname "$0")" && pwd)"

export DATA_DIR="$APP_DIR/data"
export XLSX_PATH="$APP_DIR/data/SUPERDATASETCLEANED.xlsx"
export VOL_A_HTML_CACHE_DIR=/home/vol_a_html_cache

echo "[startup] APP_DIR: $APP_DIR"
echo "[startup] XLSX exists: $([ -f "$XLSX_PATH" ] && echo YES || echo NO)"

exec python -m uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
