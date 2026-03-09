#!/bin/bash
export DATA_DIR=/home/site/wwwroot/data
export XLSX_PATH=/home/site/wwwroot/data/SUPERDATASETCLEANED.xlsx
export VOL_A_HTML_CACHE_DIR=/home/vol_a_html_cache

VENV=/home/site/venv
REQ=/home/site/wwwroot/requirements.txt

_run() {
    cd /home/site/wwwroot
    exec python -m uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
}

_verify() {
    python -c "import uvicorn, fastapi, duckdb, pandas, openpyxl" 2>/dev/null
}

# 1. Oryx-built venv (created when SCM_DO_BUILD_DURING_DEPLOYMENT=true)
if [ -f "/antenv/bin/activate" ]; then
    echo "[startup] Trying Oryx /antenv..."
    source /antenv/bin/activate
    if _verify; then
        echo "[startup] Oryx antenv OK"
        _run
    fi
    echo "[startup] Oryx antenv broken — falling back"
    deactivate 2>/dev/null
fi

# 2. Persistent venv — verify all key packages are importable
if [ -f "$VENV/bin/activate" ]; then
    source "$VENV/bin/activate"
    if _verify; then
        echo "[startup] Persistent venv OK"
        _run
    fi
    echo "[startup] Persistent venv incomplete — rebuilding..."
    deactivate 2>/dev/null
fi

# 3. Build persistent venv from scratch
echo "[startup] Creating $VENV and installing packages..."
rm -rf "$VENV"
python -m venv "$VENV"
source "$VENV/bin/activate"
pip install --no-cache-dir -q -r "$REQ"
echo "[startup] Install complete"
_run
