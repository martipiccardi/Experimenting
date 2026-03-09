#!/bin/bash
export DATA_DIR=/home/site/wwwroot/data
export XLSX_PATH=/home/site/wwwroot/data/SUPERDATASETCLEANED.xlsx
export VOL_A_HTML_CACHE_DIR=/home/vol_a_html_cache

# /tmp is local disk (fast writes). /home is Azure Files (slow SMB writes).
VENV=/tmp/venv
WHEELS=/home/site/wwwroot/wheels
REQ=/home/site/wwwroot/requirements.txt

_verify() {
    python -c "import uvicorn, fastapi, duckdb, pandas, openpyxl" 2>/dev/null
}

_run() {
    cd /home/site/wwwroot
    exec python -m uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
}

# 1. Oryx-built venv (local disk, created during deployment)
if [ -f "/antenv/bin/activate" ]; then
    source /antenv/bin/activate
    if _verify; then echo "[startup] /antenv OK"; _run; fi
    deactivate 2>/dev/null
fi

# 2. Temp venv from same container instance (survives app restarts within same container)
if [ -f "$VENV/bin/activate" ]; then
    source "$VENV/bin/activate"
    if _verify; then echo "[startup] $VENV OK"; _run; fi
    deactivate 2>/dev/null
    rm -rf "$VENV"
fi

# 3. Install from bundled wheels to /tmp (local disk — fast, fits in 230s)
echo "[startup] Installing from bundled wheels to /tmp/venv..."
python -m venv "$VENV"
source "$VENV/bin/activate"
if [ -d "$WHEELS" ]; then
    pip install --no-cache-dir --no-index --find-links "$WHEELS" -r "$REQ" \
        && echo "[startup] Wheels install OK" && _run
fi

# 4. Network fallback (last resort)
echo "[startup] Network fallback..."
pip install --no-cache-dir -q -r "$REQ"
_run
