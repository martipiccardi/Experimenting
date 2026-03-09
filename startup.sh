#!/bin/bash
export DATA_DIR=/home/site/wwwroot/data
export XLSX_PATH=/home/site/wwwroot/data/SUPERDATASETCLEANED.xlsx
export VOL_A_HTML_CACHE_DIR=/home/vol_a_html_cache

VENV=/home/site/venv
WHEELS=/home/site/wwwroot/wheels
REQ=/home/site/wwwroot/requirements.txt

_verify() {
    python -c "import uvicorn, fastapi, duckdb, pandas, openpyxl" 2>/dev/null
}

_run() {
    cd /home/site/wwwroot
    exec python -m uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
}

# 1. Oryx-built venv (created when SCM_DO_BUILD_DURING_DEPLOYMENT=true)
if [ -f "/antenv/bin/activate" ]; then
    source /antenv/bin/activate
    if _verify; then echo "[startup] /antenv OK"; _run; fi
    deactivate 2>/dev/null
fi

# 2. Persistent venv on /home (survives redeployments)
if [ -f "$VENV/bin/activate" ]; then
    source "$VENV/bin/activate"
    if _verify; then echo "[startup] $VENV OK"; _run; fi
    echo "[startup] $VENV broken — rebuilding"
    deactivate 2>/dev/null
    rm -rf "$VENV"
fi

# 3. Install from bundled manylinux wheels (fast, no network, correct glibc)
echo "[startup] Installing from bundled wheels..."
python -m venv "$VENV"
source "$VENV/bin/activate"
if [ -d "$WHEELS" ]; then
    pip install --no-cache-dir --no-index --find-links "$WHEELS" -r "$REQ" \
        && echo "[startup] Wheels install OK" && _run
fi

# 4. Last resort: pip install from network
echo "[startup] Falling back to network pip install..."
pip install --no-cache-dir -q -r "$REQ"
_run
