#!/bin/bash
export DATA_DIR=/home/site/wwwroot/data
export XLSX_PATH=/home/site/wwwroot/data/SUPERDATASETCLEANED.xlsx
export VOL_A_HTML_CACHE_DIR=/home/vol_a_html_cache

# Prefer the Oryx-built venv (packages installed during deployment, before startup timer).
# This avoids pip install at startup and eliminates ContainerTimeout risk.
if [ -f "/antenv/bin/activate" ]; then
    echo "Using Oryx venv (/antenv)."
    source /antenv/bin/activate
else
    # Fallback: persistent custom venv with stamp file to skip reinstall on restarts.
    VENV=/home/site/venv
    STAMP="$VENV/.stamp"
    EXPECTED_STAMP="v1-core-hf"
    if [ ! -f "$STAMP" ] || [ "$(cat "$STAMP")" != "$EXPECTED_STAMP" ]; then
        echo "Building venv (first run)..."
        rm -rf "$VENV"
        python -m venv "$VENV"
        source "$VENV/bin/activate"
        pip install --no-cache-dir pandas==2.2.2 openpyxl==3.1.5 xlrd==2.0.1 duckdb==1.0.0 \
            pyarrow==17.0.0 fastapi==0.115.0 "uvicorn[standard]==0.30.6" numpy requests
        if [ -z "$HF_API_TOKEN" ]; then
            echo "Installing sentence-transformers..."
            pip install --no-cache-dir torch --index-url https://download.pytorch.org/whl/cpu
            pip install --no-cache-dir sentence-transformers
        fi
        echo "$EXPECTED_STAMP" > "$STAMP"
    else
        echo "Reusing existing venv."
        source "$VENV/bin/activate"
    fi
fi

cd /home/site/wwwroot
exec uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
