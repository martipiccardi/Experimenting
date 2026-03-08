#!/bin/bash
export DATA_DIR=/home/site/wwwroot/data
export XLSX_PATH=/home/site/wwwroot/data/SUPERDATASETCLEANED.xlsx
export VOL_A_HTML_CACHE_DIR=/home/vol_a_html_cache

# Prefer the Oryx-built venv (created during deployment via SCM_DO_BUILD_DURING_DEPLOYMENT).
# Fall back to a persistent venv on /home only if Oryx venv is absent.
if [ -f "/antenv/bin/uvicorn" ]; then
    echo "Using Oryx antenv."
    source /antenv/bin/activate
elif [ -f "/home/site/wwwroot/antenv/bin/uvicorn" ]; then
    echo "Using wwwroot antenv."
    source /home/site/wwwroot/antenv/bin/activate
else
    VENV=/home/site/venv
    if [ ! -f "$VENV/bin/uvicorn" ]; then
        echo "No Oryx venv found — installing core dependencies..."
        rm -rf "$VENV"
        python -m venv "$VENV"
        source "$VENV/bin/activate"
        pip install --no-cache-dir -r /home/site/wwwroot/requirements.txt
        echo "Done."
    else
        echo "Reusing existing venv."
        source "$VENV/bin/activate"
    fi
fi

cd /home/site/wwwroot
exec uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
