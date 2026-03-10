import os
import threading
from pathlib import Path

import duckdb
import pandas as pd

_DEFAULT_DATA_DIR = Path(__file__).resolve().parent.parent.parent / "data"
DATA_DIR = Path(os.environ.get("DATA_DIR", str(_DEFAULT_DATA_DIR)))
XLSX_PATH = Path(os.environ.get("XLSX_PATH", str(DATA_DIR / "SUPERDATASETCLEANED.xlsx")))

# DB_PATH can be set independently so the DuckDB survives zip re-deployments.
# On Azure set DB_PATH=/home/enes.duckdb (outside wwwroot, persists across deploys).
_DEFAULT_DB_PATH = DATA_DIR / "enes.duckdb"
DB_PATH = Path(os.environ.get("DB_PATH", str(_DEFAULT_DB_PATH)))

# Process-level flag + lock so concurrent requests never race to CREATE TABLE.
# After the first successful init, all threads skip straight through (fast path).
_table_ready = False
_table_lock = threading.Lock()


def get_conn():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    return duckdb.connect(str(DB_PATH))


def ensure_table(con: duckdb.DuckDBPyConnection):
    """Create (or refresh) TABLE enes from Excel.

    Stores the Excel file's modification time in a metadata table.
    If the Excel has changed since the last build, the table is dropped
    and rebuilt automatically — no manual intervention required.
    """
    global _table_ready
    if _table_ready:
        return  # fast path: already confirmed ready this process lifetime

    with _table_lock:
        if _table_ready:
            return  # another thread finished while we waited

        if not XLSX_PATH.exists():
            raise FileNotFoundError(f"Excel file not found at {XLSX_PATH}")

        current_mtime = str(XLSX_PATH.stat().st_mtime)

        exists = con.execute("""
            SELECT COUNT(*) FROM information_schema.tables WHERE table_name = 'enes'
        """).fetchone()[0]

        if exists:
            try:
                row = con.execute("SELECT mtime FROM enes_meta LIMIT 1").fetchone()
                if row and row[0] == current_mtime:
                    _table_ready = True
                    return  # Excel unchanged — reuse existing DuckDB
            except Exception:
                pass  # enes_meta missing or corrupt — fall through to rebuild
            con.execute("DROP TABLE IF EXISTS enes")
            con.execute("DROP TABLE IF EXISTS enes_meta")

        df = pd.read_excel(XLSX_PATH, engine="openpyxl")
        con.register("enes_df", df)
        con.execute("CREATE TABLE enes AS SELECT * FROM enes_df")
        con.execute("CREATE TABLE enes_meta (mtime VARCHAR)")
        con.execute("INSERT INTO enes_meta VALUES (?)", [current_mtime])
        _table_ready = True





