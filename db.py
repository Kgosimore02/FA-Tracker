"""DB layer — pool, cached reads, writes, view refresh.

st.cache_data cannot hash arbitrary dicts passed as function args.
Fix: serialize params to a JSON string for the cache key, execute
with the original dict. Use keyword-only args so Streamlit hashes
only the sql + params_key strings (both hashable).
"""
import os, json
import streamlit as st
from sqlalchemy import create_engine, text
from sqlalchemy.pool import QueuePool

_engine = None

def get_engine():
    global _engine
    if _engine is None:
        url = os.getenv(
            "DATABASE_URL",
            "postgresql://neondb_owner:npg_hH6enqf0JUVT@ep-dark-tree-ay8819gw.c-5.us-east-2.aws.neon.tech/neondb?sslmode=require"
        )
        _engine = create_engine(
            url, poolclass=QueuePool,
            pool_size=10, max_overflow=20,
            pool_timeout=30, pool_pre_ping=True,
            pool_recycle=1800,
            connect_args={"connect_timeout": 10,
                          "application_name": "BLIL_FA_Tracker"},
        )
    return _engine


def _rows(result) -> list[dict]:
    """Convert SQLAlchemy result to plain dicts (cache-safe, no Row objects)."""
    keys = list(result.keys())
    return [dict(zip(keys, row)) for row in result.fetchall()]


@st.cache_data(ttl=60, show_spinner=False)
def _cached_query(sql: str, params_json: str) -> list[dict]:
    """Cache key = (sql, params_json). Both are plain strings — always hashable."""
    params = json.loads(params_json) if params_json != "{}" else {}
    with get_engine().connect() as conn:
        return _rows(conn.execute(text(sql), params))


def run_query(sql: str, params: dict = None) -> list[dict]:
    p = params or {}
    return _cached_query(sql, json.dumps(p, default=str, sort_keys=True))


def run_query_live(sql: str, params: dict = None) -> list[dict]:
    """Bypass cache — use immediately after writes."""
    with get_engine().connect() as conn:
        return _rows(conn.execute(text(sql), params or {}))


def run_write(sql: str, params: dict = None):
    with get_engine().begin() as conn:
        conn.execute(text(sql), params or {})


def run_write_many(stmts: list[tuple]):
    with get_engine().begin() as conn:
        for sql, params in stmts:
            conn.execute(text(sql), params or {})


def refresh_views():
    order = ["mv_daily_summary", "mv_weekly_appointment_summary",
             "mv_monthly_branch_summary", "mv_quarterly_summary"]
    with get_engine().begin() as conn:
        for v in order:
            try:
                conn.execute(text(f"REFRESH MATERIALIZED VIEW CONCURRENTLY {v}"))
            except Exception:
                conn.execute(text(f"REFRESH MATERIALIZED VIEW {v}"))
    _cached_query.clear()