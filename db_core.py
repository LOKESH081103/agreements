"""
db_core.py
==========
Framework-agnostic core: DB engine, schema introspection, query building,
and the full Excel-style pivot / DPD-bucket / Field-Efficiency engine --
extracted VERBATIM (same functions, same logic, only decorators/config
swapped) from the Streamlit app's pure pandas/SQLAlchemy code, so the
Streamlit UI and this FastAPI service always compute identical numbers.

Streamlit-only concerns (st.cache_data/st.cache_resource, st.secrets,
st.error/UI rendering) are NOT here -- replaced below with:
  - plain environment-variable config (no st.secrets)
  - a tiny in-process TTL cache for schema/lookup calls
  - a module-level singleton engine, created on first use / at API startup

This module raises plain exceptions instead of calling st.error(); the
FastAPI layer (api.py) turns those into HTTP error responses.
"""

from __future__ import annotations

import functools
import io
import itertools
import os
import re
import time
from datetime import date
from typing import Any, Callable

import numpy as np
import pandas as pd
import xlsxwriter
import bcrypt
from sqlalchemy import create_engine, inspect, text
from sqlalchemy import types as sa_types
from sqlalchemy.engine import Engine
from sqlalchemy.exc import SQLAlchemyError

# --------------------------------------------------------------------------- #
# Config (env vars only -- no st.secrets here; api.py runs outside Streamlit)
# --------------------------------------------------------------------------- #
DB_HOST = os.environ.get("DB_HOST", "localhost")
DB_PORT = os.environ.get("DB_PORT", "5432")
DB_NAME = os.environ.get("DB_NAME", "postgres")
DB_USER = os.environ.get("DB_USER", "postgres")
DB_PASSWORD = os.environ.get("DB_PASSWORD", "")
DB_SSLMODE = os.environ.get("DB_SSLMODE", "prefer")
DATABASE_URL = os.environ.get("DATABASE_URL")

# --------------------------------------------------------------------------- #
# Tiny in-process TTL cache (stands in for @st.cache_data outside Streamlit).
# Deliberately simple: a dict keyed by (func name, args, sorted kwargs),
# values expire after `ttl` seconds. Fine for a single-process API; swap for
# Redis/`fastapi-cache` if you run multiple worker processes and want a
# shared cache across them.
# --------------------------------------------------------------------------- #
_ttl_cache_store: dict[tuple, tuple[float, Any]] = {}


def ttl_cache(ttl: int = 300) -> Callable:
    """TTL cache that degrades gracefully to "no cache" for calls whose
    arguments aren't hashable (e.g. a DataFrame or a plain dict passed
    positionally) instead of raising -- safe to apply broadly."""
    def decorator(fn: Callable) -> Callable:
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            try:
                key = (fn.__name__, args, tuple(sorted(kwargs.items())))
                hash(key)
            except TypeError:
                return fn(*args, **kwargs)
            now = time.time()
            hit = _ttl_cache_store.get(key)
            if hit is not None and now - hit[0] < ttl:
                return hit[1]
            result = fn(*args, **kwargs)
            _ttl_cache_store[key] = (now, result)
            return result
        return wrapper
    return decorator


# --------------------------------------------------------------------------- #
# Database engine (module-level singleton, built on first use)
# --------------------------------------------------------------------------- #
_engine: Engine | None = None


def get_engine() -> Engine:
    """Build and cache a SQLAlchemy engine for the process's lifetime."""
    global _engine
    if _engine is not None:
        return _engine

    if DATABASE_URL:
        url = DATABASE_URL
    else:
        url = (
            f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}"
            f"@{DB_HOST}:{DB_PORT}/{DB_NAME}?sslmode={DB_SSLMODE}"
        )

    engine = create_engine(
        url,
        pool_pre_ping=True,
        pool_size=5,
        max_overflow=5,
        pool_recycle=1800,
    )
    with engine.connect() as conn:
        conn.execute(text("SELECT 1"))
    _engine = engine
    return engine


MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
_MONTH_FULL_TO_NUM = {name: i + 1 for i, name in enumerate(MONTH_NAMES)}
_MONTH_ABBR_TO_NUM = {name[:3]: i + 1 for i, name in enumerate(MONTH_NAMES)}

CRORE = 10_000_000  # 1,00,00,000

_XL_HEADER_FILL = "1F3864"
_XL_HEADER_FONT = "FFFFFF"
_XL_GROUP_ROW_FILL = "DCE6F1"
_XL_TOTAL_COL_FILL = "F2F2F2"
_XL_TOTAL_ROW_FILL = "DCE6F1"
_XL_CORNER_FILL = "B4C6E7"
_XL_LEVEL_FONT_COLORS = [
    "1F4E78", "2E7D32", "B45309", "6B21A8", "9C2E38",
]

# --------------------------------------------------------------------------- #
# Auth tables (same DDL as app.py's ensure_auth_tables -- both processes
# point at the identical three tables, so an employee's password/access
# level is shared between the Streamlit login screen and this API).
# --------------------------------------------------------------------------- #
_CREATE_EMPLOYEES_SQL = text(
    """
    CREATE TABLE IF NOT EXISTS authorized_employees (
        employee_code   VARCHAR(20) PRIMARY KEY,
        employee_name   VARCHAR(100) NOT NULL,
        department      VARCHAR(100),
        access_level    VARCHAR(20) NOT NULL DEFAULT 'user'
                        CHECK (access_level IN ('admin', 'user', 'analyst')),
        is_authorized   BOOLEAN NOT NULL DEFAULT TRUE,
        created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
    )
    """
)
_CREATE_USERS_SQL = text(
    """
    CREATE TABLE IF NOT EXISTS app_users (
        employee_code   VARCHAR(20) PRIMARY KEY
                        REFERENCES authorized_employees(employee_code) ON DELETE CASCADE,
        password_hash   TEXT NOT NULL,
        created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
        last_login      TIMESTAMPTZ
    )
    """
)
_CREATE_AUDIT_SQL = text(
    """
    CREATE TABLE IF NOT EXISTS login_audit (
        id              BIGSERIAL PRIMARY KEY,
        employee_code   VARCHAR(20) NOT NULL,
        event           VARCHAR(20) NOT NULL
                        CHECK (event IN ('login_success', 'login_failure', 'logout', 'account_created')),
        event_time      TIMESTAMPTZ NOT NULL DEFAULT now()
    )
    """
)


# ============================================================================
# Everything below this line is extracted verbatim from app.py's pure
# pandas/SQLAlchemy logic (schema introspection, query building, the
# Excel-style pivot / DPD-bucket / Field-Efficiency engine, and the auth
# helpers). @st.cache_data / @st.cache_resource decorators were stripped;
# the hot, small, hashable-arg lookups (list_tables, get_columns, distinct
# values, year/month lookups) are re-wrapped in @ttl_cache below them.
# ============================================================================

def get_engine() -> Engine:
    """
    Build and cache a SQLAlchemy engine for the app's lifetime.

    Using @st.cache_resource ensures a single connection pool is reused
    across reruns instead of opening a new connection on every interaction.
    """
    if DATABASE_URL:
        url = DATABASE_URL
    else:
        url = (
            f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}"
            f"@{DB_HOST}:{DB_PORT}/{DB_NAME}?sslmode={DB_SSLMODE}"
        )

    engine = create_engine(
        url,
        pool_pre_ping=True,   # detect dead connections and recycle them
        pool_size=5,
        max_overflow=5,
        pool_recycle=1800,    # recycle connections every 30 minutes
    )
    # Fail fast if credentials/host are wrong, so we can show a clean error
    # at startup rather than on the first query the user runs.
    with engine.connect() as conn:
        conn.execute(text("SELECT 1"))
    return engine


@ttl_cache(ttl=300)
def list_tables(_engine: Engine) -> list[str]:
    """List all tables in the 'public' schema."""
    inspector = inspect(_engine)
    return sorted(inspector.get_table_names(schema="public"))




@ttl_cache(ttl=300)
def get_columns(_engine: Engine, table_name: str) -> list[dict[str, Any]]:
    """Return column metadata (name + SQLAlchemy/py type) for a table."""
    inspector = inspect(_engine)
    return inspector.get_columns(table_name, schema="public")


def classify_column(col_type: Any) -> str:
    """
    Classify a SQLAlchemy column type into one of: 'date', 'numeric', 'text'.
    Falls back to 'text' for anything unrecognized (safe default).
    """
    type_str = str(col_type).upper()
    date_markers = ("DATE", "TIME", "TIMESTAMP")
    numeric_markers = (
        "INT",
        "NUMERIC",
        "DECIMAL",
        "FLOAT",
        "DOUBLE",
        "REAL",
        "MONEY",
        "SERIAL", 
    )
    if any(marker in type_str for marker in date_markers):
        return "date"
    if any(marker in type_str for marker in numeric_markers):
        return "numeric"
    return "text"




@ttl_cache(ttl=300)
def get_distinct_values(_engine: Engine, table_name: str, column_name: str, limit: int = 1000) -> list[Any]:
    """Fetch distinct non-null values for a categorical column (capped)."""
    query = text(
        f'SELECT DISTINCT "{column_name}" FROM "public"."{table_name}" '
        f'WHERE "{column_name}" IS NOT NULL ORDER BY "{column_name}" LIMIT :limit'
    )
    with _engine.connect() as conn:
        result = conn.execute(query, {"limit": limit})
        return [row[0] for row in result]




@ttl_cache(ttl=300)
def get_distinct_values_filtered(
    _engine: Engine,
    table_name: str,
    column_name: str,
    filter_conditions: tuple[tuple[str, tuple[Any, ...]], ...],
    limit: int = 1000,
) -> list[Any]:
    """
    Same as `get_distinct_values`, but narrowed by whatever OTHER Filter
    fields already have a value picked -- this is what makes the Filters
    section cascade: picking a Zone narrows Region's own dropdown down to
    just that zone's regions, instead of always listing every region in the
    whole table. Pushed down as a real SQL WHERE clause, so it stays fast
    even against a large table -- Postgres only has to look at the rows that
    already match, not every row.
    """
    params: dict[str, Any] = {"limit": limit}
    where_parts = [f'"{column_name}" IS NOT NULL']
    for i, (col, values) in enumerate(filter_conditions):
        if not values:
            continue
        placeholders = []
        for j, v in enumerate(values):
            key = f"fv_{i}_{j}"
            placeholders.append(f":{key}")
            params[key] = v
        where_parts.append(f'"{col}" IN ({", ".join(placeholders)})')
    where_clause = " AND ".join(where_parts)
    query = text(
        f'SELECT DISTINCT "{column_name}" FROM "public"."{table_name}" '
        f'WHERE {where_clause} ORDER BY "{column_name}" LIMIT :limit'
    )
    with _engine.connect() as conn:
        result = conn.execute(query, params)
        return [row[0] for row in result]




@ttl_cache(ttl=300)
def get_available_years(_engine: Engine, table_name: str, column_name: str) -> list[int]:
    """Fetch distinct years present in a date/timestamp column."""
    query = text(
        f'SELECT DISTINCT EXTRACT(YEAR FROM "{column_name}")::int AS yr '
        f'FROM "public"."{table_name}" WHERE "{column_name}" IS NOT NULL ORDER BY yr'
    )
    with _engine.connect() as conn:
        result = conn.execute(query)
        return [row[0] for row in result]




@ttl_cache(ttl=300)
def get_available_year_months(_engine: Engine, table_name: str, column_name: str) -> list[tuple[int, int]]:
    """Fetch distinct (year, month) pairs present in a date/timestamp column."""
    query = text(
        f'SELECT DISTINCT EXTRACT(YEAR FROM "{column_name}")::int AS yr, '
        f'EXTRACT(MONTH FROM "{column_name}")::int AS mo '
        f'FROM "public"."{table_name}" WHERE "{column_name}" IS NOT NULL '
        f'ORDER BY yr, mo'
    )
    with _engine.connect() as conn:
        result = conn.execute(query)
        return [(row[0], row[1]) for row in result]


MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
_MONTH_FULL_TO_NUM = {name: i + 1 for i, name in enumerate(MONTH_NAMES)}
_MONTH_ABBR_TO_NUM = {name[:3]: i + 1 for i, name in enumerate(MONTH_NAMES)}

# --------------------------------------------------------------------------- #
# Crore formatting helper
# --------------------------------------------------------------------------- #
CRORE = 10_000_000  # 1,00,00,000


def build_query(
    table_name: str,
    ref_column: str,
    output_columns: list[str],
    ref_kind: str,
    filter_payload: dict[str, Any],
) -> tuple[str, dict[str, Any]]:
    """
    Build a parameterized SQL query (string with :named placeholders + params
    dict) based on the reference column type and the filters chosen in the UI.
    Always includes the reference column itself in the SELECT list so it's
    visible in the results alongside the requested output columns.
    """
    select_cols = list(dict.fromkeys([ref_column, *output_columns]))  # de-dupe, keep order
    quoted_select = ", ".join(f'"{c}"' for c in select_cols)
    base_query = f'SELECT {quoted_select} FROM "public"."{table_name}"'
    params: dict[str, Any] = {}
    where_clause = ""

    if ref_kind == "date":
        granularity = filter_payload.get("granularity")

        if granularity == "Specific Date(s)":
            dates: list[date] = filter_payload.get("dates", [])
            if dates:
                if len(dates) == 2:
                    # If two dates are selected, use BETWEEN to get the full range
                    params["start_date"] = dates[0]
                    params["end_date"] = dates[1]
                    where_clause = f'WHERE "{ref_column}"::date BETWEEN :start_date AND :end_date'
                elif len(dates) == 1:
                    # If only one date is selected, get that exact day
                    params["single_date"] = dates[0]
                    where_clause = f'WHERE "{ref_column}"::date = :single_date'
                else:
                    # Fallback just in case
                    placeholders = []
                    for i, d in enumerate(dates):
                        key = f"date_{i}"
                        placeholders.append(f":{key}")
                        params[key] = d
                    where_clause = f'WHERE "{ref_column}"::date IN ({", ".join(placeholders)})'

        elif granularity == "Month/Year":
            year_months: list[tuple[int, int]] = filter_payload.get("year_months", [])
            if year_months:
                conditions = []
                for i, (yr, mo) in enumerate(year_months):
                    yr_key, mo_key = f"ym_yr_{i}", f"ym_mo_{i}"
                    conditions.append(
                        f'(EXTRACT(YEAR FROM "{ref_column}") = :{yr_key} '
                        f'AND EXTRACT(MONTH FROM "{ref_column}") = :{mo_key})'
                    )
                    params[yr_key] = yr
                    params[mo_key] = mo
                where_clause = "WHERE " + " OR ".join(conditions)

        elif granularity == "Year(s)":
            years: list[int] = filter_payload.get("years", [])
            if years:
                placeholders = []
                for i, yr in enumerate(years):
                    key = f"year_{i}"
                    placeholders.append(f":{key}")
                    params[key] = yr
                where_clause = f'WHERE EXTRACT(YEAR FROM "{ref_column}") IN ({", ".join(placeholders)})'

    else:  # categorical / text
        values: list[Any] = filter_payload.get("values", [])
        if values:
            placeholders = []
            for i, v in enumerate(values):
                key = f"val_{i}"
                placeholders.append(f":{key}")
                params[key] = v
            where_clause = f'WHERE "{ref_column}" IN ({", ".join(placeholders)})'

    query = f"{base_query} {where_clause}".strip()
    return query, params




@ttl_cache(ttl=60)
def run_query(_engine: Engine, query: str, params: dict[str, Any]) -> pd.DataFrame:
    """Execute the (already-parameterized) query and return a DataFrame."""
    with _engine.connect() as conn:
        return pd.read_sql(text(query), conn, params=params)


# --------------------------------------------------------------------------- #
# Excel-style 4-Quadrant Pivot Table engine
# --------------------------------------------------------------------------- #
PIVOT_AGG_FUNCS = ["sum", "mean", "count", "min", "max"]

# Column names used internally for the four pre-aggregated statistics that
# every pivot view (Summary Matrix, Field Efficiency Table, Chart, flat CSV)
# is built from. Carrying all four -- not just whichever agg_func is
# currently selected -- means switching the Values aggregation in the
# sidebar (sum -> mean -> count -> ...) never triggers a new database
# round-trip: it's just a different combination of these same four numbers,
# computed in pandas over an already-tiny, already-cached table.
STAT_SUM = "__stat_sum__"
STAT_COUNT = "__stat_count__"
STAT_MIN = "__stat_min__"
STAT_MAX = "__stat_max__"




@ttl_cache(ttl=600)
def fetch_pivot_source_data(
    _engine: Engine,
    table_name: str,
    row_cols: tuple[str, ...],
    col_cols: tuple[str, ...],
    value_col: str,
    filter_conditions: tuple[tuple[str, tuple[Any, ...]], ...],
    column_kind_map: dict[str, str],
) -> pd.DataFrame:
    """
    Push the aggregation into PostgreSQL instead of pulling raw rows.

    One GROUP BY query computes SUM/COUNT/MIN/MAX of `value_col` per
    (Rows x Columns) combination *inside the database* -- including
    auto-bucketing any date/timestamp Rows/Columns field into a Month level
    above the exact Date (mirroring Excel's automatic date grouping), via a
    `TO_CHAR(...)` expression right in the GROUP BY. For realistic BI
    dimensions this turns a 4M+ row table scan into a query that returns at
    most a few hundred/thousand aggregated rows.

    This is the single biggest lever for large tables: every downstream
    view -- Summary Matrix, Field Efficiency Table, Chart, flat CSV --
    is then built from that tiny table with fast vectorized pandas
    groupbys, never touching millions of raw rows in Python again.

    Returns a DataFrame with one row per unique (expanded) Rows x Columns
    combination, containing:
      - one column per Rows/Columns field (date fields become TWO string
        columns, "__<field>__month" ('YYYY-MM') and "__<field>__date"
        ('YYYY-MM-DD') -- exactly the field-name convention `_expand_date_fields`
        already expects, so no downstream code needs to know this was
        computed in SQL rather than in pandas)
      - STAT_SUM / STAT_COUNT / STAT_MIN / STAT_MAX for `value_col`.
    """
    group_exprs: list[str] = []
    group_by_aliases: list[str] = []
    seen: set[str] = set()

    for f in [*row_cols, *col_cols]:
        if f in seen:
            continue
        seen.add(f)
        if column_kind_map.get(f) == "date":
            month_alias, date_alias = f"__{f}__month", f"__{f}__date"
            group_exprs.append(f"TO_CHAR(\"{f}\", 'YYYY-MM') AS \"{month_alias}\"")
            group_exprs.append(f"TO_CHAR(\"{f}\", 'YYYY-MM-DD') AS \"{date_alias}\"")
            group_by_aliases.extend([f'"{month_alias}"', f'"{date_alias}"'])
        else:
            group_exprs.append(f'"{f}" AS "{f}"')
            group_by_aliases.append(f'"{f}"')

    agg_exprs = [
        f'SUM("{value_col}") AS "{STAT_SUM}"',
        f'COUNT("{value_col}") AS "{STAT_COUNT}"',
        f'MIN("{value_col}") AS "{STAT_MIN}"',
        f'MAX("{value_col}") AS "{STAT_MAX}"',
    ]
    select_clause = ", ".join([*group_exprs, *agg_exprs])
    query = f'SELECT {select_clause} FROM "public"."{table_name}"'

    params: dict[str, Any] = {}
    where_parts: list[str] = []
    for i, (col, values) in enumerate(filter_conditions):
        if not values:
            continue
        placeholders = []
        for j, v in enumerate(values):
            key = f"pf_{i}_{j}"
            placeholders.append(f":{key}")
            params[key] = v
        where_parts.append(f'"{col}" IN ({", ".join(placeholders)})')
    if where_parts:
        query += " WHERE " + " AND ".join(where_parts)

    if group_by_aliases:
        query += " GROUP BY " + ", ".join(group_by_aliases)

    with _engine.connect() as conn:
        agg_df = pd.read_sql(text(query), conn, params=params)

    for c in (STAT_SUM, STAT_COUNT, STAT_MIN, STAT_MAX):
        agg_df[c] = pd.to_numeric(agg_df[c], errors="coerce")
    return agg_df


def _combine_group_stats(df: pd.DataFrame, group_cols: list[str], agg_func: str) -> pd.Series:
    """
    Combine the pre-aggregated STAT_SUM/STAT_COUNT/STAT_MIN/STAT_MAX columns
    across `group_cols` -- a grouping that may be *coarser* than the leaf
    granularity the stats were computed at, i.e. a subtotal or a margin --
    into ONE correctly-combined value per group for `agg_func`.

    sum/count/min/max all combine correctly by re-applying the SAME
    operation across the leaf-level stats (sum of sums, sum of counts, min
    of mins, max of maxes). mean is reconstructed as sum/count -- NEVER as
    an average of the leaf-level means, which silently gives the wrong
    number the instant the underlying group sizes differ (the classic
    "average of ratios" vs "ratio of sums" trap). `group_cols=[]` combines
    across the ENTIRE table into a single-row Series (used for grand
    totals / the corner cell).
    """
    if group_cols:
        g = df.groupby(group_cols, dropna=False)
    else:
        g = df.groupby(np.zeros(len(df), dtype=int))

    if agg_func == "sum":
        return g[STAT_SUM].sum().astype(float)
    if agg_func == "count":
        return g[STAT_COUNT].sum().astype(float)
    if agg_func == "mean":
        sum_s = g[STAT_SUM].sum()
        count_s = g[STAT_COUNT].sum()
        return (sum_s / count_s.replace(0, np.nan)).fillna(0.0).astype(float)
    if agg_func == "min":
        return g[STAT_MIN].min().astype(float)
    if agg_func == "max":
        return g[STAT_MAX].max().astype(float)
    raise ValueError(f"Unsupported agg_func: {agg_func}")


def _expand_date_fields(
    df: pd.DataFrame,
    fields: list[str],
    column_kind_map: dict[str, str],
    label_map: dict[tuple[str, Any], str],
) -> tuple[pd.DataFrame, list[str]]:
    """
    Mirrors Excel's automatic date grouping: a date/timestamp Rows/Columns
    field expands into two levels -- Month ('Mon-YY') above the exact Date
    ('DD-Mon-YY'). The bucketed columns themselves ("__<field>__month" /
    "__<field>__date", both plain 'YYYY-MM' / 'YYYY-MM-DD' text) are now
    computed server-side by `fetch_pivot_source_data`'s GROUP BY -- this
    function's job is just to know the resulting effective field names and
    build the pretty display `label_map` from those already-bucketed text
    columns. Non-date fields pass through unchanged. `df` is returned
    as-is (no mutation needed any more); kept in the signature/return so
    every existing call site keeps working unchanged.
    """
    expanded: list[str] = []
    for f in fields:
        if column_kind_map.get(f) == "date":
            month_field, date_field = f"__{f}__month", f"__{f}__date"
            if month_field in df.columns:
                month_keys = df[month_field].dropna().unique()
                if len(month_keys):
                    disp = pd.to_datetime(month_keys, format="%Y-%m").strftime("%b-%y")
                    for k, d in zip(month_keys, disp):
                        label_map[(month_field, k)] = d
            if date_field in df.columns:
                date_keys = df[date_field].dropna().unique()
                if len(date_keys):
                    disp = pd.to_datetime(date_keys, format="%Y-%m-%d").strftime("%d-%b-%y")
                    for k, d in zip(date_keys, disp):
                        label_map[(date_field, k)] = d
            expanded.extend([month_field, date_field])
        else:
            expanded.append(f)
    return df, expanded


def _make_dims_cache(df: pd.DataFrame, agg_func: str):
    """
    Returns `series_for(dims)` -- combines df's pre-aggregated stats grouped
    by exactly `dims` (a tuple of field names, possibly empty), the FIRST
    time that exact combination of dims is asked for, and returns the
    cached result instantly every time after.

    This is the fix for a pivot with many Rows/Columns/Values fields
    freezing the page: naively, computing every single CELL of a pivot by
    re-scanning the whole (already-aggregated) source table with a fresh
    boolean mask is correct but O(cells x table_size) -- cheap-looking per
    call, but a pivot with, say, 40 rows x 20 columns is already 800 full
    table scans, and a deep Rows hierarchy or a Field Efficiency table's
    numerator+denominator doubles or triples that. In reality there are
    only ever a HANDFUL of distinct `dims` combinations across an entire
    pivot -- one per (row-depth, column-depth) pairing, not one per cell --
    so grouping by each distinct combination ONCE and looking up individual
    values from that small, already-computed result is dramatically faster,
    with the speedup growing the more cells the pivot has.
    """
    cache: dict[tuple[str, ...], pd.Series | float] = {}

    def series_for(dims: tuple[str, ...]) -> pd.Series | float:
        if dims not in cache:
            if not dims:
                cache[dims] = float(_combine_group_stats(df, [], agg_func).iloc[0])
            else:
                series = _combine_group_stats(df, list(dims), agg_func)
                if len(dims) > 1:
                    series = series.sort_index()  # faster repeated MultiIndex lookups
                cache[dims] = series
        return cache[dims]

    return series_for


def _lookup_dims_value(series_for, dims: tuple[str, ...], key: tuple[Any, ...]) -> float:
    """Look up one cell's value from a `series_for`-cached grouping. `dims`
    empty means the corner/grand-total case (a single scalar, no lookup
    needed); `len(dims) == 1` means the cached Series is keyed by plain
    scalars (pandas' own behavior for a single-column groupby), not
    1-tuples, so that case is unwrapped to match."""
    series_or_scalar = series_for(dims)
    if not dims:
        return float(series_or_scalar)
    lookup_key = key[0] if len(dims) == 1 else key
    val = series_or_scalar.get(lookup_key, 0.0)
    return float(val) if val is not None else 0.0


def _month_sort_key(value: Any) -> tuple[int, int, str] | None:
    """
    If `value` looks like a month name someone typed/stored as plain text —
    'January', 'Jan', 'January 2024', 'Jan-24', 'Jan 24' — return a sort key
    that orders by CALENDAR year (January -> December), year ascending when
    present. Returns None for anything that doesn't look like a month
    value, so normal columns are completely unaffected.
    """
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text:
        return None
    match = re.match(r"^([A-Za-z]+)[\s\-,]*(\d{2,4})?$", text)
    if not match:
        return None
    month_part, year_part = match.group(1), match.group(2)
    month_num = _MONTH_FULL_TO_NUM.get(month_part.title()) or _MONTH_ABBR_TO_NUM.get(month_part[:3].title())
    if month_num is None:
        return None
    if year_part:
        year_num = int(year_part)
        if year_num < 100:  # 2-digit year, e.g. "24" -> 2024
            year_num += 2000
    else:
        # No year at all -- every value is in the same single cycle.
        year_num = 0

    return (year_num, month_num, text)


def _ordered_unique(series: pd.Series) -> list[Any]:
    """
    Distinct non-null values, sorted when possible. If EVERY value in the
    column looks like a month name (plain text, not a real date/timestamp
    column — those are already handled separately and sort correctly), sort
    chronologically (January -> December, and by year first if present)
    instead of alphabetically. Otherwise falls back to a normal sort, or
    first-seen order if the values aren't sortable at all.
    """
    vals = [v for v in pd.unique(series) if pd.notna(v)]
    if not vals:
        return vals
    month_keys = [_month_sort_key(v) for v in vals]
    if all(k is not None for k in month_keys):
        return [v for _, v in sorted(zip(month_keys, vals), key=lambda pair: pair[0])]
    try:
        return sorted(vals)
    except TypeError:
        return vals


def _reorder_axis_chronologically(idx: pd.Index) -> pd.Index:
    """
    Reorder a (possibly MultiIndex) pandas axis so any month-like level
    sorts chronologically instead of alphabetically — the exact same rule
    `_ordered_unique` applies to the main Summary Matrix — while pinning a
    literal 'Total' entry at that level to the very end. Non-month levels
    keep their existing order untouched. Used for tables (like the Field
    Efficiency Table) that are built via pandas groupby/pivot_table instead
    of the row/column tree builder, so they still need this applied
    separately after the fact.
    """
    is_multi = isinstance(idx, pd.MultiIndex)
    n_levels = idx.nlevels if is_multi else 1
    tuples = list(idx) if is_multi else [(v,) for v in idx]

    level_ranks: list[dict[Any, int]] = []
    for level in range(n_levels):
        non_total_vals = [t[level] for t in tuples if t[level] not in ("Total", "")]
        ordered_vals = _ordered_unique(pd.Series(non_total_vals)) if non_total_vals else []
        level_ranks.append({v: i for i, v in enumerate(ordered_vals)})

    def _sort_key(t: tuple) -> tuple:
        key = []
        for level in range(n_levels):
            v = t[level]
            key.append((1, 0) if v in ("Total", "") else (0, level_ranks[level].get(v, 0)))
        return tuple(key)

    order = sorted(range(len(tuples)), key=lambda i: _sort_key(tuples[i]))
    new_tuples = [tuples[i] for i in order]

    if is_multi:
        return pd.MultiIndex.from_tuples(new_tuples, names=idx.names)
    return pd.Index([t[0] for t in new_tuples], name=idx.name)


def _build_row_entries(
    df: pd.DataFrame,
    row_fields: list[str],
    label_map: dict[tuple[str, Any], str],
    prefix: tuple[Any, ...] = (),
    indent: int = 0,
) -> list[dict[str, Any]]:
    """
    Excel-style row tree: a subtotal row for each parent group appears
    BEFORE its children (e.g. 'EAST' subtotal, then 'EAST_1', 'EAST_2'
    indented beneath it), recursing to arbitrary depth.

    Vectorized: this used to recurse into `df`, re-scanning the ENTIRE
    table with a fresh boolean mask at EVERY node (O(nodes x table_size)) --
    fine for a shallow hierarchy with a few dozen groups, but with a deep
    Rows hierarchy (e.g. zone > sub_zone > main_region > sub_region >
    area_office > branch) and tens of thousands of leaf combinations, the
    node count itself explodes and the whole page freezes rebuilding the
    tree before a single cell of the actual table is computed.

    Instead: rank every level's DISTINCT values once (same order
    `_ordered_unique` would have produced -- a value's relative order never
    depends on which parent it sits under, so one global ranking is exactly
    equivalent to re-deriving it per node), sort the table's distinct
    row-field combinations by those ranks, and walk that sorted table ONCE.
    That single sorted pass is precisely a depth-first, pre-order walk of
    the same tree the recursion would have produced -- so a linear scan
    (plus one sort) replaces what used to be a per-node rescan of the whole
    table.
    """
    n_levels = len(row_fields)
    if n_levels == 0:
        return []

    uniq = df[row_fields].drop_duplicates()
    if uniq.empty:
        return []

    # Rank each level's values once, globally, in the exact order
    # _ordered_unique would give them (chronological for month-like
    # columns, sorted otherwise).
    rank_cols = []
    for f in row_fields:
        order = _ordered_unique(df[f])
        rank_map = {v: i for i, v in enumerate(order)}
        rank_cols.append(uniq[f].map(rank_map).to_numpy())

    # np.lexsort's LAST key is primary, so reverse to make level 0 (zone)
    # the major sort key and the deepest level the minor one.
    sort_idx = np.lexsort(rank_cols[::-1])
    sorted_vals = uniq.iloc[sort_idx][row_fields].to_numpy()

    entries: list[dict[str, Any]] = []
    prev_row: tuple[Any, ...] | None = None
    for row_vals in sorted_vals:
        row_vals = tuple(row_vals)
        # Find the shallowest level where this row's path diverges from the
        # previous one -- everything from there down is a "new branch" and
        # needs fresh group/leaf entries emitted for it. NaN never equals
        # itself, so a NaN value naturally always counts as "changed".
        changed_at = 0
        if prev_row is not None:
            while changed_at < n_levels and row_vals[changed_at] == prev_row[changed_at]:
                changed_at += 1
        for level in range(changed_at, n_levels):
            v = row_vals[level]
            if pd.isna(v):
                # A missing value at this level means this row can't
                # contribute an entry here or any deeper -- same as the
                # original mask-based recursion, where _ordered_unique
                # simply never includes NaN as a value to recurse into.
                break
            disp = label_map.get((row_fields[level], v), str(v))
            is_last_level = level == n_levels - 1
            entries.append({
                "label": disp,
                "prefix": row_vals[: level + 1],
                "indent": level,
                "is_group": not is_last_level,
            })
        prev_row = row_vals
    return entries


def _build_col_entries(
    df: pd.DataFrame,
    col_fields: list[str],
    label_map: dict[tuple[str, Any], str],
    prefix: tuple[Any, ...] = (),
) -> list[dict[str, Any]]:
    """
    Excel-style column tree: a group's children come first, followed by that
    group's own '<value> Total' subtotal column (e.g. Jun, Jun Total, Jul,
    Jul Total, then STAGE_3 Total) — the mirror image of the row ordering.
    """
    level = len(prefix)
    if level >= len(col_fields):
        return []
    field = col_fields[level]
    mask = pd.Series(True, index=df.index)
    for f, v in zip(col_fields[:level], prefix):
        mask &= df[f] == v
    values = _ordered_unique(df.loc[mask, field])
    is_last_level = level == len(col_fields) - 1

    entries: list[dict[str, Any]] = []
    for v in values:
        new_prefix = prefix + (v,)
        if is_last_level:
            disp = label_map.get((field, v), str(v))
            entries.append({"label": disp, "prefix": new_prefix, "is_total": False})
        else:
            entries.extend(_build_col_entries(df, col_fields, label_map, new_prefix))
            disp = label_map.get((field, v), str(v))
            entries.append({"label": f"{disp} Total", "prefix": new_prefix, "is_total": True})
    return entries


def _pretty_prefix(
    prefix: tuple[Any, ...], fields: list[str], label_map: dict[tuple[str, Any], str]
) -> tuple[str, ...]:
    """Map each raw value in a prefix tuple to its display label, one per field."""
    return tuple(label_map.get((fields[i], v), str(v)) for i, v in enumerate(prefix))




@ttl_cache(ttl=300)
def build_excel_style_pivot(
    source_df: pd.DataFrame,
    row_cols: list[str],
    col_cols: list[str],
    value_col: str,
    agg_func: str,
    column_kind_map: dict[str, str],
    in_crores: bool,
    right_total_mode: str = "grand_total",
) -> tuple[pd.DataFrame, list[bool], list[bool]]:
    """
    Build a true Excel-style PivotTable: per-level row AND column subtotals
    (not just one grand-total margin), date fields auto-grouped into a
    Month level above the exact date, and values optionally shown in
    Crores. Returns (display_df, group_row_flags, total_col_flags) — the
    two flag lists are positional (aligned to display_df's row/column
    order), used only for highlighting subtotal/Total rows & columns.

    `right_total_mode` controls what appears at the far right of the header
    when 2+ Columns fields are selected (e.g. Stage > Month):
      - "grand_total"        -> single overall "Grand Total" column (default,
                                 original behaviour).
      - "deepest_field_total" -> one "<value> Total" column per distinct value
                                 of the *last* Columns field (e.g. one per
                                 month), each summed across every other
                                 Columns field (e.g. across all stages) --
                                 no single overall Grand Total column.
      - "both"                -> the per-deepest-field totals, followed by
                                 one overall Grand Total at the very end.
    With fewer than 2 Columns fields there's nothing to "sum across", so
    this always falls back to the plain single Grand Total column.
    """
    label_map: dict[tuple[str, Any], str] = {}
    work_df, row_fields = _expand_date_fields(source_df, row_cols, column_kind_map, label_map)
    work_df, col_fields = _expand_date_fields(work_df, col_cols, column_kind_map, label_map)

    row_entries = _build_row_entries(work_df, row_fields, label_map)
    row_entries.append({"label": "Grand Total", "prefix": (), "indent": 0, "is_group": True, "is_grand_total": True})

    if col_fields:
        col_entries = _build_col_entries(work_df, col_fields, label_map)
        col_depth = len(col_fields)

        # Dynamic "per-deepest-field" totals (e.g. one "Jul-26 Total" column
        # per month, summed across every Stage) -- only meaningful with 2+
        # Columns fields; with just one field there's no "other" field left
        # to sum across, so this list simply stays empty in that case.
        deepest_field_entries: list[dict[str, Any]] = []
        if right_total_mode in ("deepest_field_total", "both") and col_depth >= 2:
            deepest_field = col_fields[-1]
            for v in _ordered_unique(work_df[deepest_field]):
                disp = label_map.get((deepest_field, v), str(v))
                deepest_field_entries.append({
                    "label": f"{disp} Total",
                    "prefix": (v,),
                    "is_total": True,
                    # Aggregate by matching ONLY this field -- ignoring every
                    # other Columns field entirely, which is what makes this
                    # a total *across* stages rather than a per-stage subtotal.
                    "mask_fields": [deepest_field],
                })

        # Fall back to (or additionally include) the classic single overall
        # Grand Total column -- always included if the per-field totals
        # above ended up empty for any reason (e.g. mode requested but only
        # 1 Columns field was actually picked).
        grand_total_entries: list[dict[str, Any]] = []
        if right_total_mode in ("grand_total", "both") or not deepest_field_entries:
            grand_total_entries = [{"label": "Grand Total", "prefix": (), "is_total": True, "is_grand_total": True}]

        col_entries = col_entries + deepest_field_entries + grand_total_entries

        col_tuples: list[tuple[str, ...]] = []
        for ce in col_entries:
            if ce.get("is_grand_total"):
                tup = ("Grand Total",) + ("",) * (col_depth - 1)
            elif ce.get("mask_fields"):
                # Per-deepest-field total (e.g. "Jul-26 Total") -- shown as a
                # single top-level header label, same visual treatment as
                # Grand Total, since it cuts across every other Columns
                # field rather than nesting under one specific value of them.
                tup = (ce["label"],) + ("",) * (col_depth - 1)
            else:
                depth_used = len(ce["prefix"])
                pretty = _pretty_prefix(ce["prefix"], col_fields[:depth_used], label_map)
                if ce["is_total"]:
                    pretty = pretty[:-1] + (f"{pretty[-1]} Total",)
                tup = pretty + ("",) * (col_depth - len(pretty))
            col_tuples.append(tup)
        columns_index = pd.MultiIndex.from_tuples(col_tuples) if col_depth > 1 else pd.Index([t[0] for t in col_tuples])
        total_col_flags = [bool(ce.get("is_total") or ce.get("is_grand_total")) for ce in col_entries]
    else:
        col_entries = [{"label": f"{agg_func.title()} of {value_col}", "prefix": (), "is_total": False}]
        columns_index = pd.Index([col_entries[0]["label"]])
        total_col_flags = [False]

    divisor = CRORE if (in_crores and agg_func != "count") else 1
    series_for = _make_dims_cache(work_df, agg_func)

    row_labels: list[str] = []
    group_row_flags: list[bool] = []
    data_rows: list[list[float]] = []
    for re_ in row_entries:
        indent_txt = "    " * re_["indent"]
        if re_.get("is_grand_total"):
            label = "Grand Total"
        elif re_["is_group"]:
            label = f"{indent_txt}{re_['label']}"
        else:
            # Only show the child-arrow when this row is actually nested under a
            # parent group row (indent > 0). A single-level Rows selection has
            # no parent, so it should read as a plain flat list, e.g. "EAST",
            # not "↳ EAST".
            arrow = "↳ " if re_["indent"] > 0 else ""
            label = f"{indent_txt}{arrow}{re_['label']}"
        row_labels.append(label)
        group_row_flags.append(bool(re_["is_group"]))

        row_prefix = re_["prefix"]
        row_vals = []
        for ce in col_entries:
            col_fields_for_cell = ce.get("mask_fields", col_fields)
            col_prefix = ce["prefix"]
            dims = tuple(row_fields[: len(row_prefix)]) + tuple(col_fields_for_cell[: len(col_prefix)])
            key = tuple(row_prefix) + tuple(col_prefix)
            row_vals.append(_lookup_dims_value(series_for, dims, key) / divisor)
        data_rows.append(row_vals)

    display_df = pd.DataFrame(
        data_rows,
        index=pd.Index(row_labels, name=" / ".join(row_cols)),
        columns=columns_index,
    )
    return display_df, group_row_flags, total_col_flags


_AGGRID_NUM_FORMATTERS: dict[str, "JsCode"] = {}
_aggrid_key_counter = itertools.count()


def _split_matrix_row_label(raw_label: str) -> tuple[int, str]:
    """
    Recovers (depth, clean_text) from a Summary-Matrix row label built by
    `build_excel_style_pivot`'s compact-form indentation convention: N
    groups of 4 leading spaces per nesting level, and a "↳ " marker on leaf
    rows that are actually nested under a parent (e.g. "        ↳ EAST_1"
    is depth 2, clean text "EAST_1"). A label with no leading spaces/arrow
    -- a top-level row, or "Grand Total" -- is depth 0, returned unchanged.
    This only *reads* the existing convention; it doesn't change how labels
    are built, so the on-screen table, DPD-bucket, and Flow-Retention logic
    (which also read this same convention) are completely unaffected.
    """
    stripped = raw_label.lstrip(" ")
    depth = (len(raw_label) - len(stripped)) // 4
    if stripped.startswith("↳ "):
        stripped = stripped[2:]
    return depth, stripped




def dataframe_to_formatted_excel_bytes(
    df: pd.DataFrame,
    group_row_flags: list[bool] | None = None,
    total_col_flags: list[bool] | None = None,
    value_kind: str = "number",  # "number" | "count" | "percent"
    table_style: str = "matrix",  # "matrix" (Summary Matrix rules) | "efficiency" (Field Efficiency rules)
    sheet_name: str = "Report",
) -> bytes:
    """
    Export a (possibly MultiIndex-column) pivot/summary DataFrame to a
    properly formatted .xlsx -- unlike a flattened CSV, an outer header like
    "Sum of March 23" is written as a real merged cell spanning BOUNCED /
    CLEARED beneath it, instead of silently disappearing. Subtotal/Total
    rows and columns get the same bold/shading treatment shown on screen,
    and percent tables (Field Efficiency) are written exactly as displayed,
    e.g. "18.80%", via a custom number format -- not divided by 100 again.

    Built on `xlsxwriter` rather than cell-by-cell `openpyxl`: header labels
    are written via `merge_range` (one call per merged span, not one call
    per underlying cell), and every `Format` (font/fill/border/number-format
    combination) is created ONCE and cached/reused across every cell that
    needs it, instead of a fresh `Font`/`PatternFill` object per cell --
    object construction, not the write call itself, was the actual cost
    driver in the original openpyxl version at pivot-table sizes. Cached
    with @st.cache_data on top, so re-clicking the download button after a
    widget-only rerun serves the same bytes instantly rather than rebuilding
    the workbook.
    """
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "default_date_format": "dd-mmm-yy"})
    ws = wb.add_worksheet((sheet_name or "Report")[:31])

    # -- cached format factory: same (props) tuple always returns the same
    # Format object instead of allocating a new one per cell. -- #
    _fmt_cache: dict[tuple, Any] = {}

    def get_fmt(**props) -> Any:
        props = {k: v for k, v in props.items() if v is not None}
        key = tuple(sorted(props.items()))
        f = _fmt_cache.get(key)
        if f is None:
            f = wb.add_format(props)
            _fmt_cache[key] = f
        return f

    BORDER_COLOR = "#D0D7DE"
    base_border = {"border": 1, "border_color": BORDER_COLOR}

    has_multi_cols = isinstance(df.columns, pd.MultiIndex)
    col_levels = df.columns.nlevels if has_multi_cols else 1
    index_is_multi = isinstance(df.index, pd.MultiIndex)
    n_index_cols = df.index.nlevels if index_is_multi else 1
    raw_names = list(df.index.names) if index_is_multi else [df.index.name]
    index_names = [n if n else f"Row {i + 1}" for i, n in enumerate(raw_names)]

    header_fmt = get_fmt(
        bold=True, font_color=f"#{_XL_HEADER_FONT}", font_size=11, font_name="Calibri",
        bg_color=f"#{_XL_HEADER_FILL}", align="center", valign="vcenter", text_wrap=True, **base_border,
    )

    # --- Index (row-key) header, spanning every header row --- #
    for i, name in enumerate(index_names):
        if col_levels > 1:
            ws.merge_range(0, i, col_levels - 1, i, name, header_fmt)
        else:
            ws.write(0, i, name, header_fmt)

    data_start_col = n_index_cols

    # --- Value column headers, merged per level so parent labels (e.g.
    # "Sum of March 23") correctly span every child column beneath them --- #
    col_tuples = list(df.columns) if has_multi_cols else [(c,) for c in df.columns]
    for level in range(col_levels):
        n = len(col_tuples)
        i = 0
        while i < n:
            j = i
            key = col_tuples[i][: level + 1]
            while j + 1 < n and col_tuples[j + 1][: level + 1] == key:
                j += 1
            label = col_tuples[i][level]
            col0 = data_start_col + i
            col1 = data_start_col + j
            text = "" if label is None else str(label)
            if col1 > col0:
                ws.merge_range(level, col0, level, col1, text, header_fmt)
            else:
                ws.write(level, col0, text, header_fmt)
            i = j + 1

    # --- Body --- #
    num_fmt_str = {'percent': '0.00"%"', 'count': '#,##0'}.get(value_kind, '#,##0.00')

    flat_index_tuples = list(df.index) if index_is_multi else [(v,) for v in df.index]
    n_rows = len(df)
    header_rows = col_levels

    def is_grand_total_row(r: int) -> bool:
        # Two conventions coexist in this codebase: the Summary Matrix's
        # single compact-string column has the label last (only) element;
        # a genuine MultiIndex Total row (Field Efficiency, flat pivot) puts
        # "Total" FIRST with blanks after, e.g. ("Total", "", ""). Check both
        # ends so this works correctly under either shape.
        tup = flat_index_tuples[r]
        first_label = str(tup[0]).strip()
        last_label = str(tup[-1]).strip()
        return first_label in ("Grand Total", "Total") or last_label in ("Grand Total", "Total")

    # Only the Summary Matrix's compact-form "single Row Labels column"
    # shape (one index column, not a MultiIndex) carries the 4-space /
    # "↳ " nesting convention -- that's what unlocks per-level colour,
    # native cell indent, and real Excel row grouping (+/- outline
    # buttons) below. Every other export (flat pivot with a MultiIndex
    # row index, Field Efficiency, Flow Retention) is left exactly as
    # it already was.
    use_row_hierarchy_style = (not index_is_multi) and (n_index_cols == 1) and (group_row_flags is not None)
    # Field Efficiency (and any other export) whose Rows selection has 2+
    # fields produces a genuine pandas MultiIndex -- every row already shows
    # the full parent-to-child combination (e.g. EAST / E-North / EN-1), just
    # with the parent values repeated on every single row. This gives that
    # same shape the same treatment as the Summary Matrix: repeated parent
    # labels blanked out, each level colour-coded, and real Excel row
    # grouping -- computed purely from how many leading levels each row
    # shares with the row above it, since (unlike the Summary Matrix) there
    # are no separate one-per-level subtotal rows to key off of here.
    use_multiindex_hierarchy_style = index_is_multi and n_index_cols > 1

    if use_row_hierarchy_style or use_multiindex_hierarchy_style:
        # Summary/subtotal rows sit ABOVE their detail rows in this table
        # (e.g. "EAST" then "EAST_1"/"EAST_2" beneath it) -- symbols_below=False
        # tells Excel to draw the +/- collapse toggle next to that summary
        # row rather than assuming a summary-at-the-bottom layout.
        ws.outline_settings(symbols_below=False, symbols_right=True)

    # For the genuine-MultiIndex case, precompute how many LEADING levels
    # each row shares with the row directly above it -- vectorized via numpy
    # (compare each row's tuple to the row above's, all at once) rather than
    # a nested per-row/per-level Python loop.
    multi_common_prefix: list[int] = []
    if use_multiindex_hierarchy_style:
        idx_arr = np.array([[t[lvl] for lvl in range(n_index_cols)] for t in flat_index_tuples], dtype=object)
        grand_total_mask = np.array([is_grand_total_row(r) for r in range(n_rows)])
        prev_arr = np.empty_like(idx_arr)
        prev_arr[0] = None
        prev_arr[1:] = idx_arr[:-1]
        eq = (idx_arr == prev_arr)
        # cumulative "all True so far" across levels -> count of leading
        # matching levels for each row, in one vectorized pass.
        common = eq.cumprod(axis=1).sum(axis=1)
        common[grand_total_mask] = 0
        reset_after = np.zeros(n_rows, dtype=bool)
        reset_after[1:] = grand_total_mask[:-1]
        common[reset_after] = 0
        multi_common_prefix = common.astype(int).tolist()

    for r in range(n_rows):
        excel_row = header_rows + r
        is_group = bool(group_row_flags[r]) if group_row_flags is not None and r < len(group_row_flags) else False
        is_total_row = is_grand_total_row(r) or (table_style == "efficiency" and r == n_rows - 1)

        if use_row_hierarchy_style:
            depth, clean_label = _split_matrix_row_label(str(flat_index_tuples[r][0]))
            if depth:
                ws.set_row(excel_row, None, None, {"level": depth})
        elif use_multiindex_hierarchy_style:
            if multi_common_prefix[r]:
                ws.set_row(excel_row, None, None, {"level": multi_common_prefix[r]})
            depth, clean_label = 0, None
        else:
            depth, clean_label = 0, None

        # -- index cells -- #
        for c in range(n_index_cols):
            if use_row_hierarchy_style and c == 0:
                val = clean_label
            elif use_multiindex_hierarchy_style and not is_total_row and c < multi_common_prefix[r]:
                val = ""  # same as the row above at this level -- don't repeat it
            else:
                val = flat_index_tuples[r][c] if c < len(flat_index_tuples[r]) else ""

            if use_multiindex_hierarchy_style:
                level_color = None if is_total_row else _XL_LEVEL_FONT_COLORS[c % len(_XL_LEVEL_FONT_COLORS)]
                idx_fmt = get_fmt(
                    font_name="Calibri", font_size=10.5, bold=is_total_row,
                    font_color=(f"#{level_color}" if level_color else None),
                    bg_color=(f"#{_XL_TOTAL_ROW_FILL}" if is_total_row else None), **base_border,
                )
            elif use_row_hierarchy_style:
                level_color = _XL_LEVEL_FONT_COLORS[depth % len(_XL_LEVEL_FONT_COLORS)] if not is_total_row else None
                idx_fmt = get_fmt(
                    font_name="Calibri", font_size=10.5, bold=(is_total_row or is_group),
                    font_color=(f"#{level_color}" if level_color else None),
                    bg_color=(f"#{_XL_GROUP_ROW_FILL}" if (is_group and not is_total_row) else None),
                    indent=depth, valign="vcenter", **base_border,
                )
            else:  # efficiency table
                idx_fmt = get_fmt(
                    font_name="Calibri", font_size=10.5, bold=is_total_row,
                    bg_color=(f"#{_XL_TOTAL_ROW_FILL}" if is_total_row else None), **base_border,
                )
            ws.write(excel_row, c, val, idx_fmt)

        # -- data cells -- #
        row_vals = df.iloc[r].to_numpy()
        for j in range(df.shape[1]):
            val = row_vals[j]
            is_total_col = bool(total_col_flags[j]) if total_col_flags is not None and j < len(total_col_flags) else False

            if table_style == "matrix":
                # Matches the on-screen rule: Grand Total row = bold only (no
                # fill); subtotal/group rows = blue fill + bold; Total
                # columns = grey fill + bold, layered on top of either.
                bg = None
                if is_group and not is_total_row:
                    bg = _XL_GROUP_ROW_FILL
                if is_total_col:
                    bg = _XL_TOTAL_COL_FILL
                data_fmt = get_fmt(
                    font_name="Calibri", font_size=10.5, bold=(is_total_row or is_group or is_total_col),
                    bg_color=(f"#{bg}" if bg else None), num_format=num_fmt_str, **base_border,
                )
            else:  # efficiency table: Total row/col shaded, corner a shade deeper
                bg = None
                if is_total_row and is_total_col:
                    bg = _XL_CORNER_FILL
                elif is_total_row:
                    bg = _XL_TOTAL_ROW_FILL
                elif is_total_col:
                    bg = _XL_TOTAL_COL_FILL
                data_fmt = get_fmt(
                    font_name="Calibri", font_size=10.5, bold=(is_total_row or is_total_col),
                    bg_color=(f"#{bg}" if bg else None), num_format=num_fmt_str, **base_border,
                )

            if pd.isna(val):
                ws.write_blank(excel_row, data_start_col + j, None, data_fmt)
            else:
                ws.write_number(excel_row, data_start_col + j, float(val), data_fmt)

    ws.freeze_panes(header_rows, data_start_col)

    # -- column widths -- #
    for c in range(n_index_cols):
        longest = max([len(str(index_names[c]))] + [len(str(t[c])) for t in flat_index_tuples if c < len(t)])
        ws.set_column(c, c, min(38, max(12, longest + 2)))
    for j in range(df.shape[1]):
        header_len = max(len(str(x)) for x in col_tuples[j]) if col_tuples[j] else 8
        ws.set_column(data_start_col + j, data_start_col + j, min(22, max(11, header_len + 2)))

    wb.close()
    return buf.getvalue()


def _effective_col_level(col_cols: list[str], target_field: str, column_kind_map: dict[str, str]) -> int:
    """
    Map a field name in `col_cols` to its column-level index in the *effective*
    (date-expansion-aware) column tree built by `build_excel_style_pivot`. A
    date/timestamp field expands into TWO levels (Month, then exact Date) via
    `_expand_date_fields`, so any field listed after one in `col_cols` is
    shifted by an extra level for each date field preceding it. Non-date
    fields (like a DPD slab column) occupy exactly one level.
    """
    level = 0
    for f in col_cols:
        if f == target_field:
            return level
        level += 2 if column_kind_map.get(f) == "date" else 1
    raise ValueError(f"'{target_field}' is not one of the selected Columns fields: {col_cols}")


def add_dpd_buckets_to_excel_pivot(
    display_df: pd.DataFrame, slab_level: int, only_show_buckets: bool = False
) -> pd.DataFrame:
    """
    Append '1+', '30+', '90+' cumulative DPD macro-bucket column groups to a
    display_df produced by `build_excel_style_pivot` / `render_excel_style_pivot_table`
    (the Excel-style Summary Matrix) -- the table that already contains
    per-slab subtotal columns like '1-29 Total' and a 'Grand Total' column.

    '1+'  = sum of '1-29' + '30-59' + '60-89' + '90+'
    '30+' = sum of '30-59' + '60-89' + '90+'
    '90+' = sum of '90+' alone -- when the raw slabs are being kept
            (only_show_buckets=False) this is IDENTICAL to the existing raw
            '90+' slab column, so it's skipped rather than added as a
            confusing duplicate column with the same label. When the raw
            slabs are being dropped (only_show_buckets=True) there's no
            collision any more, so '90+' is added normally.

    `slab_level` is the column level (from `_effective_col_level`) holding the
    raw DPD slab values. When 2+ Columns fields are selected, the DPD slab
    field currently must be the FIRST Columns field (slab_level == 0) so the
    new '<bucket> Total' columns line up with the existing '<slab> Total'
    subtotal convention; with just one Columns field (flat columns) any
    position works since there's nothing else to reconstruct.

    only_show_buckets : bool, default False
        If True, the raw slab columns ('1-29', '30-59', '60-89', '90+') and
        their '<slab> Total' subtotals are removed from the output entirely,
        leaving only '1+' / '30+' / '90+' (plus 'Grand Total' and any other
        unrelated columns untouched).

    Raises ValueError (caught and shown as a friendly warning by the caller)
    if the shape isn't supported, or if no raw DPD-slab columns are found.
    """
    columns = display_df.columns
    is_multi = isinstance(columns, pd.MultiIndex)
    n_levels = columns.nlevels if is_multi else 1

    if n_levels > 1 and slab_level != 0:
        raise ValueError(
            "Adding DPD buckets currently requires the DPD slab field to be the "
            "FIRST field in your Columns list when 2+ Columns fields are selected. "
            "Reorder your Columns selection in the sidebar so the DPD slab field "
            "comes first, then try again."
        )

    known_slabs = {"1-29", "30-59", "60-89", "90+"}
    bucket_defs = {
        "1+": {"1-29", "30-59", "60-89", "90+"},
        "30+": {"30-59", "60-89", "90+"},
        "90+": {"90+"},
    }

    tuples = [t if isinstance(t, tuple) else (t,) for t in columns]

    def _slab_val(t: tuple) -> Any:
        return t[slab_level] if is_multi else t[0]

    def _is_grand_total(t: tuple) -> bool:
        return any(str(v) == "Grand Total" for v in t)

    def _is_group_total(t: tuple) -> bool:
        return any(isinstance(v, str) and v.endswith(" Total") and v != "Grand Total" for v in t)

    leaf_tuples = [
        t for t in tuples
        if not _is_grand_total(t) and not _is_group_total(t) and str(_slab_val(t)) in known_slabs
    ]
    if not leaf_tuples:
        raise ValueError(
            "No raw DPD-slab columns ('1-29', '30-59', '60-89', '90+') were found in the "
            "selected field -- double-check you picked the right Columns field."
        )

    existing_leaf_slabs = {str(_slab_val(t)) for t in leaf_tuples}

    def _other_key(t: tuple) -> tuple:
        return tuple(v for i, v in enumerate(t) if i != slab_level)

    seen_other: list[tuple] = []
    seen_set: set = set()
    for t in leaf_tuples:
        ok = _other_key(t)
        if ok not in seen_set:
            seen_set.add(ok)
            seen_other.append(ok)

    new_frames = []
    for bucket_name, slabs in bucket_defs.items():
        # Skip a bucket that's identical to an existing raw slab column --
        # but only when that raw slab column is still going to be shown.
        if not only_show_buckets and bucket_name in existing_leaf_slabs and slabs == {bucket_name}:
            continue

        bucket_leaf = [t for t in leaf_tuples if str(_slab_val(t)) in slabs]
        if not bucket_leaf:
            continue

        block_cols: list[Any] = []
        block_data: dict[Any, pd.Series] = {}
        for ok in seen_other:
            matching = [t for t in bucket_leaf if _other_key(t) == ok]
            if not matching:
                continue
            # display_df.columns is a flat Index (not MultiIndex) whenever
            # col_depth == 1, so selecting must use the plain scalar labels,
            # not the internal (label,)-wrapped tuples used everywhere else.
            select_keys = matching if is_multi else [t[0] for t in matching]
            summed = display_df[select_keys].sum(axis=1)
            template = list(matching[0])
            template[slab_level] = bucket_name
            new_tuple = tuple(template) if is_multi else template[0]
            block_data[new_tuple] = summed
            block_cols.append(new_tuple)

        block_df = pd.DataFrame(block_data)[block_cols]

        if n_levels > 1:
            total_tuple = (f"{bucket_name} Total",) + ("",) * (n_levels - 1)
            block_df[total_tuple] = block_df.sum(axis=1)

        new_frames.append(block_df)

    if not new_frames:
        return display_df.copy()

    combined_new = pd.concat(new_frames, axis=1)

    if only_show_buckets:
        # Keep ONLY the new 1+/30+/90+ bucket columns -- drop the raw slabs,
        # Grand Total, and anything else in this field (e.g. a "non-
        # delinquency" category and its own subtotal), since the user wants
        # a table that shows nothing but the cumulative buckets themselves.
        return combined_new

    # Insert the new bucket blocks right before "Grand Total" (if present) so
    # Grand Total stays the visual right-most anchor; otherwise append at the end.
    grand_total_mask = [_is_grand_total(t) for t in tuples]
    if any(grand_total_mask):
        gt_pos = grand_total_mask.index(True)
        before = display_df.iloc[:, :gt_pos]
        after = display_df.iloc[:, gt_pos:]
        result = pd.concat([before, combined_new, after], axis=1)
    else:
        result = pd.concat([display_df, combined_new], axis=1)

    return result


# --------------------------------------------------------------------------- #
# General Field Efficiency Table: each cell as a % of the total across every
# OTHER value of a user-chosen field, holding every other selected Rows/
# Columns field fixed. Unlike a hardcoded "% of Stage" table, this works for
# *any* field the user points it at (a Rows field or a Columns field) — pick
# "staging_ecl" to get % of stage-group total per (Zone, Month); pick "zone"
# instead to get % of zone-group total per (Stage, Month); same formula.
# --------------------------------------------------------------------------- #


def _date_field_leaf(field: str, column_kind_map: dict[str, str]) -> str:
    """
    The Field Efficiency Table groups by whichever field the user picked at
    its full original granularity (matching its pre-optimization behaviour,
    which grouped directly on the raw date/timestamp column). Since raw date
    columns are no longer fetched at all -- only their SQL-bucketed Month/
    Date text columns are -- a date field maps here to its exact-Date leaf
    column ("__<field>__date"), which carries the same information (one
    entry per distinct date) the raw column did.
    """
    return f"__{field}__date" if column_kind_map.get(field) == "date" else field




@ttl_cache(ttl=300)
def build_field_efficiency_pivot(
    source_df: pd.DataFrame,
    row_cols: list[str],
    col_cols: list[str],
    normalize_field: str,
    value_col: str,
    agg_func: str,
    column_kind_map: dict[str, str],
) -> tuple[pd.DataFrame, list[bool]]:
    """
    efficiency(cell) = value(cell) ÷ sum of value across every other value of
    `normalize_field`, holding every OTHER selected Rows/Columns field fixed.

    E.g. with Rows=[zone], Columns=[staging_ecl, value_date], normalize_field=
    "staging_ecl": Stage 3 / East / July = value(Stage3,East,July) ÷
    [value(Stage1,East,July) + value(Stage2,East,July) + value(Stage3,East,July)] -- e.g. "STAGE_3" is the consolidated stage-3 (default) bucket.
    Point it at "zone" instead and it normalizes across zones per (Stage, Month).

    The ROW side mirrors the Summary Matrix exactly: with 2+ Rows fields
    (e.g. Zone, Region, Sub-Zone), a subtotal row for each parent group
    appears ABOVE its children -- "EAST", then "EAST"'s regions beneath it,
    then each region's sub-zones beneath THAT -- instead of one flat row
    per full leaf combination. A group row's % is computed by aggregating
    everything nested under it FIRST, then dividing -- never by averaging
    its children's already-computed percentages (same "ratio of sums, not
    average of ratios" rule the rest of this table already follows).

    Also appends a "Total" row, "Total" column, and "Total"/"Total" corner
    cell -- each computed the same correct way as every other cell.

    `source_df` is the pre-aggregated stats table from `fetch_pivot_source_data`
    (one row per unique Rows x Columns combination, with STAT_SUM/STAT_COUNT/
    STAT_MIN/STAT_MAX columns), NOT raw per-record data.

    Returns (display_df, group_row_flags) -- group_row_flags is positional,
    aligned to display_df's rows, exactly like build_excel_style_pivot's.
    """
    field_row_cols = [_date_field_leaf(f, column_kind_map) for f in row_cols]
    field_col_cols = [_date_field_leaf(f, column_kind_map) for f in col_cols]
    field_normalize = _date_field_leaf(normalize_field, column_kind_map)

    all_dims = list(dict.fromkeys([*field_row_cols, *field_col_cols]))
    if field_normalize not in all_dims:
        raise ValueError(f"'{normalize_field}' must be one of the selected Rows/Columns fields.")

    # Pretty display labels for the row tree -- same idea as build_excel_style_pivot's
    # label_map, populated here (rather than via _expand_date_fields, which also
    # rolls a date field up to a Month level the Field Efficiency Table has never
    # shown) so an exact-date Rows field still displays as "DD-Mon-YY" group labels.
    label_map: dict[tuple[str, Any], str] = {}
    for f in row_cols:
        leaf = _date_field_leaf(f, column_kind_map)
        if column_kind_map.get(f) == "date":
            date_keys = source_df[leaf].dropna().unique()
            if len(date_keys):
                pretty = pd.to_datetime(date_keys, format="%Y-%m-%d", errors="coerce").strftime("%d-%b-%y")
                for k, d in zip(date_keys, pretty):
                    label_map[(leaf, k)] = d

    # One shared cache across every numerator AND denominator grouping below --
    # see _make_dims_cache's docstring for why grouping ONCE per distinct
    # dims combination (not once per cell) is what keeps this table fast
    # even with a deep Rows hierarchy and many Columns values.
    series_for = _make_dims_cache(source_df, agg_func)

    def _eff_grid(row_dims: tuple[str, ...], row_keys: list[tuple[Any, ...]], col_dims: tuple[str, ...], col_keys: list[tuple[Any, ...]]) -> np.ndarray:
        """
        Vectorized replacement for a per-cell `_eff_value` loop: computes an
        entire (len(row_keys) x len(col_keys)) block of the efficiency grid
        in a HANDFUL of pandas/numpy calls -- one `.reindex()` (a single
        vectorized lookup across every row/col combination in the block at
        once, pandas' C-level equivalent of `.div(totals, axis=0)`
        broadcasting for an irregularly-keyed axis) for the numerator, one
        for the denominator, then one elementwise `np.divide` -- instead of
        calling a Python function once per (row, col) cell. Safe to batch
        this way because every entry in `row_keys` shares the same
        `row_dims` (same prefix length -> same grouping), and likewise for
        `col_keys`/`col_dims`, which is exactly how callers below invoke it
        (grouped by row-hierarchy depth).
        """
        numer_dims = tuple(row_dims) + tuple(col_dims)
        pair_keys = [rk + ck for rk in row_keys for ck in col_keys]

        def _lookup_vec(dims: tuple[str, ...], keys: list[tuple[Any, ...]]) -> np.ndarray:
            if not dims:
                return np.full(len(keys), float(series_for(())), dtype=float)
            series = series_for(dims)
            flat_keys = [k[0] for k in keys] if len(dims) == 1 else [tuple(k) for k in keys]
            return series.reindex(flat_keys).fillna(0.0).to_numpy(dtype=float)

        numerator = _lookup_vec(numer_dims, pair_keys)

        if field_normalize in numer_dims:
            idx = numer_dims.index(field_normalize)
            denom_dims = numer_dims[:idx] + numer_dims[idx + 1 :]
            denom_keys = [k[:idx] + k[idx + 1 :] for k in pair_keys]
        else:
            denom_dims, denom_keys = numer_dims, pair_keys
        denominator = _lookup_vec(denom_dims, denom_keys)

        pct = np.divide(
            numerator, denominator,
            out=np.zeros_like(numerator, dtype=float),
            where=denominator != 0,
        ) * 100.0
        return pct.reshape(len(row_keys), len(col_keys))

    # --- Row tree: identical construction to the Summary Matrix --- #
    row_entries = _build_row_entries(source_df, field_row_cols, label_map)
    row_entries.append({"label": "Grand Total", "prefix": (), "indent": 0, "is_group": True, "is_grand_total": True})

    # --- Columns: plain leaf combinations + one Total column (unchanged shape) --- #
    if field_col_cols:
        occurring = (
            set(source_df[field_col_cols].dropna().itertuples(index=False, name=None))
            if len(field_col_cols) > 1
            else {(v,) for v in source_df[field_col_cols[0]].dropna().unique()}
        )
        col_value_lists = [_ordered_unique(source_df[f]) for f in field_col_cols]
        col_prefixes: list[tuple[Any, ...]] = [
            combo for combo in itertools.product(*col_value_lists) if combo in occurring
        ]
        total_col_key = ("Total",) + ("",) * (len(field_col_cols) - 1) if len(field_col_cols) > 1 else "Total"
    else:
        col_prefixes = [()]
        total_col_key = "Total"

    row_labels: list[str] = []
    group_row_flags: list[bool] = []
    for re_ in row_entries:
        indent_txt = "    " * re_["indent"]
        if re_.get("is_grand_total"):
            label = "Grand Total"
        elif re_["is_group"]:
            label = f"{indent_txt}{re_['label']}"
        else:
            arrow = "↳ " if re_["indent"] > 0 else ""
            label = f"{indent_txt}{arrow}{re_['label']}"
        row_labels.append(label)
        group_row_flags.append(bool(re_["is_group"]))

    # Batch every row entry by its prefix LENGTH (= Rows-hierarchy depth):
    # all group rows one level down share the same `row_dims`, so their
    # entire slice of the grid can be computed in one `_eff_grid` call
    # instead of one Python call per individual cell. Bounded by the number
    # of Rows fields (+1 for Grand Total), never by the row/column COUNT.
    n_data_cols = len(col_prefixes) + 1  # + the Total column
    data_matrix = np.zeros((len(row_entries), n_data_cols), dtype=float)
    by_depth: dict[int, list[int]] = {}
    for pos, re_ in enumerate(row_entries):
        by_depth.setdefault(len(re_["prefix"]), []).append(pos)

    for depth, positions in by_depth.items():
        row_dims = tuple(field_row_cols[:depth])
        row_keys = [tuple(row_entries[p]["prefix"]) for p in positions]

        leaf_block = _eff_grid(row_dims, row_keys, tuple(field_col_cols), col_prefixes)
        total_block = _eff_grid(row_dims, row_keys, (), [()])
        block = np.hstack([leaf_block, total_block])

        for local_i, pos in enumerate(positions):
            data_matrix[pos, :] = block[local_i, :]

    data_rows = data_matrix.tolist()

    if field_col_cols and len(field_col_cols) > 1:
        columns_index = pd.MultiIndex.from_tuples(col_prefixes + [total_col_key])
    else:
        flat_cols = [cp[0] if cp else "efficiency_pct" for cp in col_prefixes]
        columns_index = pd.Index(flat_cols + [total_col_key])

    display_df = pd.DataFrame(
        data_rows,
        index=pd.Index(row_labels, name=" / ".join(row_cols)),
        columns=columns_index,
    )

    # Columns still need the same chronological-month reordering as before
    # (rows don't -- _build_row_entries already sorts each level via
    # _ordered_unique, which is already chronologically correct).
    if field_col_cols:
        non_total_cols = display_df.columns[:-1]
        ordered_non_total = _reorder_axis_chronologically(non_total_cols)
        display_df = display_df[[*list(ordered_non_total), display_df.columns[-1]]]

    display_df = _prettify_date_leaf_axes(display_df, [*row_cols, *col_cols], column_kind_map)
    return display_df, group_row_flags


def _prettify_date_leaf_axes(
    df: pd.DataFrame, original_fields: list[str], column_kind_map: dict[str, str]
) -> pd.DataFrame:
    """
    Cosmetic pass for tables (like the Field Efficiency Table) built from the
    "__<field>__date" leaf column directly rather than through the row/column
    tree builder: renames that internal-looking axis name back to the
    original field name, and reformats its raw 'YYYY-MM-DD' values to
    'DD-Mon-YY' for display -- matching the date formatting used everywhere
    else in the app. Wrapped defensively: if anything about the axis shape
    is unexpected, the table is returned completely unchanged rather than
    risk breaking the view over a formatting nicety.
    """
    try:
        date_leaf_to_original = {
            f"__{f}__date": f for f in original_fields if column_kind_map.get(f) == "date"
        }
        if not date_leaf_to_original:
            return df

        def _relabel(idx: pd.Index) -> pd.Index:
            if isinstance(idx, pd.MultiIndex):
                new_names = [date_leaf_to_original.get(n, n) for n in idx.names]
                new_levels = []
                for level_i, name in enumerate(idx.names):
                    if name in date_leaf_to_original:
                        vals = idx.get_level_values(level_i)
                        parsed = pd.to_datetime(vals, format="%Y-%m-%d", errors="coerce")
                        pretty = parsed.strftime("%d-%b-%y")
                        new_levels.append([p if pd.notna(dt) else v for v, dt, p in zip(vals, parsed, pretty)])
                    else:
                        new_levels.append(list(idx.get_level_values(level_i)))
                return pd.MultiIndex.from_arrays(new_levels, names=new_names)
            if idx.name in date_leaf_to_original:
                parsed = pd.to_datetime(idx, format="%Y-%m-%d", errors="coerce")
                pretty = parsed.strftime("%d-%b-%y")
                new_vals = [p if pd.notna(dt) else v for v, dt, p in zip(idx, parsed, pretty)]
                return pd.Index(new_vals, name=date_leaf_to_original[idx.name])
            return idx

        df = df.copy()
        df.index = _relabel(df.index)
        df.columns = _relabel(df.columns)
        return df
    except Exception:
        return df


def _clean_efficiency_column_label(label: str) -> str:
    """'Sum of march_23' -> 'march_23', 'Mean of april_23' -> 'april_23', etc. Leaves anything else untouched."""
    for agg_word in ("Sum", "Mean", "Count", "Min", "Max"):
        prefix = f"{agg_word} of "
        if label.startswith(prefix):
            return label[len(prefix):]
    return label


def slice_field_efficiency_to_status(
    pivot_eff: pd.DataFrame,
    normalize_field: str,
    status_value: str,
) -> pd.DataFrame:
    """
    Slice an ALREADY-COMPUTED Field Efficiency Table down to just one value
    of `normalize_field` for display (e.g. only 'BOUNCED'), dropping every
    other status value, the 'Total' column, and the now-redundant
    normalize_field header level. Also strips any leading 'Sum of ' /
    'Mean of ' etc. from the remaining Values-column labels.

    CRITICAL: `pivot_eff` must have been built by build_field_efficiency_pivot
    over the FULL dataset (every status value, e.g. BOUNCED *and* CLEARED)
    BEFORE calling this. This function only slices the finished % table for
    display -- it never touches the computation. Filtering the source data
    to just BOUNCED before computing would collapse the denominator to
    BOUNCED/BOUNCED = 100%, which is exactly what this two-step order avoids.
    """
    cols = pivot_eff.columns

    if isinstance(cols, pd.MultiIndex):
        level_names = list(cols.names)
        level = (
            level_names.index(normalize_field)
            if normalize_field in level_names
            else next((i for i in range(cols.nlevels) if status_value in cols.get_level_values(i)), None)
        )
        if level is None:
            raise ValueError(f"Could not find '{status_value}' in any column level of {level_names}.")
        sliced = pivot_eff.xs(status_value, axis=1, level=level, drop_level=True)
    else:
        if status_value not in cols:
            raise ValueError(f"'{status_value}' not found in columns: {list(cols)}")
        sliced = pivot_eff[[status_value]]

    sliced = sliced.copy()
    if isinstance(sliced.columns, pd.MultiIndex):
        sliced.columns = pd.MultiIndex.from_tuples(
            tuple(_clean_efficiency_column_label(v) if isinstance(v, str) else v for v in tup)
            for tup in sliced.columns
        )
    else:
        sliced.columns = [_clean_efficiency_column_label(c) if isinstance(c, str) else c for c in sliced.columns]

    return sliced


def build_pivot_table(
    agg_df: pd.DataFrame,
    row_cols: list[str],
    col_cols: list[str],
    agg_func: str,
) -> pd.DataFrame:
    """
    Build a flat pivot matrix (single overall Grand Total row/column, no
    per-level subtotals) directly from the pre-aggregated stats table --
    used for the Chart and the flat CSV download. `row_cols`/`col_cols` here
    are the *effective* (already date-expanded) field names, i.e. exactly
    the columns present in `agg_df`.

    Grand-total margins are reconstructed CORRECTLY for every agg_func
    (sum/count/mean/min/max) by re-combining the underlying sum/count/min/
    max stats via `_combine_group_stats` -- never by summing already
    per-cell-aggregated values, which is exactly right for sum/count but
    silently wrong for mean/min/max (pandas' own `pivot_table(margins=True)`
    makes this mistake for anything other than sum/count).
    """
    leaf = _combine_group_stats(agg_df, [*row_cols, *col_cols], agg_func)
    leaf_flat = leaf.reset_index(name="value")
    pivot = leaf_flat.pivot_table(
        index=row_cols, columns=col_cols if col_cols else None, values="value", aggfunc="sum", fill_value=0,
    )
    # aggfunc="sum" above is inert/safe: `leaf` already has exactly one row
    # per unique Rows x Columns combination (the full grouping granularity),
    # so this pivot_table call only *reshapes* -- it never combines more
    # than one raw value into a cell.
    original_columns = pivot.columns

    row_total = _combine_group_stats(agg_df, row_cols, agg_func)
    total_col_key = ("Total",) + ("",) * (pivot.columns.nlevels - 1) if isinstance(pivot.columns, pd.MultiIndex) else "Total"
    pivot[total_col_key] = row_total.reindex(pivot.index).fillna(0.0).to_numpy()

    if col_cols:
        col_total = _combine_group_stats(agg_df, col_cols, agg_func).reindex(original_columns).fillna(0.0)
    else:
        col_total = pd.Series([_combine_group_stats(agg_df, [], agg_func).iloc[0]] * len(original_columns), index=original_columns)
    corner = _combine_group_stats(agg_df, [], agg_func).iloc[0]

    if isinstance(pivot.index, pd.MultiIndex):
        total_row_index = pd.MultiIndex.from_tuples(
            [("Total",) + ("",) * (pivot.index.nlevels - 1)], names=pivot.index.names
        )
    else:
        # Preserve the original index name (e.g. "del_month_name") explicitly --
        # pd.concat sets the resulting index name to None whenever the pieces
        # being concatenated disagree on it, and an unnamed pd.Index(["Total"])
        # always disagrees with the real (named) pivot index. Losing the name
        # here breaks every downstream reset_index()/melt() call that expects
        # a column literally called by that field's name (e.g. the Chart).
        total_row_index = pd.Index(["Total"], name=pivot.index.name)
    total_row_values = list(col_total.to_numpy()) + [corner]
    total_row_df = pd.DataFrame([total_row_values], index=total_row_index, columns=pivot.columns)

    pivot = pd.concat([pivot, total_row_df], axis=0)
    pivot = pivot.reindex(index=_reorder_axis_chronologically(pivot.index))
    pivot = pivot.reindex(columns=_reorder_axis_chronologically(pivot.columns))
    return pivot


# --------------------------------------------------------------------------- #
# Authentication & Authorization
# --------------------------------------------------------------------------- #
# Two tables, on purpose:
#   authorized_employees -- WHO is allowed to have an account at all (admin-managed)
#   app_users            -- the password hash for employees who created one
# Kept separate so access can be revoked (is_authorized = FALSE) without
# touching the password, and so "authorized" and "has an account" are two
# independently-true-or-false things.
_AUTH_SESSION_KEYS = ("auth_employee_code", "auth_employee_name", "auth_access_level")

_CREATE_EMPLOYEES_SQL = text(
    """
    CREATE TABLE IF NOT EXISTS authorized_employees (
        employee_code   VARCHAR(20) PRIMARY KEY,
        employee_name   VARCHAR(100) NOT NULL,
        department      VARCHAR(100),
        access_level    VARCHAR(20) NOT NULL DEFAULT 'user'
                        CHECK (access_level IN ('admin', 'user', 'analyst')),
        is_authorized   BOOLEAN NOT NULL DEFAULT TRUE,
        created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
    )
    """
)
_CREATE_USERS_SQL = text(
    """
    CREATE TABLE IF NOT EXISTS app_users (
        employee_code   VARCHAR(20) PRIMARY KEY
                        REFERENCES authorized_employees(employee_code) ON DELETE CASCADE,
        password_hash   TEXT NOT NULL,
        created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
        last_login      TIMESTAMPTZ
    )
    """
)
_CREATE_AUDIT_SQL = text(
    """
    CREATE TABLE IF NOT EXISTS login_audit (
        id              BIGSERIAL PRIMARY KEY,
        employee_code   VARCHAR(20) NOT NULL,
        event           VARCHAR(20) NOT NULL
                        CHECK (event IN ('login_success', 'login_failure', 'logout', 'account_created')),
        event_time      TIMESTAMPTZ NOT NULL DEFAULT now()
    )
    """
)




@ttl_cache(ttl=3600)
def ensure_auth_tables(_engine: Engine) -> None:
    """Create the auth tables if they don't exist yet. Runs once per app process."""
    with _engine.begin() as conn:
        conn.execute(_CREATE_EMPLOYEES_SQL)
        conn.execute(_CREATE_USERS_SQL)
        conn.execute(_CREATE_AUDIT_SQL)


def _is_employee_authorized(engine: Engine, employee_code: str) -> tuple[bool, str | None, str | None]:
    """Return (is_authorized, employee_name, access_level); (False, None, None) if unknown/disabled."""
    query = text(
        "SELECT employee_name, access_level, is_authorized "
        "FROM authorized_employees WHERE employee_code = :code"
    )
    with engine.connect() as conn:
        row = conn.execute(query, {"code": employee_code}).mappings().fetchone()
    if row is None or not row["is_authorized"]:
        return False, None, None
    return True, row["employee_name"], row["access_level"]


def _get_password_hash(engine: Engine, employee_code: str) -> str | None:
    query = text("SELECT password_hash FROM app_users WHERE employee_code = :code")
    with engine.connect() as conn:
        row = conn.execute(query, {"code": employee_code}).fetchone()
    return row[0] if row else None


def _account_exists(engine: Engine, employee_code: str) -> bool:
    return _get_password_hash(engine, employee_code) is not None


def _create_account(engine: Engine, employee_code: str, password: str) -> None:
    password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    query = text("INSERT INTO app_users (employee_code, password_hash) VALUES (:code, :hash)")
    with engine.begin() as conn:
        conn.execute(query, {"code": employee_code, "hash": password_hash})
    _log_audit_event(engine, employee_code, "account_created")


def _verify_password(password: str, password_hash: str) -> bool:
    try:
        return bcrypt.checkpw(password.encode("utf-8"), password_hash.encode("utf-8"))
    except (ValueError, TypeError):
        return False


def _touch_last_login(engine: Engine, employee_code: str) -> None:
    query = text("UPDATE app_users SET last_login = now() WHERE employee_code = :code")
    with engine.begin() as conn:
        conn.execute(query, {"code": employee_code})


def _log_audit_event(engine: Engine, employee_code: str, event: str) -> None:
    query = text("INSERT INTO login_audit (employee_code, event) VALUES (:code, :event)")
    try:
        with engine.begin() as conn:
            conn.execute(query, {"code": employee_code, "event": event})
    except SQLAlchemyError:
        pass  # audit logging must never be the reason a login fails


def _too_many_recent_failures(
    engine: Engine, employee_code: str, max_attempts: int = 5, window_minutes: int = 15
) -> bool:
    """Basic brute-force throttle: block further attempts if too many recent failures.

    Checked against the database (not st.session_state) so it can't be bypassed
    by simply reloading the page or opening a new browser tab.
    """
    query = text(
        "SELECT COUNT(*) FROM login_audit "
        "WHERE employee_code = :code AND event = 'login_failure' "
        "AND event_time > now() - (:window || ' minutes')::interval"
    )
    with engine.connect() as conn:
        count = conn.execute(query, {"code": employee_code, "window": window_minutes}).scalar()
    return (count or 0) >= max_attempts