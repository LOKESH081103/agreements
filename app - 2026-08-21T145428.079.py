"""
Dynamic PostgreSQL Explorer & Report Generator
================================================
A production-ready Streamlit application that lets a user pick any table
in a PostgreSQL database, filter it by a reference column (date/timestamp
or categorical), pick one or more output columns, and view/download the
resulting report with auto-generated KPI cards for numeric columns.

Run with:
    streamlit run app.py

Configuration is read from environment variables (see .env.example / README).
"""

from __future__ import annotations

import io
import math
import os
from datetime import date
from typing import Any

import pandas as pd
import plotly.express as px
import streamlit as st
from sqlalchemy import create_engine, inspect, text
from sqlalchemy.engine import Engine
from sqlalchemy.exc import SQLAlchemyError

# --------------------------------------------------------------------------- #
# Page configuration
# --------------------------------------------------------------------------- #
st.set_page_config(
    page_title="DB Explorer & Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --------------------------------------------------------------------------- #
# Configuration / Environment
# --------------------------------------------------------------------------- #
# Priority: st.secrets (if running on Streamlit Cloud / has a secrets.toml)
# falls back to plain environment variables. This keeps local dev (.env via
# os.environ, e.g. loaded through `python-dotenv` or exported in shell) and
# cloud deployment (st.secrets) both working without code changes.


def _get_config_value(key: str, default: str | None = None) -> str | None:
    """Fetch a config value from st.secrets first, then environment vars."""
    try:
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        # st.secrets raises if no secrets.toml exists at all -- that's fine,
        # we just fall through to environment variables.
        pass
    return os.environ.get(key, default)


DB_HOST = _get_config_value("DB_HOST", "localhost")
DB_PORT = _get_config_value("DB_PORT", "5432")
DB_NAME = _get_config_value("DB_NAME", "postgres")
DB_USER = _get_config_value("DB_USER", "postgres")
DB_PASSWORD = _get_config_value("DB_PASSWORD", "")
DB_SSLMODE = _get_config_value("DB_SSLMODE", "prefer")  # e.g. "require" for cloud DBs

# Allow a full DATABASE_URL to override the individual pieces above.
DATABASE_URL = _get_config_value("DATABASE_URL")


# --------------------------------------------------------------------------- #
# Database Engine (cached across reruns / sessions)
# --------------------------------------------------------------------------- #
@st.cache_resource(show_spinner="Connecting to database...")
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


def get_engine_safe() -> Engine | None:
    """Wrapper that turns connection failures into a clean st.error message."""
    try:
        return get_engine()
    except SQLAlchemyError as exc:
        st.error(
            "❌ Could not connect to the database. Please check your connection "
            f"settings.\n\nDetails: `{exc.__class__.__name__}: {exc}`"
        )
    except Exception as exc:  # noqa: BLE001 - surface any other startup issue
        st.error(f"❌ Unexpected error while connecting to the database: {exc}")
    return None


# --------------------------------------------------------------------------- #
# Schema introspection helpers (cached so we don't re-hit the DB every rerun)
# --------------------------------------------------------------------------- #
@st.cache_data(ttl=300, show_spinner=False)
def list_tables(_engine: Engine) -> list[str]:
    """List all tables in the 'public' schema."""
    inspector = inspect(_engine)
    return sorted(inspector.get_table_names(schema="public"))


@st.cache_data(ttl=300, show_spinner=False)
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


@st.cache_data(ttl=300, show_spinner=False)
def get_distinct_values(_engine: Engine, table_name: str, column_name: str, limit: int = 1000) -> list[Any]:
    """Fetch distinct non-null values for a categorical column (capped)."""
    query = text(
        f'SELECT DISTINCT "{column_name}" FROM "public"."{table_name}" '
        f'WHERE "{column_name}" IS NOT NULL ORDER BY "{column_name}" LIMIT :limit'
    )
    with _engine.connect() as conn:
        result = conn.execute(query, {"limit": limit})
        return [row[0] for row in result]


@st.cache_data(ttl=300, show_spinner=False)
def get_available_years(_engine: Engine, table_name: str, column_name: str) -> list[int]:
    """Fetch distinct years present in a date/timestamp column."""
    query = text(
        f'SELECT DISTINCT EXTRACT(YEAR FROM "{column_name}")::int AS yr '
        f'FROM "public"."{table_name}" WHERE "{column_name}" IS NOT NULL ORDER BY yr'
    )
    with _engine.connect() as conn:
        result = conn.execute(query)
        return [row[0] for row in result]


@st.cache_data(ttl=300, show_spinner=False)
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

# --------------------------------------------------------------------------- #
# Crore formatting helper
# --------------------------------------------------------------------------- #
CRORE = 10_000_000  # 1,00,00,000


def format_amount(value: float, in_crores: bool, decimals: int = 2) -> str:
    """Format a numeric total either as-is or converted to Crores (÷1,00,00,000)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "0.00 Cr" if in_crores else "0.00"
    if in_crores:
        return f"{value / CRORE:,.{decimals}f} Cr"
    return f"{value:,.{decimals}f}"


def series_to_crores(series: pd.Series) -> pd.Series:
    """Convert a numeric pandas Series to Crores."""
    return series / CRORE


@st.cache_data(ttl=300, show_spinner=False)
def get_monthly_summary(
    _engine: Engine,
    table_name: str,
    date_col: str,
    category_col: str | None = None,
    sum_cols: tuple[str, ...] = (),
) -> pd.DataFrame:
    """
    Aggregate row counts (and, optionally, SUM of one or more numeric columns)
    per calendar month based on `date_col`, optionally broken down by a second
    column (e.g. payment_mode). A record dated 01-03-2024 is grouped under
    March 2024 regardless of the day of month — this mirrors
    EXTRACT(MONTH FROM ...), same as the existing Month/Year filter above.

    Each requested sum column `col` shows up in the result as `sum_<col>`.
    """
    select_parts = [
        f'EXTRACT(YEAR FROM "{date_col}")::int AS yr',
        f'EXTRACT(MONTH FROM "{date_col}")::int AS mo',
    ]
    group_parts = ["yr", "mo"]

    if category_col:
        select_parts.append(f'"{category_col}"')
        group_parts.append(f'"{category_col}"')

    select_parts.append("COUNT(*) AS record_count")
    for col in sum_cols:
        select_parts.append(f'SUM("{col}") AS "sum_{col}"')

    query = text(
        f'SELECT {", ".join(select_parts)} '
        f'FROM "public"."{table_name}" '
        f'WHERE "{date_col}" IS NOT NULL '
        f'GROUP BY {", ".join(group_parts)} '
        f'ORDER BY yr, mo'
    )
    with _engine.connect() as conn:
        return pd.read_sql(query, conn)


def render_labeled_bar_chart(
    df: pd.DataFrame,
    x_col: str,
    y_col: str,
    color_col: str | None,
    title: str,
    filename_prefix: str,
    key_suffix: str,
    res_scale: int = 3,
    y_label: str = "Count",
    value_format: str = ",",
) -> None:
    """
    Render a Plotly bar chart with the value of each bar printed on top of
    it, styled for high-resolution export. The base canvas is 1600x800px;
    at res_scale=3 the PNG export is 4800x2400 and at res_scale=4 it's
    6400x3200 — both comfortably exceed 4K (3840x2160).

    `y_label` controls the y-axis/legend title text (e.g. "Sum of emi (Cr)")
    and `value_format` controls the on-bar number format (Plotly d3-format
    spec) -- use ",.2f" for decimal amounts like Crore totals, or the
    default "," for whole-number counts.
    """
    fig = px.bar(
        df,
        x=x_col,
        y=y_col,
        color=color_col,
        barmode="group" if color_col else "relative",
        template="plotly_white",
        text=y_col,
        labels={x_col: x_col.replace("_", " ").title(), y_col: y_label},
        title=title,
    )
    fig.update_traces(
        texttemplate="%{text:" + value_format + "}",
        textposition="outside",
        textfont_size=13,
        cliponaxis=False,
    )
    fig.update_layout(
        width=1600,
        height=800,
        font=dict(size=14),
        title_font_size=18,
        uniformtext_minsize=10,
        uniformtext_mode="hide",
        margin=dict(t=80, l=60, r=40, b=60),
        legend_title_text=color_col or "",
        yaxis_title=y_label,
    )

    config = {
        "displaylogo": False,
        "toImageButtonOptions": {
            "format": "png",
            "filename": filename_prefix,
            "scale": res_scale,
        },
    }
    st.plotly_chart(fig, use_container_width=True, config=config, key=f"chart_{key_suffix}")
    st.caption(
        "💡 Click the camera icon in the chart's toolbar (top-right of the chart on hover) "
        f"to export a high-resolution PNG at ≈{1600 * res_scale}×{800 * res_scale}px."
    )


def render_split_bar_charts(
    df: pd.DataFrame,
    x_col: str,
    y_col: str,
    color_col: str | None,
    title: str,
    filename_prefix: str,
    key_suffix: str,
    res_scale: int = 3,
    num_splits: int = 1,
    y_label: str = "Count",
    value_format: str = ",",
) -> None:
    """
    Render one bar chart, or break the x-axis categories (e.g. months) into
    `num_splits` roughly-equal groups and render one chart per group. This
    keeps a chart with many bars from becoming too congested/cluttered to
    read — each smaller chart gets its own title, key, and PNG filename.
    """
    x_order = df[x_col].drop_duplicates().tolist()
    n = len(x_order)

    if num_splits <= 1 or n <= num_splits:
        render_labeled_bar_chart(
            df, x_col, y_col, color_col, title, filename_prefix, key_suffix, res_scale,
            y_label=y_label, value_format=value_format,
        )
        return

    chunk_size = math.ceil(n / num_splits)
    for i in range(0, n, chunk_size):
        chunk_labels = x_order[i : i + chunk_size]
        chunk_df = df[df[x_col].isin(chunk_labels)]
        part_num = i // chunk_size + 1
        render_labeled_bar_chart(
            chunk_df,
            x_col=x_col,
            y_col=y_col,
            color_col=color_col,
            title=f"{title} — Part {part_num} ({chunk_labels[0]} to {chunk_labels[-1]})",
            filename_prefix=f"{filename_prefix}_part{part_num}",
            key_suffix=f"{key_suffix}_part{part_num}",
            res_scale=res_scale,
            y_label=y_label,
            value_format=value_format,
        )


# --------------------------------------------------------------------------- #
# Query builder
# --------------------------------------------------------------------------- #
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


@st.cache_data(ttl=60, show_spinner=False)
def run_query(_engine: Engine, query: str, params: dict[str, Any]) -> pd.DataFrame:
    """Execute the (already-parameterized) query and return a DataFrame."""
    with _engine.connect() as conn:
        return pd.read_sql(text(query), conn, params=params)


# --------------------------------------------------------------------------- #
# Excel-style 4-Quadrant Pivot Table engine
# --------------------------------------------------------------------------- #
PIVOT_AGG_FUNCS = ["sum", "mean", "count", "min", "max"]


@st.cache_data(ttl=60, show_spinner=False)
def fetch_pivot_source_data(
    _engine: Engine,
    table_name: str,
    columns_needed: tuple[str, ...],
    filter_conditions: tuple[tuple[str, tuple[Any, ...]], ...],
) -> pd.DataFrame:
    """
    Pull only the columns required to build the pivot (Filters + Rows +
    Columns + Value) from the DB, applying any chosen Filter values as a
    server-side parameterized WHERE ... IN (...) clause so we don't drag the
    whole table across the wire just to build a small pivot matrix.
    """
    quoted_cols = ", ".join(f'"{c}"' for c in columns_needed)
    query = f'SELECT {quoted_cols} FROM "public"."{table_name}"'
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

    with _engine.connect() as conn:
        return pd.read_sql(text(query), conn, params=params)

def _expand_date_fields(
    df: pd.DataFrame,
    fields: list[str],
    column_kind_map: dict[str, str],
    label_map: dict[tuple[str, Any], str],
) -> tuple[pd.DataFrame, list[str]]:
    """
    Mirrors Excel's automatic date grouping: if a chosen Rows/Columns field is
    a date/timestamp column, split it into two synthetic sortable levels —
    'YYYY-MM' (displayed as 'Mon-YY') and 'YYYY-MM-DD' (displayed as
    'DD-Mon-YY') — so a single date field automatically produces the same
    Month > Date nesting Excel shows when you drop a date into the
    PivotTable. The synthetic fields sort correctly because 'YYYY-MM' /
    'YYYY-MM-DD' strings sort lexically the same as chronologically;
    `label_map` carries the pretty text shown to the user for each raw key.
    """
    work_df = df.copy()
    expanded: list[str] = []
    for f in fields:
        if column_kind_map.get(f) == "date":
            ts = pd.to_datetime(work_df[f], errors="coerce")
            month_field, date_field = f"__{f}__month", f"__{f}__date"
            work_df[month_field] = ts.dt.strftime("%Y-%m")
            work_df[date_field] = ts.dt.strftime("%Y-%m-%d")

            month_keys = work_df[month_field].dropna().unique()
            if len(month_keys):
                for k, disp in zip(month_keys, pd.to_datetime(month_keys).strftime("%b-%y")):
                    label_map[(month_field, k)] = disp
            date_keys = work_df[date_field].dropna().unique()
            if len(date_keys):
                for k, disp in zip(date_keys, pd.to_datetime(date_keys).strftime("%d-%b-%y")):
                    label_map[(date_field, k)] = disp

            expanded.extend([month_field, date_field])
        else:
            expanded.append(f)
    return work_df, expanded


def _pivot_agg_value(
    df: pd.DataFrame,
    value_col: str,
    agg_func: str,
    row_fields: list[str],
    row_prefix: tuple[Any, ...],
    col_fields: list[str],
    col_prefix: tuple[Any, ...],
) -> float:
    """Aggregate value_col over exactly the rows matching row_prefix + col_prefix
    (a *partial* prefix aggregates over every deeper level — this is what makes
    subtotal cells correct for sum/mean/count/min/max alike)."""
    mask = pd.Series(True, index=df.index)
    for f, v in zip(row_fields, row_prefix):
        mask &= df[f] == v
    for f, v in zip(col_fields, col_prefix):
        mask &= df[f] == v
    subset = df.loc[mask, value_col].dropna()
    if subset.empty:
        return 0.0
    if agg_func == "count":
        return float(subset.count())
    return float(getattr(subset, agg_func)())


def _ordered_unique(series: pd.Series) -> list[Any]:
    """Distinct non-null values, sorted when possible (falls back to first-seen order)."""
    vals = [v for v in pd.unique(series) if pd.notna(v)]
    try:
        return sorted(vals)
    except TypeError:
        return vals


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
    """
    level = len(prefix)
    if level >= len(row_fields):
        return []
    field = row_fields[level]
    mask = pd.Series(True, index=df.index)
    for f, v in zip(row_fields[:level], prefix):
        mask &= df[f] == v
    values = _ordered_unique(df.loc[mask, field])
    is_last_level = level == len(row_fields) - 1

    entries: list[dict[str, Any]] = []
    for v in values:
        new_prefix = prefix + (v,)
        disp = label_map.get((field, v), str(v))
        if is_last_level:
            entries.append({"label": disp, "prefix": new_prefix, "indent": indent, "is_group": False})
        else:
            entries.append({"label": disp, "prefix": new_prefix, "indent": indent, "is_group": True})
            entries.extend(_build_row_entries(df, row_fields, label_map, new_prefix, indent + 1))
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
    Jul Total, then STAGE_1 Total) — the mirror image of the row ordering.
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


@st.cache_data(ttl=300, show_spinner=False)
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

        row_vals = [
            _pivot_agg_value(
                work_df, value_col, agg_func, row_fields, re_["prefix"],
                ce.get("mask_fields", col_fields), ce["prefix"],
            ) / divisor
            for ce in col_entries
        ]
        data_rows.append(row_vals)

    display_df = pd.DataFrame(
        data_rows,
        index=pd.Index(row_labels, name=" / ".join(row_cols)),
        columns=columns_index,
    )
    return display_df, group_row_flags, total_col_flags


def render_excel_style_pivot_table(
    source_df: pd.DataFrame,
    row_cols: list[str],
    col_cols: list[str],
    value_col: str,
    agg_func: str,
    column_kind_map: dict[str, str],
    in_crores: bool,
    right_total_mode: str = "grand_total",
) -> pd.DataFrame:
    """Build the Excel-style nested pivot, render it (styled where possible), return it for CSV export."""
    display_df, group_row_flags, total_col_flags = build_excel_style_pivot(
        source_df, row_cols, col_cols, value_col, agg_func, column_kind_map, in_crores, right_total_mode
    )

    crore_suffix = " (in Cr)" if (in_crores and agg_func != "count") else ""
    st.markdown(f"**{agg_func.title()} of {value_col}{crore_suffix}**")

    try:
        num_fmt = "{:,.0f}" if agg_func == "count" else "{:,.2f}"

        def _apply_styles(df: pd.DataFrame) -> pd.DataFrame:
            styles = pd.DataFrame("", index=df.index, columns=df.columns)
            for i, (label, is_group) in enumerate(zip(row_labels_for_style, group_row_flags)):
                if label == "Grand Total":
                    styles.iloc[i, :] = "font-weight:bold"
                elif is_group:
                    styles.iloc[i, :] = "background-color:#dbe5f1; font-weight:bold"
            for j, is_total in enumerate(total_col_flags):
                if is_total:
                    for i in range(len(df)):
                        existing = styles.iloc[i, j]
                        styles.iloc[i, j] = f"{existing}; background-color:#f2f2f2; font-weight:bold" if existing else "background-color:#f2f2f2; font-weight:bold"
            return styles

        row_labels_for_style = list(display_df.index)
        styler = display_df.style.apply(_apply_styles, axis=None).format(num_fmt)
        st.dataframe(styler, use_container_width=True)
    except Exception:
        # If styling fails for any reason (older pandas/streamlit version, etc.),
        # fall back to a plain, still-correct, unstyled table rather than crashing.
        st.dataframe(display_df, use_container_width=True)

    return display_df


# --------------------------------------------------------------------------- #
# General Field Efficiency Table: each cell as a % of the total across every
# OTHER value of a user-chosen field, holding every other selected Rows/
# Columns field fixed. Unlike a hardcoded "% of Stage" table, this works for
# *any* field the user points it at (a Rows field or a Columns field) — pick
# "STAGE_ECL" to get % of stage-group total per (Zone, Month); pick "zone"
# instead to get % of zone-group total per (Stage, Month); same formula.
# --------------------------------------------------------------------------- #
@st.cache_data(ttl=300, show_spinner=False)
def build_field_efficiency_pivot(
    source_df: pd.DataFrame,
    row_cols: list[str],
    col_cols: list[str],
    normalize_field: str,
    value_col: str,
    agg_func: str,
) -> pd.DataFrame:
    """
    efficiency(cell) = value(cell) ÷ sum of value across every other value of
    `normalize_field`, holding every OTHER selected Rows/Columns field fixed.

    E.g. with Rows=[zone], Columns=[STAGE_ECL, value_date], normalize_field=
    "STAGE_ECL": Stage 1 / East / July = value(Stage1,East,July) ÷
    [value(Stage1,East,July) + value(Stage2,East,July) + value(Stage3,East,July)].
    Point it at "zone" instead and it normalizes across zones per (Stage, Month).

    Returns a DataFrame shaped exactly like the raw pivot (same row/column
    axes), just with each cell replaced by its % share (already ×100).
    """
    all_dims = list(dict.fromkeys([*row_cols, *col_cols]))  # de-dup, preserve order
    if normalize_field not in all_dims:
        raise ValueError(f"'{normalize_field}' must be one of the selected Rows/Columns fields.")

    grouped = source_df.groupby(all_dims, dropna=False)[value_col]
    leaf_agg = (grouped.count() if agg_func == "count" else grouped.agg(agg_func)).astype(float)

    other_fields = [c for c in all_dims if c != normalize_field]
    if other_fields:
        denom = leaf_agg.groupby(level=other_fields).transform("sum")
    else:
        denom = pd.Series(leaf_agg.sum(), index=leaf_agg.index)

    efficiency = (leaf_agg.where(denom != 0) / denom.where(denom != 0) * 100.0).fillna(0.0)

    eff_flat = efficiency.reset_index(name="efficiency_pct")
    pivot_eff = eff_flat.pivot_table(
        index=row_cols,
        columns=col_cols if col_cols else None,
        values="efficiency_pct",
        aggfunc="sum",
        fill_value=0.0,
    )
    return pivot_eff


def render_field_efficiency_table(
    source_df: pd.DataFrame,
    row_cols: list[str],
    col_cols: list[str],
    normalize_field: str,
    value_col: str,
    agg_func: str,
) -> pd.DataFrame:
    """Build the general Field Efficiency table, render it, return it for CSV export."""
    pivot_eff = build_field_efficiency_pivot(source_df, row_cols, col_cols, normalize_field, value_col, agg_func)

    st.markdown(f"**{agg_func.title()} of {value_col} — % of `{normalize_field}` group total**")
    st.caption(
        f"Each cell = its value ÷ the sum across every value of **`{normalize_field}`** for that same "
        "combination of the other selected Rows/Columns fields — computed automatically for every "
        "row and column in the table, however many there are. Pick a different field above to "
        "normalize a different way (e.g. % across zones instead of % across stages)."
    )

    try:
        styler = pivot_eff.style.format("{:,.2f}%")
        st.dataframe(styler, use_container_width=True)
    except Exception:
        st.dataframe(pivot_eff, use_container_width=True)

    return pivot_eff


def build_pivot_table(
    df: pd.DataFrame,
    row_cols: list[str],
    col_cols: list[str],
    value_col: str,
    agg_func: str,
) -> pd.DataFrame:
    """
    Build an Excel-style pivot matrix: Rows down the side, Columns across the
    top, the chosen Value aggregated at each intersection, with a grand-total
    row and column (margins=True) — exactly like Excel's PivotTable.
    """
    return pd.pivot_table(
        df,
        index=row_cols,
        columns=col_cols if col_cols else None,
        values=value_col,
        aggfunc=agg_func,
        fill_value=0,
        margins=True,
        margins_name="Total",
    )


def _drop_pivot_totals(pivot_df: pd.DataFrame) -> pd.DataFrame:
    """Strip the 'Total' margin row/column so charts show only real categories."""
    trimmed = pivot_df.copy()
    if "Total" in trimmed.index:
        trimmed = trimmed.drop(index="Total")
    if isinstance(trimmed.columns, pd.MultiIndex):
        mask = trimmed.columns.get_level_values(0) != "Total"
        trimmed = trimmed.loc[:, mask]
    elif "Total" in trimmed.columns:
        trimmed = trimmed.drop(columns="Total")
    return trimmed


def render_pivot_chart(
    pivot_df: pd.DataFrame,
    row_cols: list[str],
    col_cols: list[str],
    value_col: str,
    agg_func: str,
    key_suffix: str,
) -> None:
    """Render an interactive Plotly grouped bar chart matching the pivoted dimensions."""
    chart_df = _drop_pivot_totals(pivot_df)
    if chart_df.empty:
        st.info("Nothing to chart once the grand-total row/column is excluded.")
        return

    # Flatten a MultiIndex column axis (happens when 2+ "Columns" fields are chosen)
    # into single readable string labels before melting for Plotly.
    if isinstance(chart_df.columns, pd.MultiIndex):
        chart_df.columns = [
            " | ".join(str(part) for part in col if str(part) != "") if isinstance(col, tuple) else str(col)
            for col in chart_df.columns
        ]

    chart_df = chart_df.reset_index()
    value_vars = [c for c in chart_df.columns if c not in row_cols]
    if not value_vars:
        st.info("No values available to chart.")
        return

    long_df = chart_df.melt(id_vars=row_cols, value_vars=value_vars, var_name="__pivot_col__", value_name="__value__")

    if len(row_cols) == 1:
        x_col = row_cols[0]
    else:
        x_col = "__row_label__"
        long_df[x_col] = long_df[row_cols].astype(str).agg(" | ".join, axis=1)

    color_col = "__pivot_col__" if col_cols else None
    title_bits = f"{agg_func.title()} of {value_col} by {', '.join(row_cols)}"
    if col_cols:
        title_bits += f", split by {', '.join(col_cols)}"

    fig = px.bar(
        long_df,
        x=x_col,
        y="__value__",
        color=color_col,
        barmode="group",
        template="plotly_white",
        title=title_bits,
        labels={
            "__value__": f"{agg_func}({value_col})",
            x_col: " | ".join(row_cols),
            "__pivot_col__": " | ".join(col_cols) if col_cols else "",
        },
    )
    fig.update_layout(
        height=550,
        font=dict(size=13),
        title_font_size=16,
        margin=dict(t=70, l=50, r=30, b=60),
        legend_title_text=" | ".join(col_cols) if col_cols else "",
    )
    config = {"displaylogo": False, "toImageButtonOptions": {"format": "png", "filename": f"pivot_chart_{key_suffix}", "scale": 3}}
    st.plotly_chart(fig, use_container_width=True, config=config, key=f"pivot_chart_{key_suffix}")


# --------------------------------------------------------------------------- #
# UI - Sidebar: connection status
# --------------------------------------------------------------------------- #
st.title("📊 Dynamic Database Explorer & Report Generator")
st.caption("Select a table, filter by any column, and export a custom report.")

engine = get_engine_safe()
if engine is None:
    st.info(
        "Set your database credentials as environment variables (or in "
        "`.streamlit/secrets.toml`) and reload the app. See the README / "
        "setup instructions for details."
    )
    st.stop()

with st.sidebar:
    st.header("⚙️ Configuration")
    st.success(f"Connected to **{DB_NAME}**@`{DB_HOST}`", icon="✅")

    try:
        tables = list_tables(engine)
    except SQLAlchemyError as exc:
        st.error(f"Failed to list tables: {exc}")
        st.stop()

    if not tables:
        st.warning("No tables found in the `public` schema.")
        st.stop()

    st.divider()
    selected_table = st.selectbox("📁 Select a table", options=tables, index=0)

# --------------------------------------------------------------------------- #
# Load column metadata for the selected table
# --------------------------------------------------------------------------- #
try:
    columns_meta = get_columns(engine, selected_table)
except SQLAlchemyError as exc:
    st.error(f"Failed to fetch columns for `{selected_table}`: {exc}")
    st.stop()

column_names = [c["name"] for c in columns_meta]
column_kind_map = {c["name"]: classify_column(c["type"]) for c in columns_meta}

with st.sidebar:
    st.divider()
    st.subheader("1️⃣ Reference (Filter) Column")
    ref_column = st.selectbox("Reference column", options=column_names, key="ref_col")

    detected_kind = column_kind_map.get(ref_column, "text")
    kind_override = st.radio(
        "Treat this column as",
        options=["Date/Timestamp", "Categorical/Text"],
        index=0 if detected_kind == "date" else 1,
        help="Auto-detected from the database column type; override if needed.",
    )
    ref_kind = "date" if kind_override == "Date/Timestamp" else "text"

    filter_payload: dict[str, Any] = {}

    if ref_kind == "date":
        granularity = st.radio(
            "Granularity",
            options=["Specific Date(s)", "Month/Year", "Year(s)"],
            key="granularity",
        )
        filter_payload["granularity"] = granularity

        if granularity == "Specific Date(s)":
            picked_dates = st.date_input(
                "Pick one or more dates",
                value=[],
                help="Click multiple dates; use the calendar's range/multi picker.",
            )
            # st.date_input returns a single date, a tuple, or a list depending
            # on interaction state -- normalize to a list of dates.
            if isinstance(picked_dates, (list, tuple)):
                filter_payload["dates"] = list(picked_dates)
            elif picked_dates:
                filter_payload["dates"] = [picked_dates]
            else:
                filter_payload["dates"] = []

        elif granularity == "Month/Year":
            try:
                available_ym = get_available_year_months(engine, selected_table, ref_column)
            except SQLAlchemyError as exc:
                st.error(f"Failed to fetch available months: {exc}")
                available_ym = []
            ym_labels = {f"{MONTH_NAMES[mo - 1]} {yr}": (yr, mo) for yr, mo in available_ym}
            selected_labels = st.multiselect("Select Month/Year", options=list(ym_labels.keys()))
            filter_payload["year_months"] = [ym_labels[lbl] for lbl in selected_labels]

        elif granularity == "Year(s)":
            try:
                available_years = get_available_years(engine, selected_table, ref_column)
            except SQLAlchemyError as exc:
                st.error(f"Failed to fetch available years: {exc}")
                available_years = []
            filter_payload["years"] = st.multiselect("Select Year(s)", options=available_years)

    else:
        try:
            distinct_values = get_distinct_values(engine, selected_table, ref_column)
        except SQLAlchemyError as exc:
            st.error(f"Failed to fetch distinct values: {exc}")
            distinct_values = []
        filter_payload["values"] = st.multiselect(
            f"Select value(s) of `{ref_column}`", options=distinct_values
        )

    st.divider()
    st.subheader("2️⃣ Target / Output Columns")
    output_candidates = [c for c in column_names if c != ref_column]
    output_columns = st.multiselect(
        "Select one or more output columns",
        options=output_candidates,
        default=output_candidates[:1] if output_candidates else [],
    )

    st.divider()
    st.subheader("3️⃣ Monthly Cumulative Summary (optional)")
    show_monthly_summary = st.checkbox(
        "Show monthly cumulative summary",
        value=False,
        help=(
            "Counts records per calendar month — a record dated 01-03-2024 is "
            "counted as a March record. Optionally break it down by another "
            "column, e.g. payment_mode (Cash / Cheque / DD)."
        ),
    )

    summary_date_col: str | None = None
    breakdown_col: str | None = None
    focus_year_months: list[tuple[int, int]] = []
    focus_years_summary: list[int] = []
    focus_categories: list[Any] = []
    chart_res_scale: int = 3
    sum_columns: list[str] = []
    show_in_crores: bool = False
    num_chart_splits: int = 1

    if show_monthly_summary:
        date_columns = [c for c in column_names if column_kind_map.get(c) == "date"]
        if not date_columns:
            st.warning("No date/timestamp columns found in this table — can't build a monthly summary.")
            show_monthly_summary = False
        else:
            default_date_col = ref_column if ref_kind == "date" and ref_column in date_columns else date_columns[0]
            summary_date_col = st.selectbox(
                "Date column to summarize by month",
                options=date_columns,
                index=date_columns.index(default_date_col),
                key=f"summary_date_col_{selected_table}",
            )

            breakdown_candidates = ["(None)"] + [c for c in column_names if c != summary_date_col]
            breakdown_choice = st.selectbox(
                "Breakdown by (optional) — e.g. payment_mode, or a numeric column like emi to sum it",
                options=breakdown_candidates,
                key=f"breakdown_choice_{selected_table}",
                help=(
                    "Pick a text/category column (e.g. payment_mode) to split the chart "
                    "into groups. Pick a numeric column (e.g. emi) to total it instead — "
                    "numeric columns are never used to split/color the chart, since they "
                    "usually have far too many distinct values to group by."
                ),
            )

            if breakdown_choice == "(None)":
                breakdown_col = None
            elif column_kind_map.get(breakdown_choice) == "numeric":
                # Numeric columns can't sensibly be used to group/color a chart (too many
                # distinct values -> hundreds of tiny stacked slivers + a slow, useless
                # continuous color legend). Treat this as "sum this column" instead.
                breakdown_col = None
                st.caption(
                    f"ℹ️ `{breakdown_choice}` is numeric, so it'll be **summed**, not used to split the chart."
                )
                want_sum_choice = st.radio(
                    f"➕ Show the SUM of `{breakdown_choice}` for each month/year?",
                    options=["No", "Yes"],
                    index=0,
                    key=f"want_sum_{selected_table}_{breakdown_choice}",
                    horizontal=True,
                )
                if want_sum_choice == "Yes":
                    sum_columns = [breakdown_choice]
                    crores_choice = st.radio(
                        f"💰 Display the sum of `{breakdown_choice}` in Crores (÷ 1,00,00,000)?",
                        options=["No", "Yes"],
                        index=0,
                        key=f"crores_choice_{selected_table}_{breakdown_choice}",
                        horizontal=True,
                    )
                    show_in_crores = crores_choice == "Yes"
            else:
                # Categorical/text column -> normal group/color breakdown, same as before.
                breakdown_col = breakdown_choice

            try:
                available_ym_summary = get_available_year_months(engine, selected_table, summary_date_col)
            except SQLAlchemyError as exc:
                st.error(f"Failed to fetch available months: {exc}")
                available_ym_summary = []
            ym_labels_summary = {f"{MONTH_NAMES[mo - 1]} {yr}": (yr, mo) for yr, mo in available_ym_summary}
            focus_labels = st.multiselect(
                "Focus on specific month(s) — leave empty to see every month",
                options=list(ym_labels_summary.keys()),
                key=f"focus_months_{selected_table}_{summary_date_col}",
            )
            focus_year_months = [ym_labels_summary[lbl] for lbl in focus_labels]

            try:
                available_years_summary = get_available_years(engine, selected_table, summary_date_col)
            except SQLAlchemyError as exc:
                st.error(f"Failed to fetch available years: {exc}")
                available_years_summary = []
            focus_years_summary = st.multiselect(
                "Focus on specific year(s) — leave empty to see every year",
                options=available_years_summary,
                key=f"focus_years_summary_{selected_table}_{summary_date_col}",
            )
            if focus_year_months and focus_years_summary:
                st.caption(
                    "⚠️ Both a specific-month filter and a year filter are set — "
                    "the specific-month filter takes priority below. Clear it if you "
                    "just want the whole year(s) you picked."
                )

            if breakdown_col:
                try:
                    breakdown_values = get_distinct_values(engine, selected_table, breakdown_col)
                except SQLAlchemyError as exc:
                    st.error(f"Failed to fetch values for `{breakdown_col}`: {exc}")
                    breakdown_values = []
                focus_categories = st.multiselect(
                    f"Focus on specific {breakdown_col} value(s) — e.g. Cash — leave empty for all",
                    options=breakdown_values,
                    key=f"focus_categories_{selected_table}_{breakdown_col}",
                )

            chart_res_scale = st.select_slider(
                "🖼️ Chart clarity (export resolution)",
                options=[1, 2, 3, 4],
                value=3,
                key="chart_res_scale",
                help=(
                    "Base chart is 1600×800px. 3× ≈ 4800×2400px and 4× ≈ 6400×3200px — "
                    "both exceed 4K (3840×2160) for crisp downloads."
                ),
            )

            num_chart_splits = st.slider(
                "📊 Split the monthly chart into how many charts?",
                min_value=1,
                max_value=12,
                value=1,
                key="num_chart_splits",
                help=(
                    "If you have many months, one chart can get congested. "
                    "Raise this to break it into that many smaller charts "
                    "(e.g. 3 → roughly a third of the months per chart)."
                ),
            )

    st.divider()
    st.subheader("4️⃣ Summary table Analysis (optional)")
    show_pivot_table = st.checkbox(
        "Show monthly cumulative summary ",
        value=False,
        key=f"show_pivot_table_{selected_table}",
        help=(
            "Build a 4-quadrant Excel-style PivotTable: pick Filters, Rows, "
            "Columns, and a Value + aggregation, just like Excel's PivotTable."
        ),
    )

    pivot_filter_cols: list[str] = []
    pivot_filter_values: dict[str, list[Any]] = {}
    pivot_row_cols: list[str] = []
    pivot_col_cols: list[str] = []
    pivot_value_col: str | None = None
    pivot_agg_func: str = "sum"
    pivot_show_in_crores: bool = False
    pivot_right_total_mode: str = "grand_total"
    pivot_run_clicked = False
    pivot_is_built = st.session_state.get(f"pivot_built_{selected_table}", False)

    if show_pivot_table:
        numeric_cols = [c for c in column_names if column_kind_map.get(c) == "numeric"]

        st.markdown("**🔍 Filters**")
        pivot_filter_cols = st.multiselect(
            "Column(s) to filter the dataset before pivoting",
            options=column_names,
            key=f"pivot_filter_cols_{selected_table}",
        )
        for fc in pivot_filter_cols:
            try:
                fc_values = get_distinct_values(engine, selected_table, fc)
            except SQLAlchemyError as exc:
                st.error(f"Failed to fetch values for `{fc}`: {exc}")
                fc_values = []
            pivot_filter_values[fc] = st.multiselect(
                f"↳ Value(s) of `{fc}` — leave empty to include all",
                options=fc_values,
                key=f"pivot_filter_vals_{selected_table}_{fc}",
            )

        st.markdown("**≡ Rows**")
        pivot_row_cols = st.multiselect(
            "Column(s) to group vertically (e.g. ZONE, BRANCH)",
            options=[c for c in column_names if c not in pivot_filter_cols],
            key=f"pivot_row_cols_{selected_table}",
        )

        st.markdown("**|||| Columns**")
        pivot_col_cols = st.multiselect(
            "Column(s) to split horizontally across headers (e.g. STAGE_ECL, STATUS)",
            options=[c for c in column_names if c not in pivot_row_cols],
            key=f"pivot_col_cols_{selected_table}",
        )

        if len(pivot_col_cols) >= 2:
            deepest_label = pivot_col_cols[-1]
            outer_labels = ", ".join(pivot_col_cols[:-1])
            right_total_help = {
                "deepest_field_total": (
                    f"One '<value> Total' column per {deepest_label} (e.g. one per month), "
                    f"each summed across every {outer_labels} — the single overall Grand Total "
                    "column is dropped."
                ),
                "grand_total": "The original single 'Grand Total' column, summing everything together.",
                "both": (
                    f"The per-{deepest_label} Total columns, plus one overall Grand Total "
                    "at the very end."
                ),
            }
            pivot_right_total_mode = st.radio(
                f"↳ Rightmost total column(s) for `{deepest_label}`",
                options=["deepest_field_total", "grand_total", "both"],
                format_func=lambda k: {
                    "deepest_field_total": f"Per-{deepest_label} totals (new)",
                    "grand_total": "Single overall Grand Total (original)",
                    "both": "Both",
                }[k],
                index=0,
                key=f"pivot_right_total_mode_{selected_table}",
                help=right_total_help["deepest_field_total"],
            )
            st.caption(right_total_help[pivot_right_total_mode])
        else:
            pivot_right_total_mode = "grand_total"

        st.markdown("**∑ Values & Aggregation**")
        if numeric_cols:
            pivot_value_col = st.selectbox(
                "Numeric column to aggregate (e.g. GA in Crs)",
                options=numeric_cols,
                key=f"pivot_value_col_{selected_table}",
            )
        else:
            st.warning("No numeric columns found in this table to aggregate.")
        pivot_agg_func = st.selectbox(
            "Aggregation function",
            options=PIVOT_AGG_FUNCS,
            key=f"pivot_agg_func_{selected_table}",
        )

        if pivot_value_col and pivot_agg_func != "count":
            crores_pivot_choice = st.radio(
                f"💰 Display `{pivot_value_col}` in Crores (÷ 1,00,00,000) in the Summary Table?",
                options=["No", "Yes"],
                index=0,
                key=f"pivot_crores_{selected_table}_{pivot_value_col}",
                horizontal=True,
            )
            pivot_show_in_crores = crores_pivot_choice == "Yes"

        st.caption("A Grand Total row is always added. Column totals follow your choice above (if shown).")
        pivot_run_clicked = st.button(
            "🧮 Build Summary", type="primary", use_container_width=True, key=f"pivot_run_{selected_table}"
        )
        if pivot_run_clicked:
            # A button's True value only lasts for the one rerun it was clicked
            # on — the very next rerun (e.g. touching a widget in the ratio
            # calculator below, or tweaking Rows/Columns) it reports False
            # again. Persist "this table's pivot has been built" separately so
            # the whole Summary Table doesn't vanish and reappear stale on
            # every unrelated interaction; it now stays live and always
            # reflects the *current* Rows/Columns/Values choices.
            st.session_state[f"pivot_built_{selected_table}"] = True
        pivot_is_built = st.session_state.get(f"pivot_built_{selected_table}", False)

    st.divider()
    run_clicked = st.button("🚀 Generate Report", type="primary", use_container_width=True)

# --------------------------------------------------------------------------- #
# Main UI - Tabs Layout
# --------------------------------------------------------------------------- #
tab_report, tab_pivot, tab_sql = st.tabs(
    ["📊 Dynamic Report Builder", "🧮 Summary Table", "🧑‍💻 Custom SQL Query"]
)

with tab_report:
    # --------------------------------------------------------------------------- #
    # Monthly Cumulative Summary (independent of the report section below)
    # --------------------------------------------------------------------------- #
    if show_monthly_summary and summary_date_col:
        st.subheader("🗓️ Monthly Cumulative Summary")
        st.caption(
            f"Every record is counted under the calendar month of `{summary_date_col}` "
            "— e.g. a record dated 01-03-2024 is counted as a **March** record, regardless of the day."
        )

        try:
            with st.spinner("Building monthly summary..."):
                summary_df = get_monthly_summary(
                    engine, selected_table, summary_date_col, breakdown_col, tuple(sum_columns)
                )
        except SQLAlchemyError as exc:
            st.error(f"❌ Failed to build monthly summary: {exc}")
            summary_df = pd.DataFrame()
        except Exception as exc:  # noqa: BLE001
            st.error(f"❌ Unexpected error while building monthly summary: {exc}")
            summary_df = pd.DataFrame()

        if summary_df.empty:
            st.info("No data available to summarize for this date column.")
        else:
            summary_df["month_label"] = summary_df.apply(
                lambda r: f"{MONTH_NAMES[int(r['mo']) - 1]} {int(r['yr'])}", axis=1
            )
            summary_df = summary_df.sort_values(["yr", "mo"])

            # Apply optional focus filters chosen in the sidebar.
            # If specific month(s) are picked, that's the most precise filter and wins.
            # Otherwise, fall back to year-only focus so picking "2023" here also
            # narrows this monthly view down to just 2023 (previously this only
            # affected the Yearly section below, which made the monthly chart/table/
            # CSV look "unfiltered" even after picking a year).
            display_df = summary_df.copy()
            if focus_year_months:
                focus_set = {(int(yr), int(mo)) for yr, mo in focus_year_months}
                display_df = display_df[
                    display_df.apply(lambda r: (int(r["yr"]), int(r["mo"])) in focus_set, axis=1)
                ]
            elif focus_years_summary:
                focus_years_set = {int(y) for y in focus_years_summary}
                display_df = display_df[display_df["yr"].astype(int).isin(focus_years_set)]
            if breakdown_col and focus_categories:
                display_df = display_df[display_df[breakdown_col].isin(focus_categories)]

            if display_df.empty:
                st.warning(
                    "No records match the selected month/category focus. "
                    f"Active filters — Month(s): {focus_labels or 'none'}, "
                    f"Year(s): {focus_years_summary or 'none'}, "
                    f"{breakdown_col or 'category'}: {focus_categories or 'none'}. "
                    "Clear one of these in the sidebar to see results."
                )
            elif breakdown_col:
                # --- Grouped comparison chart: month x category (with value labels) ---
                render_split_bar_charts(
                    display_df,
                    x_col="month_label",
                    y_col="record_count",
                    color_col=breakdown_col,
                    title=f"Records per month by {breakdown_col}",
                    filename_prefix=f"{selected_table}_monthly_{breakdown_col}",
                    key_suffix="monthly_breakdown",
                    res_scale=chart_res_scale,
                    num_splits=num_chart_splits,
                )

                # --- Excel-style pivot: rows = month, columns = category ---
                pivot_df = display_df.pivot_table(
                    index="month_label",
                    columns=breakdown_col,
                    values="record_count",
                    aggfunc="sum",
                    fill_value=0,
                    margins=True,
                    margins_name="Total",
                )
                month_order = [
                    m for m in display_df.sort_values(["yr", "mo"])["month_label"].unique() if m in pivot_df.index
                ]
                if "Total" in pivot_df.index:
                    month_order.append("Total")
                pivot_df = pivot_df.reindex(month_order)
                st.dataframe(pivot_df, use_container_width=True)

                summary_csv_buffer = io.StringIO()
                pivot_df.to_csv(summary_csv_buffer)
                st.download_button(
                    "⬇️ Download Monthly Breakdown CSV",
                    data=summary_csv_buffer.getvalue(),
                    file_name=f"{selected_table}_monthly_{breakdown_col}_summary.csv",
                    mime="text/csv",
                    key="download_monthly_breakdown",
                )

                # --- Plain row-by-row data (Month | category | Count) ---------
                with st.expander("📋 View all rows (Month × " + breakdown_col + ")", expanded=False):
                    flat_df = display_df[["month_label", breakdown_col, "record_count"]].rename(
                        columns={"month_label": "Month", "record_count": "Count"}
                    )
                    st.dataframe(flat_df, use_container_width=True, hide_index=True)
                    flat_csv_buffer = io.StringIO()
                    flat_df.to_csv(flat_csv_buffer, index=False)
                    st.download_button(
                        "⬇️ Download Filtered Rows CSV",
                        data=flat_csv_buffer.getvalue(),
                        file_name=f"{selected_table}_monthly_{breakdown_col}_rows.csv",
                        mime="text/csv",
                        key="download_monthly_breakdown_rows",
                    )
            else:
                # --- Simple month-over-month comparison (with value labels) ---
                render_split_bar_charts(
                    display_df,
                    x_col="month_label",
                    y_col="record_count",
                    color_col=None,
                    title="Records per month",
                    filename_prefix=f"{selected_table}_monthly_summary",
                    key_suffix="monthly_simple",
                    res_scale=chart_res_scale,
                    num_splits=num_chart_splits,
                )

                if focus_year_months:
                    cols = st.columns(min(len(display_df), 4) or 1)
                    for i, (_, row) in enumerate(display_df.iterrows()):
                        cols[i % len(cols)].metric(row["month_label"], f"{int(row['record_count']):,}")

                st.dataframe(
                    display_df[["month_label", "record_count"]].rename(
                        columns={"month_label": "Month", "record_count": "Count"}
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

                summary_csv_buffer = io.StringIO()
                display_df[["month_label", "record_count"]].to_csv(summary_csv_buffer, index=False)
                st.download_button(
                    "⬇️ Download Monthly Summary CSV",
                    data=summary_csv_buffer.getvalue(),
                    file_name=f"{selected_table}_monthly_summary.csv",
                    mime="text/csv",
                    key="download_monthly_summary",
                )

            # --- Optional SUM columns (e.g. "emi") per month ------------------- #
            if not display_df.empty and sum_columns:
                unit_note = " (shown in Crores)" if show_in_crores else ""
                for sc in sum_columns:
                    sum_field = f"sum_{sc}"
                    if sum_field not in display_df.columns:
                        continue

                    st.markdown(f"#### 💵 Sum of `{sc}` per month{unit_note}")

                    chart_df = display_df.copy()
                    if show_in_crores:
                        chart_df[sum_field] = chart_df[sum_field] / CRORE
                    y_axis_label = f"Sum of {sc}" + (" (Cr)" if show_in_crores else "")
                    bar_value_format = ",.2f" if show_in_crores else ",.0f"

                    if breakdown_col:
                        render_split_bar_charts(
                            chart_df,
                            x_col="month_label",
                            y_col=sum_field,
                            color_col=breakdown_col,
                            title=f"Sum of {sc} per month by {breakdown_col}{unit_note}",
                            filename_prefix=f"{selected_table}_monthly_sum_{sc}_{breakdown_col}",
                            key_suffix=f"monthly_sum_{sc}_breakdown",
                            res_scale=chart_res_scale,
                            num_splits=num_chart_splits,
                            y_label=y_axis_label,
                            value_format=bar_value_format,
                        )

                        sum_pivot_df = display_df.pivot_table(
                            index="month_label",
                            columns=breakdown_col,
                            values=sum_field,
                            aggfunc="sum",
                            fill_value=0,
                            margins=True,
                            margins_name="Total",
                        )
                        sum_pivot_df = sum_pivot_df.reindex(
                            [m for m in month_order if m in sum_pivot_df.index]
                        )
                        sum_pivot_view = (
                            sum_pivot_df.div(CRORE) if show_in_crores else sum_pivot_df
                        ).round(2)
                        st.dataframe(sum_pivot_view, use_container_width=True)
                        if show_in_crores:
                            st.caption("Table above is in Crores. Downloaded CSV keeps raw (non-Crore) values.")

                        sum_csv_buffer = io.StringIO()
                        sum_pivot_df.to_csv(sum_csv_buffer)  # always raw values in CSV
                        st.download_button(
                            f"⬇️ Download Monthly Sum({sc}) Breakdown CSV",
                            data=sum_csv_buffer.getvalue(),
                            file_name=f"{selected_table}_monthly_sum_{sc}_{breakdown_col}.csv",
                            mime="text/csv",
                            key=f"download_monthly_sum_{sc}_breakdown",
                        )
                    else:
                        render_split_bar_charts(
                            chart_df,
                            x_col="month_label",
                            y_col=sum_field,
                            color_col=None,
                            title=f"Sum of {sc} per month{unit_note}",
                            filename_prefix=f"{selected_table}_monthly_sum_{sc}",
                            key_suffix=f"monthly_sum_{sc}",
                            res_scale=chart_res_scale,
                            num_splits=num_chart_splits,
                            y_label=y_axis_label,
                            value_format=bar_value_format,
                        )

                        if focus_year_months:
                            m_cols = st.columns(min(len(display_df), 4) or 1)
                            for i, (_, row) in enumerate(display_df.iterrows()):
                                m_cols[i % len(m_cols)].metric(
                                    row["month_label"], format_amount(row[sum_field], show_in_crores)
                                )

                        sum_table_df = display_df[["month_label", sum_field]].rename(
                            columns={"month_label": "Month", sum_field: f"Sum({sc})"}
                        )
                        table_view = sum_table_df.copy()
                        if show_in_crores:
                            table_view[f"Sum({sc}) [Cr]"] = series_to_crores(table_view[f"Sum({sc})"]).round(2)
                            table_view = table_view.drop(columns=[f"Sum({sc})"])
                        st.dataframe(table_view, use_container_width=True, hide_index=True)

                        sum_csv_buffer = io.StringIO()
                        sum_table_df.to_csv(sum_csv_buffer, index=False)  # raw values in CSV
                        st.download_button(
                            f"⬇️ Download Monthly Sum({sc}) CSV",
                            data=sum_csv_buffer.getvalue(),
                            file_name=f"{selected_table}_monthly_sum_{sc}.csv",
                            mime="text/csv",
                            key=f"download_monthly_sum_{sc}",
                        )
                    st.divider()
        st.divider()

        # --------------------------------------------------------------------------- #
        # Yearly Cumulative Summary — same data, rolled up to 2023 / 2024 / 2025 ...
        # --------------------------------------------------------------------------- #
        st.subheader("📅 Yearly Cumulative Summary")
        st.caption(
            f"Same data as above, rolled up to the calendar year of `{summary_date_col}` "
            "(e.g. 2023, 2024, 2025)."
        )

        group_cols = ["yr"] + ([breakdown_col] if breakdown_col else [])
        yearly_agg_cols = ["record_count"] + [f"sum_{sc}" for sc in sum_columns if f"sum_{sc}" in summary_df.columns]
        yearly_df = summary_df.groupby(group_cols, as_index=False)[yearly_agg_cols].sum()
        yearly_df["year_label"] = yearly_df["yr"].astype(int).astype(str)
        yearly_df = yearly_df.sort_values("yr")

        yearly_display_df = yearly_df.copy()
        if focus_years_summary:
            focus_years_set = {int(y) for y in focus_years_summary}
            yearly_display_df = yearly_display_df[yearly_display_df["yr"].astype(int).isin(focus_years_set)]
        if breakdown_col and focus_categories:
            yearly_display_df = yearly_display_df[yearly_display_df[breakdown_col].isin(focus_categories)]

        if yearly_display_df.empty:
            st.warning(
                "No records match the selected year/category focus. "
                f"Active filters — Year(s): {focus_years_summary or 'none'}, "
                f"{breakdown_col or 'category'}: {focus_categories or 'none'}. "
                "Clear one of these in the sidebar to see results."
            )
        elif breakdown_col:
            render_split_bar_charts(
                yearly_display_df,
                x_col="year_label",
                y_col="record_count",
                color_col=breakdown_col,
                title=f"Records per year by {breakdown_col}",
                filename_prefix=f"{selected_table}_yearly_{breakdown_col}",
                key_suffix="yearly_breakdown",
                res_scale=chart_res_scale,
                num_splits=num_chart_splits,
            )

            yearly_pivot_df = yearly_display_df.pivot_table(
                index="year_label",
                columns=breakdown_col,
                values="record_count",
                aggfunc="sum",
                fill_value=0,
                margins=True,
                margins_name="Total",
            )
            year_order = [
                y for y in yearly_display_df.sort_values("yr")["year_label"].unique() if y in yearly_pivot_df.index
            ]
            if "Total" in yearly_pivot_df.index:
                year_order.append("Total")
            yearly_pivot_df = yearly_pivot_df.reindex(year_order)
            st.dataframe(yearly_pivot_df, use_container_width=True)

            yearly_csv_buffer = io.StringIO()
            yearly_pivot_df.to_csv(yearly_csv_buffer)
            st.download_button(
                "⬇️ Download Yearly Breakdown CSV",
                data=yearly_csv_buffer.getvalue(),
                file_name=f"{selected_table}_yearly_{breakdown_col}_summary.csv",
                mime="text/csv",
                key="download_yearly_breakdown",
            )

            # --- Plain row-by-row data (Year | category | Count) -------------
            with st.expander("📋 View all rows (Year × " + breakdown_col + ")", expanded=False):
                yearly_flat_df = yearly_display_df[["year_label", breakdown_col, "record_count"]].rename(
                    columns={"year_label": "Year", "record_count": "Count"}
                )
                st.dataframe(yearly_flat_df, use_container_width=True, hide_index=True)
                yearly_flat_csv_buffer = io.StringIO()
                yearly_flat_df.to_csv(yearly_flat_csv_buffer, index=False)
                st.download_button(
                    "⬇️ Download Filtered Rows CSV",
                    data=yearly_flat_csv_buffer.getvalue(),
                    file_name=f"{selected_table}_yearly_{breakdown_col}_rows.csv",
                    mime="text/csv",
                    key="download_yearly_breakdown_rows",
                )
        else:
            render_split_bar_charts(
                yearly_display_df,
                x_col="year_label",
                y_col="record_count",
                color_col=None,
                title="Records per year",
                filename_prefix=f"{selected_table}_yearly_summary",
                key_suffix="yearly_simple",
                res_scale=chart_res_scale,
                num_splits=num_chart_splits,
            )

            if focus_years_summary:
                cols = st.columns(min(len(yearly_display_df), 4) or 1)
                for i, (_, row) in enumerate(yearly_display_df.iterrows()):
                    cols[i % len(cols)].metric(row["year_label"], f"{int(row['record_count']):,}")

            st.dataframe(
                yearly_display_df[["year_label", "record_count"]].rename(
                    columns={"year_label": "Year", "record_count": "Count"}
                ),
                use_container_width=True,
                hide_index=True,
            )

            yearly_csv_buffer = io.StringIO()
            yearly_display_df[["year_label", "record_count"]].to_csv(yearly_csv_buffer, index=False)
            st.download_button(
                "⬇️ Download Yearly Summary CSV",
                data=yearly_csv_buffer.getvalue(),
                file_name=f"{selected_table}_yearly_summary.csv",
                mime="text/csv",
                key="download_yearly_summary",
            )

        # --- Optional SUM columns (e.g. "emi") per year ------------------- #
        if not yearly_display_df.empty and sum_columns:
            unit_note = " (shown in Crores)" if show_in_crores else ""
            for sc in sum_columns:
                sum_field = f"sum_{sc}"
                if sum_field not in yearly_display_df.columns:
                    continue

                st.markdown(f"#### 💵 Sum of `{sc}` per year{unit_note}")

                yearly_chart_df = yearly_display_df.copy()
                if show_in_crores:
                    yearly_chart_df[sum_field] = yearly_chart_df[sum_field] / CRORE
                yearly_y_axis_label = f"Sum of {sc}" + (" (Cr)" if show_in_crores else "")
                yearly_bar_value_format = ",.2f" if show_in_crores else ",.0f"

                if breakdown_col:
                    render_split_bar_charts(
                        yearly_chart_df,
                        x_col="year_label",
                        y_col=sum_field,
                        color_col=breakdown_col,
                        title=f"Sum of {sc} per year by {breakdown_col}{unit_note}",
                        filename_prefix=f"{selected_table}_yearly_sum_{sc}_{breakdown_col}",
                        key_suffix=f"yearly_sum_{sc}_breakdown",
                        res_scale=chart_res_scale,
                        num_splits=num_chart_splits,
                        y_label=yearly_y_axis_label,
                        value_format=yearly_bar_value_format,
                    )

                    yearly_sum_pivot_df = yearly_display_df.pivot_table(
                        index="year_label",
                        columns=breakdown_col,
                        values=sum_field,
                        aggfunc="sum",
                        fill_value=0,
                        margins=True,
                        margins_name="Total",
                    )
                    yearly_sum_pivot_df = yearly_sum_pivot_df.reindex(
                        [y for y in year_order if y in yearly_sum_pivot_df.index]
                    )
                    yearly_sum_pivot_view = (
                        yearly_sum_pivot_df.div(CRORE) if show_in_crores else yearly_sum_pivot_df
                    ).round(2)
                    st.dataframe(yearly_sum_pivot_view, use_container_width=True)
                    if show_in_crores:
                        st.caption("Table above is in Crores. Downloaded CSV keeps raw (non-Crore) values.")

                    yearly_sum_csv_buffer = io.StringIO()
                    yearly_sum_pivot_df.to_csv(yearly_sum_csv_buffer)  # raw values
                    st.download_button(
                        f"⬇️ Download Yearly Sum({sc}) Breakdown CSV",
                        data=yearly_sum_csv_buffer.getvalue(),
                        file_name=f"{selected_table}_yearly_sum_{sc}_{breakdown_col}.csv",
                        mime="text/csv",
                        key=f"download_yearly_sum_{sc}_breakdown",
                    )
                else:
                    render_split_bar_charts(
                        yearly_chart_df,
                        x_col="year_label",
                        y_col=sum_field,
                        color_col=None,
                        title=f"Sum of {sc} per year{unit_note}",
                        filename_prefix=f"{selected_table}_yearly_sum_{sc}",
                        key_suffix=f"yearly_sum_{sc}",
                        res_scale=chart_res_scale,
                        num_splits=num_chart_splits,
                        y_label=yearly_y_axis_label,
                        value_format=yearly_bar_value_format,
                    )

                    if focus_years_summary:
                        y_cols = st.columns(min(len(yearly_display_df), 4) or 1)
                        for i, (_, row) in enumerate(yearly_display_df.iterrows()):
                            y_cols[i % len(y_cols)].metric(
                                row["year_label"], format_amount(row[sum_field], show_in_crores)
                            )

                    yearly_sum_table_df = yearly_display_df[["year_label", sum_field]].rename(
                        columns={"year_label": "Year", sum_field: f"Sum({sc})"}
                    )
                    yearly_table_view = yearly_sum_table_df.copy()
                    if show_in_crores:
                        yearly_table_view[f"Sum({sc}) [Cr]"] = series_to_crores(
                            yearly_table_view[f"Sum({sc})"]
                        ).round(2)
                        yearly_table_view = yearly_table_view.drop(columns=[f"Sum({sc})"])
                    st.dataframe(yearly_table_view, use_container_width=True, hide_index=True)

                    yearly_sum_csv_buffer = io.StringIO()
                    yearly_sum_table_df.to_csv(yearly_sum_csv_buffer, index=False)  # raw values
                    st.download_button(
                        f"⬇️ Download Yearly Sum({sc}) CSV",
                        data=yearly_sum_csv_buffer.getvalue(),
                        file_name=f"{selected_table}_yearly_sum_{sc}.csv",
                        mime="text/csv",
                        key=f"download_yearly_sum_{sc}",
                    )
                st.divider()
        st.divider()

    # --------------------------------------------------------------------------- #
    # Main panel - results
    # --------------------------------------------------------------------------- #
    if not output_columns:
        st.info("👈 Select at least one output column in the sidebar to build a report.")
    elif not run_clicked:
        st.info("👈 Configure your filters and click **Generate Report** to run the query.")
    else:
        query, params = build_query(selected_table, ref_column, output_columns, ref_kind, filter_payload)

        with st.expander("🔍 View generated SQL", expanded=False):
            st.code(query, language="sql")
            if params:
                st.caption("Parameters:")
                st.json({k: str(v) for k, v in params.items()})

        try:
            with st.spinner("Running query..."):
                df = run_query(engine, query, params)
                
            st.success(f"Query returned **{len(df):,}** row(s).")

            # --- KPI cards for numeric output columns ---
            numeric_output_cols = [
                c for c in output_columns
                if c in df.columns and pd.api.types.is_numeric_dtype(df[c])
            ]

            if numeric_output_cols:
                st.subheader("📈 KPI Summary")
                if show_in_crores:
                    st.caption("💰 Sum/Average shown in Crores (÷1,00,00,000).")
                for col in numeric_output_cols:
                    series = df[col].dropna()
                    c1, c2, c3 = st.columns(3)
                    c1.metric(f"Total Sum — {col}", format_amount(series.sum(), show_in_crores))
                    c2.metric(
                        f"Average — {col}",
                        format_amount(series.mean(), show_in_crores) if len(series) else format_amount(0, show_in_crores),
                    )
                    c3.metric(f"Count — {col}", f"{series.count():,}")
                st.divider()

            # --- Data table ---
            st.subheader("📋 Results")
            st.dataframe(df, use_container_width=True, height=450)
            st.caption(f"Total rows: **{len(df):,}** | Columns: {', '.join(df.columns)}")

            # --- Download button ---
            csv_buffer = io.StringIO()
            df.to_csv(csv_buffer, index=False)
            st.download_button(
                label="⬇️ Download CSV",
                data=csv_buffer.getvalue(),
                file_name=f"{selected_table}_report.csv",
                mime="text/csv",
                use_container_width=True,
            )
        except SQLAlchemyError as exc:
            st.error(f"❌ Query failed: {exc}")
        except Exception as exc:  # noqa: BLE001
            st.error(f"❌ Unexpected error while running the query: {exc}")

with tab_pivot:
    st.subheader("🧮 Quadrant Summary Table")

    if not show_pivot_table:
        st.info(
            "👈 Enable **'Show monthly cumulative summary / Pivot Table'** in the sidebar "
            "(section 4️⃣) to build a pivot table."
        )
    elif not pivot_row_cols:
        st.info("👈 Choose at least one **Rows** column in the sidebar to build a pivot table.")
    elif not pivot_value_col:
        st.info("👈 Choose a numeric **Values** column in the sidebar to aggregate.")
    elif not pivot_is_built:
        st.info("👈 Configure Filters / Rows / Columns / Values in the sidebar, then click **Build Summary** once. It'll then stay live and update automatically as you change your selections or use the calculator below.")
    else:
        try:
            columns_needed = tuple(
                dict.fromkeys([*pivot_filter_cols, *pivot_row_cols, *pivot_col_cols, pivot_value_col])
            )
            filter_conditions = tuple(
                (fc, tuple(pivot_filter_values.get(fc, []))) for fc in pivot_filter_cols
            )

            with st.spinner("Fetching data for pivot table..."):
                pivot_source_df = fetch_pivot_source_data(
                    engine, selected_table, columns_needed, filter_conditions
                )

            if pivot_source_df.empty:
                st.warning("No rows match the selected filters — nothing to pivot.")
            else:
                with st.spinner("Building pivot table..."):
                    pivot_df = build_pivot_table(
                        pivot_source_df, pivot_row_cols, pivot_col_cols, pivot_value_col, pivot_agg_func
                    )

                st.success(
                    f"Pivot built from **{len(pivot_source_df):,}** source row(s) — "
                    f"**{pivot_agg_func}({pivot_value_col})** by **{', '.join(pivot_row_cols)}**"
                    + (f" × **{', '.join(pivot_col_cols)}**" if pivot_col_cols else "")
                )

                st.markdown("#### 📋 Summary Matrix")
                display_pivot_df = render_excel_style_pivot_table(
                    pivot_source_df,
                    pivot_row_cols,
                    pivot_col_cols,
                    pivot_value_col,
                    pivot_agg_func,
                    column_kind_map,
                    pivot_show_in_crores,
                    pivot_right_total_mode,
                )

                available_efficiency_fields = list(dict.fromkeys([*pivot_row_cols, *pivot_col_cols]))
                if available_efficiency_fields:
                    st.markdown("#### 🎯 Field Efficiency Table (% of a chosen field's group total)")
                    st.caption(
                        "Pick any one of your selected Rows/Columns fields below — the table will show "
                        "each cell as that field's % share of the total across its own group, holding "
                        "every other selected field fixed. Works the same way whether you point it at a "
                        "Rows field or a Columns field."
                    )
                    normalize_field = st.selectbox(
                        "Normalize as % across:",
                        options=available_efficiency_fields,
                        key=f"pivot_normalize_field_{selected_table}",
                    )
                    efficiency_display_df = render_field_efficiency_table(
                        pivot_source_df,
                        pivot_row_cols,
                        pivot_col_cols,
                        normalize_field,
                        pivot_value_col,
                        pivot_agg_func,
                    )
                    efficiency_csv_bytes = efficiency_display_df.to_csv().encode("utf-8-sig")
                    st.download_button(
                        "⬇️ Download Field Efficiency Table CSV",
                        data=efficiency_csv_bytes,
                        file_name=f"{selected_table}_pivot_efficiency_{normalize_field}.csv",
                        mime="text/csv",
                        use_container_width=True,
                        key=f"download_pivot_efficiency_{selected_table}",
                    )
                else:
                    efficiency_display_df = None

                st.markdown("#### 📈 Chart")
                render_pivot_chart(
                    pivot_df,
                    pivot_row_cols,
                    pivot_col_cols,
                    pivot_value_col,
                    pivot_agg_func,
                    key_suffix=selected_table,
                )

                dl_col1, dl_col2 = st.columns(2)
                with dl_col1:
                    pivot_csv_bytes = pivot_df.to_csv().encode("utf-8-sig")
                    st.download_button(
                        "⬇️ Download Pivot CSV (flat)",
                        data=pivot_csv_bytes,
                        file_name=f"{selected_table}_pivot_summary.csv",
                        mime="text/csv",
                        use_container_width=True,
                        key=f"download_pivot_{selected_table}",
                    )
                with dl_col2:
                    # utf-8-sig adds a BOM so Excel detects UTF-8 correctly instead of
                    # mangling non-ASCII characters like the "↳" indent arrow.
                    display_csv_bytes = display_pivot_df.to_csv().encode("utf-8-sig")
                    st.download_button(
                        "⬇️ Download Summary Matrix CSV (with subtotals)",
                        data=display_csv_bytes,
                        file_name=f"{selected_table}_pivot_excel_style.csv",
                        mime="text/csv",
                        use_container_width=True,
                        key=f"download_pivot_excel_style_{selected_table}",
                    )
        except SQLAlchemyError as exc:
            st.error(f"❌ Failed to fetch data for the pivot table: {exc}")
        except Exception as exc:  # noqa: BLE001
            st.error(f"❌ Unexpected error while building the pivot table: {exc}")

with tab_sql:
    st.subheader("🧑‍💻 Custom SQL Execution")
    st.info("Write and execute standard PostgreSQL queries directly against your database.")
    
    # Text area for the user to type queries
    custom_query = st.text_area("SQL Query", height=200, placeholder='SELECT * FROM "public"."agreements" LIMIT 100;')
    
    if st.button("▶️ Run Custom Query", type="primary"):
        if custom_query.strip():
            try:
                with st.spinner("Executing query..."):
                    with engine.connect() as conn:
                        # Uses the existing engine to run the custom text query
                        custom_df = pd.read_sql(text(custom_query), conn)
                
                st.success(f"Query returned **{len(custom_df):,}** row(s).")
                st.dataframe(custom_df, use_container_width=True, height=400)
                
                # Dynamic CSV Download for Custom Queries
                custom_csv = io.StringIO()
                custom_df.to_csv(custom_csv, index=False)
                st.download_button(
                    label="⬇️ Download Custom Results (CSV)",
                    data=custom_csv.getvalue(),
                    file_name="custom_sql_results.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"❌ Query failed: {exc}")
        else:
            st.warning("Please enter a SQL query first.")