"""
PostgreSQL BI Analytics Dashboard
==================================
A production-ready Streamlit BI tool that connects to PostgreSQL, lets a
user apply an unlimited number of simultaneous column filters (date range /
month-year / year, numeric range sliders, categorical multi-select), then
explore the filtered dataset through three views:

    1. Filtered Raw Data   - searchable table, KPI cards, CSV export
    2. Interactive Pivot   - Excel-style pivot_table with row/col totals
    3. Charts & Analytics  - Plotly Express charts (bar/line/area/pie/treemap)

Run with:
    streamlit run app.py

Configuration is read from environment variables or st.secrets
(see .env.example / .streamlit/secrets.toml.example / README).
"""

from __future__ import annotations

import io
from datetime import datetime
import os
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
    page_title="BI Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --------------------------------------------------------------------------- #
# Configuration / Environment
# --------------------------------------------------------------------------- #
# Priority: st.secrets (Streamlit Cloud / secrets.toml) -> environment vars.


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


@st.cache_data(ttl=300, show_spinner=False)
def get_column_min_max(_engine: Engine, table_name: str, column_name: str) -> tuple[Any, Any]:
    """Fetch MIN/MAX for a date or numeric column, safely handling all-NULL columns."""
    query = text(f'SELECT MIN("{column_name}"), MAX("{column_name}") FROM "public"."{table_name}"')
    with _engine.connect() as conn:
        row = conn.execute(query).first()
    if row is None:
        return None, None
    lo, hi = row[0], row[1]
    # Normalize datetimes to plain dates for st.date_input / st.slider friendliness.
    if isinstance(lo, datetime):
        lo = lo.date()
    if isinstance(hi, datetime):
        hi = hi.date()
    return lo, hi


MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


# --------------------------------------------------------------------------- #
# Multi-column filter engine
# --------------------------------------------------------------------------- #
def render_date_filter(
    engine: Engine, table: str, col: str, idx: int
) -> tuple[str | None, dict[str, Any]]:
    """Render sidebar widgets for one date/timestamp filter column and return
    a (SQL condition, params) pair, or (None, {}) if the user set no filter."""
    mode = st.sidebar.radio(
        f"`{col}` — mode",
        options=["Date Range", "Month/Year", "Year(s)"],
        key=f"mode_{table}_{col}",
        horizontal=True,
    )

    if mode == "Date Range":
        lo, hi = get_column_min_max(engine, table, col)
        if lo is None or hi is None:
            st.sidebar.caption("No data available to determine a date range.")
            return None, {}
        picked = st.sidebar.date_input(
            f"`{col}` between",
            value=(lo, hi),
            min_value=lo,
            max_value=hi,
            key=f"daterange_{table}_{col}",
        )
        if isinstance(picked, (list, tuple)) and len(picked) == 2:
            start, end = picked
            cond = f'"{col}" BETWEEN :f{idx}_start AND :f{idx}_end'
            return cond, {f"f{idx}_start": start, f"f{idx}_end": end}
        return None, {}

    if mode == "Month/Year":
        try:
            available_ym = get_available_year_months(engine, table, col)
        except SQLAlchemyError as exc:
            st.sidebar.error(f"Could not load months for `{col}`: {exc}")
            return None, {}
        ym_labels = {f"{MONTH_NAMES[mo - 1]} {yr}": (yr, mo) for yr, mo in available_ym}
        chosen = st.sidebar.multiselect(
            f"`{col}` — select Month/Year", options=list(ym_labels.keys()), key=f"ym_{table}_{col}"
        )
        if not chosen:
            return None, {}
        conditions, params = [], {}
        for i, label in enumerate(chosen):
            yr, mo = ym_labels[label]
            yr_key, mo_key = f"f{idx}_yr_{i}", f"f{idx}_mo_{i}"
            conditions.append(
                f'(EXTRACT(YEAR FROM "{col}") = :{yr_key} AND EXTRACT(MONTH FROM "{col}") = :{mo_key})'
            )
            params[yr_key] = yr
            params[mo_key] = mo
        return "(" + " OR ".join(conditions) + ")", params

    # mode == "Year(s)"
    try:
        available_years = get_available_years(engine, table, col)
    except SQLAlchemyError as exc:
        st.sidebar.error(f"Could not load years for `{col}`: {exc}")
        return None, {}
    chosen_years = st.sidebar.multiselect(
        f"`{col}` — select Year(s)", options=available_years, key=f"years_{table}_{col}"
    )
    if not chosen_years:
        return None, {}
    placeholders, params = [], {}
    for i, yr in enumerate(chosen_years):
        key = f"f{idx}_year_{i}"
        placeholders.append(f":{key}")
        params[key] = yr
    return f'EXTRACT(YEAR FROM "{col}") IN ({", ".join(placeholders)})', params


def render_numeric_filter(
    engine: Engine, table: str, col: str, idx: int
) -> tuple[str | None, dict[str, Any]]:
    """Render a range slider for a numeric filter column."""
    lo, hi = get_column_min_max(engine, table, col)
    if lo is None or hi is None:
        st.sidebar.caption(f"`{col}` has no non-null values to filter on.")
        return None, {}
    lo_f, hi_f = float(lo), float(hi)
    if lo_f == hi_f:
        st.sidebar.caption(f"`{col}` is constant ({lo_f:,.2f}) — no range to filter.")
        return None, {}
    chosen = st.sidebar.slider(
        f"`{col}` range",
        min_value=lo_f,
        max_value=hi_f,
        value=(lo_f, hi_f),
        key=f"numrange_{table}_{col}",
    )
    # Only add a WHERE clause if the user actually narrowed the full range.
    if chosen[0] <= lo_f and chosen[1] >= hi_f:
        return None, {}
    cond = f'"{col}" BETWEEN :f{idx}_min AND :f{idx}_max'
    return cond, {f"f{idx}_min": chosen[0], f"f{idx}_max": chosen[1]}


def render_categorical_filter(
    engine: Engine, table: str, col: str, idx: int
) -> tuple[str | None, dict[str, Any]]:
    """Render a multi-select for a categorical/text filter column."""
    try:
        distinct_values = get_distinct_values(engine, table, col)
    except SQLAlchemyError as exc:
        st.sidebar.error(f"Could not load values for `{col}`: {exc}")
        return None, {}
    if len(distinct_values) >= 1000:
        st.sidebar.caption(f"`{col}` has 1000+ distinct values — showing first 1000.")
    chosen = st.sidebar.multiselect(f"`{col}` values", options=distinct_values, key=f"cat_{table}_{col}")
    if not chosen:
        return None, {}
    placeholders, params = [], {}
    for i, v in enumerate(chosen):
        key = f"f{idx}_v{i}"
        placeholders.append(f":{key}")
        params[key] = v
    return f'"{col}" IN ({", ".join(placeholders)})', params


def build_filtered_query(
    table: str, filter_columns: list[str], column_kind_map: dict[str, str], engine: Engine, row_limit: int
) -> tuple[str, dict[str, Any]]:
    """
    Render one filter widget per selected column (unlimited columns) and
    build a single parameterized SQL query combining all active filters
    with AND. Returns the full SELECT * query and its bound parameters.
    """
    conditions: list[str] = []
    params: dict[str, Any] = {}

    for idx, col in enumerate(filter_columns):
        kind = column_kind_map.get(col, "text")
        st.sidebar.markdown(f"**🔸 {col}**")
        if kind == "date":
            cond, p = render_date_filter(engine, table, col, idx)
        elif kind == "numeric":
            cond, p = render_numeric_filter(engine, table, col, idx)
        else:
            cond, p = render_categorical_filter(engine, table, col, idx)
        if cond:
            conditions.append(cond)
            params.update(p)

    where_sql = ("WHERE " + " AND ".join(conditions)) if conditions else ""
    query = f'SELECT * FROM "public"."{table}" {where_sql} LIMIT :row_limit'.strip()
    params["row_limit"] = row_limit
    return query, params


@st.cache_data(ttl=60, show_spinner=False)
def run_query(_engine: Engine, query: str, params: dict[str, Any]) -> pd.DataFrame:
    """Execute the (already-parameterized) query and return a DataFrame."""
    with _engine.connect() as conn:
        return pd.read_sql(text(query), conn, params=params)


# --------------------------------------------------------------------------- #
# UI - Header & connection
# --------------------------------------------------------------------------- #
st.title("📊 BI Analytics Dashboard")
st.caption("Filter any combination of columns, pivot your data Excel-style, and build interactive charts.")

engine = get_engine_safe()
if engine is None:
    st.info(
        "Set your database credentials as environment variables (or in "
        "`.streamlit/secrets.toml`) and reload the app. See the README / "
        "setup instructions for details."
    )
    st.stop()

with st.sidebar:
    st.header("⚙️ Data Source")
    st.success(f"Connected to **{DB_NAME}**@`{DB_HOST}`", icon="✅")

    try:
        tables = list_tables(engine)
    except SQLAlchemyError as exc:
        st.error(f"Failed to list tables: {exc}")
        st.stop()

    if not tables:
        st.warning("No tables found in the `public` schema.")
        st.stop()

    selected_table = st.selectbox("📁 Select a table", options=tables, index=0)

# --------------------------------------------------------------------------- #
# Load column metadata for the selected table
# --------------------------------------------------------------------------- #
try:
    columns_meta = get_columns(engine, selected_table)
except SQLAlchemyError as exc:
    st.error(f"Failed to fetch columns for `{selected_table}`: {exc}")
    st.stop()

if not columns_meta:
    st.warning(f"Table `{selected_table}` has no readable columns.")
    st.stop()

column_names = [c["name"] for c in columns_meta]
column_kind_map = {c["name"]: classify_column(c["type"]) for c in columns_meta}

with st.sidebar:
    st.divider()
    st.header("🧭 Filters")
    st.caption("Add as many filters as you like, across any columns.")
    filter_columns = st.multiselect(
        "Filter on these columns", options=column_names, key=f"filtercols_{selected_table}"
    )

    row_limit = st.number_input(
        "Max rows to fetch (safety cap)",
        min_value=1000,
        max_value=1_000_000,
        value=50_000,
        step=1000,
        help="Protects against accidentally loading an enormous table into memory.",
    )

    if filter_columns:
        st.divider()
        query, params = build_filtered_query(selected_table, filter_columns, column_kind_map, engine, row_limit)
    else:
        query = f'SELECT * FROM "public"."{selected_table}" LIMIT :row_limit'
        params = {"row_limit": row_limit}

    st.divider()
    run_clicked = st.button("🚀 Run / Refresh Dashboard", type="primary", use_container_width=True)

# --------------------------------------------------------------------------- #
# Execute query
# --------------------------------------------------------------------------- #
if "last_df" not in st.session_state:
    st.session_state["last_df"] = None

if run_clicked or st.session_state["last_df"] is None:
    with st.expander("🔍 View generated SQL", expanded=False):
        st.code(query, language="sql")
        display_params = {k: str(v) for k, v in params.items()}
        st.json(display_params)

    try:
        with st.spinner("Running query..."):
            df = run_query(engine, query, params)
    except SQLAlchemyError as exc:
        st.error(f"❌ Query failed: {exc}")
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"❌ Unexpected error while running the query: {exc}")
        st.stop()
    st.session_state["last_df"] = df
else:
    df = st.session_state["last_df"]

if df is None:
    st.info("👈 Configure filters (optional) and click **Run / Refresh Dashboard**.")
    st.stop()

if len(df) == row_limit:
    st.warning(
        f"⚠️ Result was truncated at the {row_limit:,}-row safety cap. "
        "Narrow your filters or raise the cap in the sidebar for a complete dataset."
    )

st.success(f"Loaded **{len(df):,}** row(s) × **{len(df.columns)}** column(s) from `{selected_table}`.")

if df.empty:
    st.warning("No rows match the current filters. Adjust your filters and re-run.")
    st.stop()

numeric_cols_all = df.select_dtypes(include="number").columns.tolist()
all_cols = df.columns.tolist()

# --------------------------------------------------------------------------- #
# Tabs
# --------------------------------------------------------------------------- #
tab_raw, tab_pivot, tab_charts = st.tabs(
    ["📋 Filtered Raw Data", "📊 Interactive Pivot Table", "📈 Charts & Analytics"]
)

# --------------------------------------------------------------------------- #
# TAB 1 — Filtered Raw Data
# --------------------------------------------------------------------------- #
with tab_raw:
    display_cols = st.multiselect(
        "Columns to display", options=all_cols, default=all_cols, key="display_cols"
    )
    view_df = df[display_cols] if display_cols else df

    numeric_display_cols = [c for c in display_cols if c in numeric_cols_all]
    if numeric_display_cols:
        st.subheader("📈 KPI Summary")
        kpi_col_choice = st.multiselect(
            "Show KPI cards for", options=numeric_display_cols, default=numeric_display_cols[:3]
        )
        for col in kpi_col_choice:
            series = df[col].dropna()
            c1, c2, c3, c4 = st.columns(4)
            c1.metric(f"Sum — {col}", f"{series.sum():,.2f}" if len(series) else "0.00")
            c2.metric(f"Average — {col}", f"{series.mean():,.2f}" if len(series) else "0.00")
            c3.metric(f"Min — {col}", f"{series.min():,.2f}" if len(series) else "—")
            c4.metric(f"Max — {col}", f"{series.max():,.2f}" if len(series) else "—")
        st.divider()

    st.subheader("📋 Results")
    st.dataframe(view_df, use_container_width=True, height=450)
    st.caption(f"Total rows: **{len(view_df):,}** | Columns shown: {', '.join(display_cols) if display_cols else '—'}")

    csv_buffer = io.StringIO()
    view_df.to_csv(csv_buffer, index=False)
    st.download_button(
        label="⬇️ Download CSV",
        data=csv_buffer.getvalue(),
        file_name=f"{selected_table}_filtered.csv",
        mime="text/csv",
        use_container_width=True,
        key="download_raw",
    )

# --------------------------------------------------------------------------- #
# TAB 2 — Interactive Pivot Table
# --------------------------------------------------------------------------- #
with tab_pivot:
    st.subheader("🧮 Build a Pivot Table")
    p1, p2, p3, p4 = st.columns(4)
    with p1:
        pivot_rows = st.multiselect("Rows (Index)", options=all_cols, key="pivot_rows")
    with p2:
        pivot_col_choice = st.selectbox("Columns (optional)", options=["(None)"] + all_cols, key="pivot_col")
    with p3:
        pivot_values = st.multiselect("Values (numeric)", options=numeric_cols_all, key="pivot_values")
    with p4:
        agg_label = st.selectbox("Aggregation", options=["Sum", "Mean", "Count", "Min", "Max"], key="pivot_agg")

    agg_map = {"Sum": "sum", "Mean": "mean", "Count": "count", "Min": "min", "Max": "max"}

    if not pivot_rows or not pivot_values:
        st.info("Select at least one **Row** and one **Value** column to build a pivot table.")
    else:
        pivot_cols_param = None if pivot_col_choice == "(None)" else pivot_col_choice
        if pivot_cols_param:
            n_unique = df[pivot_cols_param].nunique(dropna=True)
            if n_unique > 60:
                st.warning(
                    f"`{pivot_cols_param}` has {n_unique:,} distinct values — the pivot table "
                    "may be very wide. Consider a different Columns field."
                )
        try:
            pivot_df = pd.pivot_table(
                df,
                index=pivot_rows,
                columns=pivot_cols_param,
                values=pivot_values,
                aggfunc=agg_map[agg_label],
                fill_value=0,
                margins=True,
                margins_name="Total",
                dropna=False,
            )
            st.dataframe(pivot_df, use_container_width=True, height=480)

            pivot_csv_buffer = io.StringIO()
            pivot_df.to_csv(pivot_csv_buffer)
            st.download_button(
                label="⬇️ Download Pivot Table CSV",
                data=pivot_csv_buffer.getvalue(),
                file_name=f"{selected_table}_pivot.csv",
                mime="text/csv",
                use_container_width=True,
                key="download_pivot",
            )
        except Exception as exc:  # noqa: BLE001
            st.error(f"❌ Could not build pivot table: {exc}")

# --------------------------------------------------------------------------- #
# TAB 3 — Charts & Analytics
# --------------------------------------------------------------------------- #
with tab_charts:
    st.subheader("📈 Build a Chart")

    if not numeric_cols_all:
        st.info("No numeric columns are available in the current result set to chart.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            chart_type = st.selectbox(
                "Chart type", options=["Bar Chart", "Line Chart", "Area Chart", "Pie Chart", "Treemap"]
            )
        with c2:
            x_col = st.selectbox("X-axis (Dimension)", options=all_cols, key="chart_x")
        with c3:
            y_col = st.selectbox("Y-axis (Metric)", options=numeric_cols_all, key="chart_y")
        with c4:
            color_choice = st.selectbox("Color / Group By (optional)", options=["(None)"] + all_cols, key="chart_color")
        color_col = None if color_choice == "(None)" else color_choice

        res_scale = st.slider(
            "Export resolution scale (higher = sharper PNG download via the chart's camera icon)",
            min_value=1,
            max_value=4,
            value=3,
            help="Base render is 1600×900px; scale ×3 or ×4 produces a 4K-and-beyond PNG export.",
        )

        chart_df = df.dropna(subset=[x_col, y_col]) if x_col and y_col else df
        fig = None
        try:
            if chart_df.empty:
                st.warning("No non-null data available for the selected X/Y columns.")
            elif chart_type == "Bar Chart":
                fig = px.bar(chart_df, x=x_col, y=y_col, color=color_col, barmode="group", template="plotly_white")
            elif chart_type == "Line Chart":
                plot_df = chart_df.sort_values(by=x_col)
                fig = px.line(plot_df, x=x_col, y=y_col, color=color_col, markers=True, template="plotly_white")
            elif chart_type == "Area Chart":
                plot_df = chart_df.sort_values(by=x_col)
                fig = px.area(plot_df, x=x_col, y=y_col, color=color_col, template="plotly_white")
            elif chart_type == "Pie Chart":
                if (chart_df[y_col] < 0).any():
                    st.warning("Pie charts require non-negative values; negative values were dropped.")
                    chart_df = chart_df[chart_df[y_col] >= 0]
                fig = px.pie(chart_df, names=x_col, values=y_col, template="plotly_white")
            elif chart_type == "Treemap":
                path = [color_col, x_col] if color_col else [x_col]
                if (chart_df[y_col] < 0).any():
                    st.warning("Treemaps require non-negative values; negative values were dropped.")
                    chart_df = chart_df[chart_df[y_col] >= 0]
                fig = px.treemap(chart_df, path=path, values=y_col, template="plotly_white")
        except Exception as exc:  # noqa: BLE001
            st.error(f"❌ Could not build chart: {exc}")
            fig = None

        if fig is not None:
            fig.update_layout(
                width=1600,
                height=900,
                font=dict(size=14),
                title=f"{chart_type}: {y_col} by {x_col}" + (f" (colored by {color_col})" if color_col else ""),
                margin=dict(t=60, l=40, r=40, b=40),
                hovermode="closest",
            )
            config = {
                "displaylogo": False,
                "toImageButtonOptions": {
                    "format": "png",
                    "filename": f"{selected_table}_{chart_type.replace(' ', '_').lower()}",
                    "scale": res_scale,
                },
            }
            st.plotly_chart(fig, use_container_width=True, config=config)
            st.caption(
                "Use the camera icon in the chart toolbar to export a high-resolution PNG "
                f"(≈{1600 * res_scale}×{900 * res_scale}px at the selected scale)."
            )
