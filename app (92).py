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
) -> None:
    """
    Render a Plotly bar chart with the value of each bar printed on top of
    it, styled for high-resolution export. The base canvas is 1600x800px;
    at res_scale=3 the PNG export is 4800x2400 and at res_scale=4 it's
    6400x3200 — both comfortably exceed 4K (3840x2160).
    """
    fig = px.bar(
        df,
        x=x_col,
        y=y_col,
        color=color_col,
        barmode="group" if color_col else "relative",
        template="plotly_white",
        text=y_col,
        labels={x_col: x_col.replace("_", " ").title(), y_col: "Count"},
        title=title,
    )
    fig.update_traces(
        texttemplate="%{text:,}",
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
            df, x_col, y_col, color_col, title, filename_prefix, key_suffix, res_scale
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
                "Breakdown by (optional) — e.g. payment_mode",
                options=breakdown_candidates,
                key=f"breakdown_choice_{selected_table}",
            )
            breakdown_col = None if breakdown_choice == "(None)" else breakdown_choice

            numeric_cols_for_sum = [c for c in column_names if column_kind_map.get(c) == "numeric"]
            if numeric_cols_for_sum:
                want_sum_choice = st.radio(
                    "➕ Show the SUM of a column for each month/year? (e.g. total `emi` for March)",
                    options=["No", "Yes"],
                    index=0,
                    key=f"want_sum_{selected_table}",
                    horizontal=True,
                )
                if want_sum_choice == "Yes":
                    sum_columns = st.multiselect(
                        "Which numeric column(s) should be summed?",
                        options=numeric_cols_for_sum,
                        key=f"sum_columns_{selected_table}",
                        help=(
                            "Computes SUM(column) for every month and year "
                            "(further split by the breakdown column above, if you set one)."
                        ),
                    )
                    if sum_columns:
                        crores_choice = st.radio(
                            "💰 Display these sums in Crores (÷ 1,00,00,000)?",
                            options=["No", "Yes"],
                            index=0,
                            key=f"crores_choice_{selected_table}",
                            horizontal=True,
                        )
                        show_in_crores = crores_choice == "Yes"
            else:
                st.caption("ℹ️ No numeric columns found in this table to sum.")

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
    run_clicked = st.button("🚀 Generate Report", type="primary", use_container_width=True)

# --------------------------------------------------------------------------- #
# Main UI - Tabs Layout
# --------------------------------------------------------------------------- #
tab_report, tab_sql = st.tabs(["📊 Dynamic Report Builder", "🧑‍💻 Custom SQL Query"])

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

                    if breakdown_col:
                        render_split_bar_charts(
                            display_df,
                            x_col="month_label",
                            y_col=sum_field,
                            color_col=breakdown_col,
                            title=f"Sum of {sc} per month by {breakdown_col}{unit_note}",
                            filename_prefix=f"{selected_table}_monthly_sum_{sc}_{breakdown_col}",
                            key_suffix=f"monthly_sum_{sc}_breakdown",
                            res_scale=chart_res_scale,
                            num_splits=num_chart_splits,
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

                if breakdown_col:
                    render_split_bar_charts(
                        yearly_display_df,
                        x_col="year_label",
                        y_col=sum_field,
                        color_col=breakdown_col,
                        title=f"Sum of {sc} per year by {breakdown_col}{unit_note}",
                        filename_prefix=f"{selected_table}_yearly_sum_{sc}_{breakdown_col}",
                        key_suffix=f"yearly_sum_{sc}_breakdown",
                        res_scale=chart_res_scale,
                        num_splits=num_chart_splits,
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