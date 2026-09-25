"""
api.py
======
FastAPI backend running ALONGSIDE the Streamlit app -- same Postgres
database, same `app_employees` auth table, same pivot/DPD/field-efficiency
math (via db_core.py, extracted verbatim from app.py) -- exposed as JSON/
XLSX REST endpoints for other tools to consume.

Run with:
    uvicorn api:app --host 0.0.0.0 --port 8000

Auth: JWT bearer tokens issued by POST /auth/login, checked against the
SAME `authorized_employees` / `app_users` / `login_audit` tables the
Streamlit app's login screen uses (bcrypt password hash, access_level,
is_authorized) -- an employee's Streamlit login works here too, and
vice versa.

Env vars: same DB_HOST/DB_PORT/DB_NAME/DB_USER/DB_PASSWORD/DB_SSLMODE or
DATABASE_URL as app.py, plus:
    JWT_SECRET        (required in production; a random default is used
                       otherwise, logged loudly at startup)
    JWT_EXPIRES_MIN    default 480 (8 hours)
    CORS_ORIGINS       comma-separated list, default "*"
"""

from __future__ import annotations

import os
import secrets as pysecrets
from datetime import datetime, timedelta, timezone
from typing import Any, Literal
from dotenv import load_dotenv
load_dotenv()
import jwt
import pandas as pd
from fastapi import Depends, FastAPI, HTTPException, Security
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from pydantic import BaseModel, Field
from sqlalchemy.exc import SQLAlchemyError

import db_core as core

# --------------------------------------------------------------------------- #
# App / CORS
# --------------------------------------------------------------------------- #
app = FastAPI(title="DB Explorer API", version="1.0.0")

_cors_origins = os.environ.get("CORS_ORIGINS", "*")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"] if _cors_origins.strip() == "*" else [o.strip() for o in _cors_origins.split(",")],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

JWT_SECRET = os.environ.get("JWT_SECRET")
if not JWT_SECRET:
    JWT_SECRET = pysecrets.token_urlsafe(32)
    print(
        "⚠️  JWT_SECRET not set -- using a random secret generated for this "
        "process only. Tokens will stop working on restart, and won't be "
        "valid across multiple worker processes. Set JWT_SECRET in "
        "production."
    )
JWT_ALGO = "HS256"
JWT_EXPIRES_MIN = int(os.environ.get("JWT_EXPIRES_MIN", "480"))

bearer_scheme = HTTPBearer(auto_error=False)


@app.on_event("startup")
def _startup() -> None:
    engine = core.get_engine()
    core.ensure_auth_tables(engine)


# --------------------------------------------------------------------------- #
# Auth
# --------------------------------------------------------------------------- #
class LoginRequest(BaseModel):
    employee_code: str
    password: str


class LoginResponse(BaseModel):
    token: str
    employee_code: str
    access_level: str
    expires_at: str


class CurrentUser(BaseModel):
    employee_code: str
    access_level: str


def _make_token(employee_code: str, access_level: str) -> tuple[str, datetime]:
    expires_at = datetime.now(timezone.utc) + timedelta(minutes=JWT_EXPIRES_MIN)
    payload = {"sub": employee_code, "access_level": access_level, "exp": expires_at}
    token = jwt.encode(payload, JWT_SECRET, algorithm=JWT_ALGO)
    return token, expires_at


def get_current_user(creds: HTTPAuthorizationCredentials | None = Security(bearer_scheme)) -> CurrentUser:
    if creds is None:
        raise HTTPException(401, "Missing bearer token")
    try:
        payload = jwt.decode(creds.credentials, JWT_SECRET, algorithms=[JWT_ALGO])
    except jwt.ExpiredSignatureError:
        raise HTTPException(401, "Token expired")
    except jwt.InvalidTokenError:
        raise HTTPException(401, "Invalid token")
    return CurrentUser(employee_code=payload["sub"], access_level=payload.get("access_level", "user"))


def require_access(*levels: str):
    def dep(user: CurrentUser = Depends(get_current_user)) -> CurrentUser:
        if user.access_level not in levels:
            raise HTTPException(403, f"Requires access level: {', '.join(levels)}")
        return user
    return dep


@app.post("/auth/login", response_model=LoginResponse)
def login(body: LoginRequest):
    """
    Checked against the SAME three tables app.py's login screen uses
    (`authorized_employees`, `app_users`, `login_audit`) -- an employee's
    Streamlit password works here too, and vice versa.
    """
    engine = core.get_engine()
    if core._too_many_recent_failures(engine, body.employee_code):
        raise HTTPException(429, "Too many failed login attempts. Please try again in a few minutes.")

    # (is_authorized, employee_name, access_level)
    authorized, _employee_name, access_level = core._is_employee_authorized(engine, body.employee_code)
    password_hash = core._get_password_hash(engine, body.employee_code)
    if not authorized or not password_hash or not core._verify_password(body.password, password_hash):
        core._log_audit_event(engine, body.employee_code, "login_failure")
        raise HTTPException(401, "Invalid employee code or password.")

    core._touch_last_login(engine, body.employee_code)
    core._log_audit_event(engine, body.employee_code, "login_success")
    token, expires_at = _make_token(body.employee_code, access_level or "user")
    return LoginResponse(
        token=token, employee_code=body.employee_code,
        access_level=access_level or "user", expires_at=expires_at.isoformat(),
    )


@app.get("/auth/me", response_model=CurrentUser)
def me(user: CurrentUser = Depends(get_current_user)):
    return user


# --------------------------------------------------------------------------- #
# Error handling helper
# --------------------------------------------------------------------------- #
def _run_safely(fn, *args, **kwargs):
    try:
        return fn(*args, **kwargs)
    except SQLAlchemyError as exc:
        raise HTTPException(500, f"Database error: {exc}")
    except ValueError as exc:
        raise HTTPException(400, str(exc))


# --------------------------------------------------------------------------- #
# Schema
# --------------------------------------------------------------------------- #
@app.get("/tables", response_model=list[str])
def list_tables(_: CurrentUser = Depends(get_current_user)):
    engine = core.get_engine()
    return _run_safely(core.list_tables, engine)


class ColumnMeta(BaseModel):
    name: str
    data_type: str
    kind: str  # "date" | "numeric" | "text"


@app.get("/tables/{table_name}/columns", response_model=list[ColumnMeta])
def get_columns(table_name: str, _: CurrentUser = Depends(get_current_user)):
    engine = core.get_engine()
    cols = _run_safely(core.get_columns, engine, table_name)
    return [
        ColumnMeta(name=c["name"], data_type=str(c["type"]), kind=core.classify_column(c["type"]))
        for c in cols
    ]


@app.get("/tables/{table_name}/distinct/{column_name}", response_model=list[Any])
def distinct_values(table_name: str, column_name: str, limit: int = 1000, _: CurrentUser = Depends(get_current_user)):
    engine = core.get_engine()
    return _run_safely(core.get_distinct_values, engine, table_name, column_name, limit)


@app.get("/tables/{table_name}/year-months/{column_name}", response_model=list[list[int]])
def year_months(table_name: str, column_name: str, _: CurrentUser = Depends(get_current_user)):
    engine = core.get_engine()
    pairs = _run_safely(core.get_available_year_months, engine, table_name, column_name)
    return [[y, m] for y, m in pairs]


def _column_kind_map(table_name: str, cols: list[str]) -> dict[str, str]:
    engine = core.get_engine()
    all_cols = core.get_columns(engine, table_name)
    kind_by_name = {c["name"]: core.classify_column(c["type"]) for c in all_cols}
    missing = [c for c in cols if c not in kind_by_name]
    if missing:
        raise HTTPException(400, f"Unknown column(s): {', '.join(missing)}")
    return kind_by_name


# --------------------------------------------------------------------------- #
# Filtered report (flat table + KPIs)
# --------------------------------------------------------------------------- #
class ReportRequest(BaseModel):
    ref_column: str
    output_columns: list[str] = Field(default_factory=list)
    ref_kind: Literal["date", "numeric", "text"]
    filter_payload: dict[str, Any] = Field(default_factory=dict)
    limit: int = 5000


@app.post("/tables/{table_name}/report")
def run_report(table_name: str, body: ReportRequest, _: CurrentUser = Depends(get_current_user)):
    engine = core.get_engine()
    query, params = _run_safely(
        core.build_query, table_name, body.ref_column, body.output_columns, body.ref_kind, body.filter_payload,
    )
    query = f"{query} LIMIT {int(body.limit)}"
    df = _run_safely(core.run_query, engine, query, params)
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    kpis = {
        "record_count": int(len(df)),
        **{
            c: {
                "sum": float(df[c].sum()) if len(df) else 0.0,
                "avg": float(df[c].mean()) if len(df) else 0.0,
                "min": float(df[c].min()) if len(df) else 0.0,
                "max": float(df[c].max()) if len(df) else 0.0,
            }
            for c in numeric_cols
        },
    }
    return {"columns": list(df.columns), "rows": df.to_dict(orient="records"), "kpis": kpis}


# --------------------------------------------------------------------------- #
# Pivot: Summary Matrix (Excel-style, with subtotal trees)
# --------------------------------------------------------------------------- #
class PivotRequest(BaseModel):
    row_cols: list[str]
    col_cols: list[str] = Field(default_factory=list)
    value_col: str
    agg_func: Literal["sum", "mean", "count", "min", "max"] = "sum"
    in_crores: bool = False
    filters: dict[str, list[Any]] = Field(default_factory=dict)  # {column: [allowed values]}
    dpd_slab_field: str | None = None  # if set, adds cumulative 1+/30+/90+ buckets
    dpd_only_show_buckets: bool = True


def _fetch_pivot_source(table_name: str, body: PivotRequest) -> tuple[pd.DataFrame, dict[str, str]]:
    engine = core.get_engine()
    column_kind_map = _column_kind_map(table_name, [body.value_col, *body.row_cols, *body.col_cols, *body.filters.keys()])
    filter_conditions = tuple((col, tuple(vals)) for col, vals in body.filters.items() if vals)
    source_df = _run_safely(
        core.fetch_pivot_source_data,
        engine, table_name, tuple(body.row_cols), tuple(body.col_cols), body.value_col,
        filter_conditions, column_kind_map,
    )
    return source_df, column_kind_map


def _df_to_matrix_json(display_df: pd.DataFrame, group_row_flags: list[bool], total_col_flags: list[bool]) -> dict:
    has_multi_cols = isinstance(display_df.columns, pd.MultiIndex)
    col_tuples = [list(t) for t in display_df.columns] if has_multi_cols else [[c] for c in display_df.columns]
    row_labels = [str(v) for v in display_df.index]
    values = display_df.to_numpy(dtype="float64", na_value=None)
    rows = [
        {"label": row_labels[i], "is_group": bool(group_row_flags[i]) if i < len(group_row_flags) else False,
         "values": [None if pd.isna(v) else float(v) for v in values[i]]}
        for i in range(len(display_df))
    ]
    return {
        "row_axis_name": display_df.index.name or "",
        "columns": [{"tuple": t, "is_total": bool(total_col_flags[j]) if j < len(total_col_flags) else False} for j, t in enumerate(col_tuples)],
        "rows": rows,
    }


@app.post("/tables/{table_name}/pivot/matrix")
def pivot_matrix(table_name: str, body: PivotRequest, _: CurrentUser = Depends(get_current_user)):
    if not body.row_cols:
        raise HTTPException(400, "At least one Rows field is required.")
    source_df, column_kind_map = _fetch_pivot_source(table_name, body)
    display_df, group_row_flags, total_col_flags = _run_safely(
        core.build_excel_style_pivot,
        source_df, body.row_cols, body.col_cols, body.value_col, body.agg_func, column_kind_map, body.in_crores,
    )

    dpd_warning = None
    if body.dpd_slab_field:
        try:
            slab_level = core._effective_col_level(body.col_cols, body.dpd_slab_field, column_kind_map)
            display_df = core.add_dpd_buckets_to_excel_pivot(display_df, slab_level, only_show_buckets=body.dpd_only_show_buckets)
            if isinstance(display_df.columns, pd.MultiIndex):
                total_col_flags = [str(t[-1]).strip().endswith("Total") for t in display_df.columns]
            else:
                total_col_flags = [str(c).strip().endswith("Total") for c in display_df.columns]
        except ValueError as exc:
            dpd_warning = str(exc)

    result = _df_to_matrix_json(display_df, group_row_flags, total_col_flags)
    if dpd_warning:
        result["dpd_warning"] = dpd_warning
    return result


@app.post("/tables/{table_name}/pivot/matrix.xlsx")
def pivot_matrix_xlsx(table_name: str, body: PivotRequest, _: CurrentUser = Depends(get_current_user)):
    if not body.row_cols:
        raise HTTPException(400, "At least one Rows field is required.")
    source_df, column_kind_map = _fetch_pivot_source(table_name, body)
    display_df, group_row_flags, total_col_flags = _run_safely(
        core.build_excel_style_pivot,
        source_df, body.row_cols, body.col_cols, body.value_col, body.agg_func, column_kind_map, body.in_crores,
    )
    if body.dpd_slab_field:
        try:
            slab_level = core._effective_col_level(body.col_cols, body.dpd_slab_field, column_kind_map)
            display_df = core.add_dpd_buckets_to_excel_pivot(display_df, slab_level, only_show_buckets=body.dpd_only_show_buckets)
            if isinstance(display_df.columns, pd.MultiIndex):
                total_col_flags = [str(t[-1]).strip().endswith("Total") for t in display_df.columns]
            else:
                total_col_flags = [str(c).strip().endswith("Total") for c in display_df.columns]
        except ValueError:
            pass  # export without buckets if they don't apply

    value_kind = "count" if body.agg_func == "count" else "number"
    xlsx_bytes = core.dataframe_to_formatted_excel_bytes(
        display_df, group_row_flags, total_col_flags, value_kind=value_kind,
        table_style="matrix", sheet_name="Summary Matrix",
    )
    return StreamingResponse(
        iter([xlsx_bytes]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{table_name}_pivot.xlsx"'},
    )


# --------------------------------------------------------------------------- #
# Pivot: Field Efficiency (% of a chosen field's group total)
# --------------------------------------------------------------------------- #
class EfficiencyRequest(PivotRequest):
    normalize_field: str
    isolate_status: str | None = None


@app.post("/tables/{table_name}/pivot/efficiency")
def pivot_efficiency(table_name: str, body: EfficiencyRequest, _: CurrentUser = Depends(get_current_user)):
    if not body.row_cols:
        raise HTTPException(400, "At least one Rows field is required.")
    source_df, column_kind_map = _fetch_pivot_source(table_name, body)
    display_df, group_row_flags = _run_safely(
        core.build_field_efficiency_pivot,
        source_df, body.row_cols, body.col_cols, body.normalize_field, body.value_col, body.agg_func, column_kind_map,
    )
    total_col_flags = [False] * (display_df.shape[1] - 1) + [True]

    if body.isolate_status:
        display_df = _run_safely(core.slice_field_efficiency_to_status, display_df, body.normalize_field, body.isolate_status)
        total_col_flags = [False] * display_df.shape[1]

    return _df_to_matrix_json(display_df, group_row_flags, total_col_flags)


@app.post("/tables/{table_name}/pivot/efficiency.xlsx")
def pivot_efficiency_xlsx(table_name: str, body: EfficiencyRequest, _: CurrentUser = Depends(get_current_user)):
    if not body.row_cols:
        raise HTTPException(400, "At least one Rows field is required.")
    source_df, column_kind_map = _fetch_pivot_source(table_name, body)
    display_df, group_row_flags = _run_safely(
        core.build_field_efficiency_pivot,
        source_df, body.row_cols, body.col_cols, body.normalize_field, body.value_col, body.agg_func, column_kind_map,
    )
    total_col_flags = [False] * (display_df.shape[1] - 1) + [True]
    if body.isolate_status:
        display_df = _run_safely(core.slice_field_efficiency_to_status, display_df, body.normalize_field, body.isolate_status)
        total_col_flags = [False] * display_df.shape[1]

    xlsx_bytes = core.dataframe_to_formatted_excel_bytes(
        display_df, group_row_flags, total_col_flags, value_kind="percent",
        table_style="efficiency", sheet_name="Field Efficiency",
    )
    return StreamingResponse(
        iter([xlsx_bytes]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{table_name}_pivot_efficiency.xlsx"'},
    )


# --------------------------------------------------------------------------- #
# Pivot: flat (no subtotals) -- handy for charting / other tools
# --------------------------------------------------------------------------- #
@app.post("/tables/{table_name}/pivot/flat")
def pivot_flat(table_name: str, body: PivotRequest, _: CurrentUser = Depends(get_current_user)):
    if not body.row_cols:
        raise HTTPException(400, "At least one Rows field is required.")
    source_df, _kind_map = _fetch_pivot_source(table_name, body)
    flat_df = _run_safely(core.build_pivot_table, source_df, body.row_cols, body.col_cols, body.agg_func)
    has_multi_cols = isinstance(flat_df.columns, pd.MultiIndex)
    columns = [list(t) for t in flat_df.columns] if has_multi_cols else [[c] for c in flat_df.columns]
    values = flat_df.to_numpy(dtype="float64", na_value=None)
    rows = [
        {"label": str(idx), "values": [None if pd.isna(v) else float(v) for v in values[i]]}
        for i, idx in enumerate(flat_df.index)
    ]
    return {"columns": columns, "rows": rows}
