# FastAPI backend (alongside Streamlit)

Three files, run alongside your existing `app.py`, same database, same
login credentials:

- **`db_core.py`** — the pure DB/pivot engine (schema introspection, query
  building, the Excel-style pivot / DPD-bucket / Field-Efficiency math,
  Excel export, auth helpers). Extracted **verbatim** from `app.py`'s
  non-Streamlit code, so both apps compute identical numbers from one
  source of truth — no separate implementation to drift out of sync.
- **`api.py`** — the FastAPI app: JWT auth + REST/JSON endpoints wrapping
  `db_core.py`.
- **`requirements-api.txt`** — dependencies for this service specifically.

`app.py` itself is **unchanged** — it keeps its own copies of these
functions and keeps working exactly as before. If you'd rather have zero
duplication (both apps importing the same `db_core.py`), say the word and
I'll do that refactor as a follow-up; it touches `app.py`, which this pass
deliberately didn't.

## Install & run

```bash
pip install -r requirements-api.txt

export DATABASE_URL=postgresql://user:pass@host:5432/dbname   # same as app.py
export JWT_SECRET=$(python -c "import secrets; print(secrets.token_urlsafe(32))")
export CORS_ORIGINS="*"          # or a comma-separated allow-list

uvicorn api:app --host 0.0.0.0 --port 8000 --reload
```

Interactive docs at `http://localhost:8000/docs` (FastAPI's auto-generated
Swagger UI) — the fastest way to try every endpoint by hand.

## Auth

```bash
curl -X POST localhost:8000/auth/login \
  -H "Content-Type: application/json" \
  -d '{"employee_code": "E123", "password": "..."}'
# -> {"token": "...", "employee_code": "E123", "access_level": "analyst", ...}
```

Same `authorized_employees` / `app_users` / `login_audit` tables `app.py`
already uses — an employee's Streamlit password works here too, and every
API login is written to the same `login_audit` table Streamlit's admin
panel already reads.

Use the token as a bearer header on everything else:

```bash
curl localhost:8000/tables -H "Authorization: Bearer <token>"
```

## Endpoints

| Method | Path | Purpose |
|---|---|---|
| POST | `/auth/login` | Get a JWT |
| GET | `/auth/me` | Whoami |
| GET | `/tables` | List tables |
| GET | `/tables/{table}/columns` | Column names + kind (date/numeric/text) |
| GET | `/tables/{table}/distinct/{column}` | Distinct values for a filter dropdown |
| GET | `/tables/{table}/year-months/{column}` | Available (year, month) pairs for a date column |
| POST | `/tables/{table}/report` | Filtered flat report + auto KPIs |
| POST | `/tables/{table}/pivot/matrix` | Summary Matrix (subtotal tree), JSON |
| POST | `/tables/{table}/pivot/matrix.xlsx` | Same, as a formatted .xlsx download |
| POST | `/tables/{table}/pivot/efficiency` | Field Efficiency Table (% of group), JSON |
| POST | `/tables/{table}/pivot/efficiency.xlsx` | Same, as .xlsx |
| POST | `/tables/{table}/pivot/flat` | Flat pivot, no subtotals (for charting) |

`pivot/matrix` and `pivot/efficiency` both accept an optional
`dpd_slab_field` to add cumulative 1+/30+/90+ DPD buckets, same rules as
the Streamlit checkbox (needs one Columns field holding raw slab labels).

Example:

```bash
curl -X POST localhost:8000/tables/loans/pivot/matrix \
  -H "Authorization: Bearer <token>" -H "Content-Type: application/json" \
  -d '{
        "row_cols": ["zone", "sub_zone"],
        "col_cols": ["staging_ecl"],
        "value_col": "outstanding_amount",
        "agg_func": "sum",
        "in_crores": true,
        "filters": {"region": ["EAST", "WEST"]},
        "dpd_slab_field": "staging_ecl"
      }'
```

## Scope / what's not ported

- The **multi-Values-column side-by-side** composition
  (`render_multi_value_pivot_table` / `render_multi_value_field_efficiency_table`
  in `app.py`) isn't in the API yet — every pivot/efficiency endpoint here
  takes one `value_col`. The single-value engine underneath (subtotal
  trees, DPD buckets, field efficiency, Excel export) is fully ported and
  identical to `app.py`.
- **Flow Retention** (the Slab/Flow/Normalised/Stabilized-specific table)
  isn't ported — it's a thin derived view on top of the Summary Matrix
  that's easy to add if you need it via the API too.
- The **Custom SQL** tab and **file-upload-to-table** tab weren't ported —
  intentionally: exposing either over a public API is a materially
  different security decision than a logged-in Streamlit session, worth
  a deliberate choice rather than a default "yes."
- `ttl_cache` in `db_core.py` is a simple in-process dict — fine for one
  worker; if you run `uvicorn` with multiple workers, each has its own
  cache (harmless, just less cache-hit benefit) — swap in Redis if that
  matters to you.
