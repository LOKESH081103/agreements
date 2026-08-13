# 📊 Dynamic Database Explorer & Report Generator

A Streamlit app that connects to any PostgreSQL database, lets you pick a
table, filter it by a reference (date or categorical) column, choose output
columns, and view/download the resulting report — with automatic KPI cards
for numeric columns.

## 1. Install dependencies

```bash
python3 -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## 2. Configure the database connection

The app reads config from **`st.secrets`** first, then falls back to
**environment variables**. Use whichever fits your workflow — you don't need
both.

### Option A — Environment variables (local dev)

```bash
cp .env.example .env
# edit .env with your real DB_HOST / DB_PORT / DB_NAME / DB_USER / DB_PASSWORD
export $(grep -v '^#' .env | xargs)
streamlit run app.py
```

Or export manually:

```bash
export DB_HOST=localhost
export DB_PORT=5432
export DB_NAME=mydb
export DB_USER=myuser
export DB_PASSWORD=mypassword
export DB_SSLMODE=prefer          # use "require" for most managed/cloud DBs
streamlit run app.py
```

### Option B — `secrets.toml` (Streamlit Community Cloud or local)

```bash
mkdir -p .streamlit
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# edit .streamlit/secrets.toml with your real credentials
streamlit run app.py
```

**Never commit `.env` or `.streamlit/secrets.toml` to version control.** Both
example files are safe templates only — add the real files to `.gitignore`.

### Option C — Single connection string

Instead of the individual `DB_*` variables, you can set one `DATABASE_URL`
(as an env var or in secrets.toml) and it will take priority:

```
DATABASE_URL=postgresql+psycopg2://user:password@host:5432/dbname?sslmode=require
```

## 3. Run the app

```bash
streamlit run app.py
```

Open the URL Streamlit prints (typically `http://localhost:8501`).

## How it works

1. **Sidebar → Select a table**: the app inspects the `public` schema via
   `sqlalchemy.inspect()` and lists every table.
2. **Reference (Filter) Column**: pick any column; the app auto-detects
   whether it's a date/timestamp or categorical/text type (you can override
   this). For dates you choose a granularity — Specific Date(s), Month/Year,
   or Year(s) — and the app builds a query using `EXTRACT(...)` / date
   equality filters. For categorical columns you get a multi-select of
   distinct values.
3. **Output Columns**: pick one or more columns to retrieve. Any numeric
   output column automatically gets `st.metric` KPI cards (Sum, Average,
   Count).
4. **Generate Report**: builds and runs a fully parameterized SQL query (no
   string-interpolated values — only whitelisted, quoted identifiers plus
   `:named` bind parameters), shows the SQL used (expandable), displays the
   results in a searchable/sortable `st.dataframe`, and offers a CSV
   download.

## Security notes

- All filter **values** are passed as bound parameters (`:param`), never
  concatenated into the SQL string — this prevents SQL injection from filter
  values.
- Table and column **names** come only from `sqlalchemy.inspect()` (i.e.
  they're read from the database's own catalog, not typed by the user as
  free text), and are double-quote-escaped when interpolated into the
  `SELECT`/`WHERE` clause, which is the standard, safe way to reference
  dynamic identifiers in PostgreSQL.
- The DB engine is created once via `@st.cache_resource` and reused across
  reruns/sessions (connection pooling via SQLAlchemy).
- Connection and query failures are caught and shown via `st.error` with a
  clean message instead of a raw traceback.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| "Could not connect to the database" | Wrong host/port/credentials, or DB not reachable from this machine (check firewall / `pg_hba.conf`) |
| SSL error | Set `DB_SSLMODE=require` for most managed Postgres (RDS, Supabase, Neon, etc.) |
| No tables listed | Your DB user may lack privileges on the `public` schema, or all tables live in a different schema |
| Query works in `psql` but fails here | Check for reserved-word column names — the app already quotes identifiers, but exotic types (arrays, JSONB) may need custom handling |
