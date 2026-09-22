"""
reconcile.py
============
Reconciles loan-agreement party data extracted from PDFs (`extracted_notices.xlsx`)
against a wide-format company master database (`company_data.xlsx`).

USAGE
-----
    python reconcile.py

By default it looks for:
    extracted_notices.xlsx   (long format: one row per party)
    company_data.xlsx        (wide format: one row per agreement)
in the current directory, and writes:
    reconciliation_report.xlsx

Edit the CONFIG block below if your file names / sheet names differ.

DEPENDENCIES
------------
    pip install pandas openpyxl rapidfuzz
(If rapidfuzz isn't available the script automatically falls back to
Python's built-in difflib, so it will still run - just a bit slower.)
"""

import re
import sys
import pandas as pd

# --------------------------------------------------------------------------
# CONFIG - adjust paths / thresholds here
# --------------------------------------------------------------------------
EXTRACTED_FILE = "extracted_notices.xlsx"
COMPANY_FILE = "company_data.xlsx"
OUTPUT_FILE = "reconciliation_report.xlsx"

# Similarity thresholds (0-100). Tune to taste.
MATCH_THRESHOLD = 90          # >= this on BOTH name & address  -> MATCH
PARTIAL_THRESHOLD = 60        # >= this on BOTH (but not MATCH) -> PARTIAL MATCH
# anything below PARTIAL_THRESHOLD, or with no company row      -> MISMATCH / NOT FOUND

# --------------------------------------------------------------------------
# Fuzzy matching backend (rapidfuzz preferred, difflib fallback)
# --------------------------------------------------------------------------
try:
    from rapidfuzz import fuzz

    def similarity(a: str, b: str) -> float:
        """0-100 similarity score, robust to word order / partial overlap."""
        if not a and not b:
            return 100.0
        if not a or not b:
            return 0.0
        # token_sort handles word-order differences; token_set handles
        # "full name contains shortened name" / extra tokens gracefully.
        return max(fuzz.token_sort_ratio(a, b), fuzz.token_set_ratio(a, b))

except ImportError:  # pragma: no cover - fallback path
    from difflib import SequenceMatcher

    def similarity(a: str, b: str) -> float:
        if not a and not b:
            return 100.0
        if not a or not b:
            return 0.0
        a_tokens = set(a.split())
        b_tokens = set(b.split())
        token_set_score = 0.0
        if a_tokens and b_tokens:
            common = a_tokens & b_tokens
            token_set_score = (
                100.0 * 2 * len(common) / (len(a_tokens) + len(b_tokens))
            )
        seq_score = 100.0 * SequenceMatcher(None, a, b).ratio()
        return max(seq_score, token_set_score)


# --------------------------------------------------------------------------
# Text normalization
# --------------------------------------------------------------------------
def normalize(text) -> str:
    """Uppercase, strip punctuation/extra whitespace, for fair comparison."""
    if pd.isna(text):
        return ""
    text = str(text).upper()
    text = re.sub(r"[^A-Z0-9\s]", " ", text)   # drop punctuation
    text = re.sub(r"\s+", " ", text).strip()   # collapse whitespace
    return text


def normalize_agreement_no(text) -> str:
    if pd.isna(text):
        return ""
    return str(text).strip().upper()


# --------------------------------------------------------------------------
# Step 1: Load extracted notices (already long format)
# --------------------------------------------------------------------------
def load_extracted(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)

    required = ["Agreement No", "Party Number", "Name", "Address"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"'{path}' is missing expected column(s): {missing}")

    df["Agreement No"] = df["Agreement No"].apply(normalize_agreement_no)
    df["Party Number"] = pd.to_numeric(df["Party Number"], errors="coerce").astype("Int64")

    if "Cheque No" not in df.columns:
        df["Cheque No"] = ""
    if "Source File" not in df.columns:
        df["Source File"] = ""

    df["_name_norm"] = df["Name"].apply(normalize)
    df["_addr_norm"] = df["Address"].apply(normalize)

    return df[
        [
            "Agreement No",
            "Party Number",
            "Name",
            "Address",
            "Cheque No",
            "Source File",
            "_name_norm",
            "_addr_norm",
        ]
    ]


# --------------------------------------------------------------------------
# Step 2: Load company master data and unpivot wide -> long
# --------------------------------------------------------------------------
def find_agreement_col(columns) -> str:
    candidates = ["Agreement No", "AGREEMENTNO", "Loan Account No", "LOAN ACCOUNT NO"]
    for c in candidates:
        if c in columns:
            return c
    # fallback: fuzzy search for something containing "agreement" or "loan" + "account"
    for c in columns:
        cu = str(c).upper()
        if "AGREEMENT" in cu or ("LOAN" in cu and "ACCOUNT" in cu):
            return c
    raise ValueError(
        "Could not find an Agreement/Loan-Account-No column in company data. "
        f"Available columns: {list(columns)}"
    )


def find_borrower_cols(columns):
    """
    Returns (name_col, addr_col) for the primary borrower, trying a few
    common naming conventions.
    """
    name_candidates = ["Borrower Name", "Borrower", "BORROWER NAME", "BORROWER"]
    addr_candidates = ["Borrower Add", "Borrower Address", "BORROWER ADD", "BORROWER ADDRESS"]

    name_col = next((c for c in name_candidates if c in columns), None)
    addr_col = next((c for c in addr_candidates if c in columns), None)

    if name_col is None or addr_col is None:
        raise ValueError(
            "Could not find primary Borrower name/address columns in company data. "
            f"Available columns: {list(columns)}"
        )
    return name_col, addr_col


def find_coborrower_cols(columns):
    """
    Finds every Co-Borrower N Name / Co-Borrower N Address pair, regardless
    of exact separator style (e.g. 'Co-Borrower-1', 'Co-Borrower 1 Name',
    'Co-Borrower1Name').
    Returns a dict: {N: {"name": col, "add": col}}
    """
    pairs = {}
    name_pattern = re.compile(r"CO[\s\-_]*BORROWER[\s\-_]*(\d+)(?:[\s\-_]*(NAME))?$", re.I)
    addr_pattern = re.compile(r"CO[\s\-_]*BORROWER[\s\-_]*(\d+)[\s\-_]*ADD(?:RESS)?$", re.I)

    for col in columns:
        col_clean = str(col).strip()
        m_addr = addr_pattern.match(col_clean.upper())
        if m_addr:
            n = int(m_addr.group(1))
            pairs.setdefault(n, {})["add"] = col
            continue
        m_name = name_pattern.match(col_clean.upper())
        if m_name:
            n = int(m_name.group(1))
            pairs.setdefault(n, {})["name"] = col

    # keep only complete pairs (both name & add found)
    return {n: v for n, v in pairs.items() if "name" in v and "add" in v}


def load_company_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)

    agreement_col = find_agreement_col(df.columns)
    borrower_name_col, borrower_addr_col = find_borrower_cols(df.columns)
    coborrower_cols = find_coborrower_cols(df.columns)

    records = []
    for _, row in df.iterrows():
        agreement_no = normalize_agreement_no(row[agreement_col])
        if not agreement_no:
            continue

        # Party Number 1 = primary borrower
        records.append(
            {
                "Agreement No": agreement_no,
                "Party Number": 1,
                "Company Name": row[borrower_name_col],
                "Company Address": row[borrower_addr_col],
            }
        )

        # Party Number 2, 3, ... = Co-Borrower 1, 2, ...
        for n in sorted(coborrower_cols):
            name_val = row[coborrower_cols[n]["name"]]
            addr_val = row[coborrower_cols[n]["add"]]
            if pd.isna(name_val) or str(name_val).strip() == "":
                continue  # skip empty co-borrower slots
            records.append(
                {
                    "Agreement No": agreement_no,
                    "Party Number": n + 1,
                    "Company Name": name_val,
                    "Company Address": addr_val,
                }
            )

    long_df = pd.DataFrame(records)
    long_df["Party Number"] = long_df["Party Number"].astype("Int64")
    long_df["_name_norm"] = long_df["Company Name"].apply(normalize)
    long_df["_addr_norm"] = long_df["Company Address"].apply(normalize)
    return long_df


# --------------------------------------------------------------------------
# Step 3: Merge + fuzzy compare
# --------------------------------------------------------------------------
def classify(name_sim: float, addr_sim: float, has_company_row: bool) -> str:
    if not has_company_row:
        return "NOT FOUND IN COMPANY DATA"
    if name_sim >= MATCH_THRESHOLD and addr_sim >= MATCH_THRESHOLD:
        return "MATCH"
    if name_sim >= PARTIAL_THRESHOLD and addr_sim >= PARTIAL_THRESHOLD:
        return "PARTIAL MATCH"
    return "MISMATCH"


def reconcile(extracted: pd.DataFrame, company: pd.DataFrame) -> pd.DataFrame:
    merged = extracted.merge(
        company,
        on=["Agreement No", "Party Number"],
        how="left",
        suffixes=("_ext", "_co"),
    )

    results = []
    for _, row in merged.iterrows():
        has_company_row = pd.notna(row.get("Company Name")) or pd.notna(row.get("Company Address"))

        name_sim = similarity(row["_name_norm_ext"], row.get("_name_norm_co", "") or "")
        addr_sim = similarity(row["_addr_norm_ext"], row.get("_addr_norm_co", "") or "")

        if not has_company_row:
            name_sim = 0.0
            addr_sim = 0.0

        status = classify(name_sim, addr_sim, has_company_row)

        results.append(
            {
                "Agreement No": row["Agreement No"],
                "Party Number": row["Party Number"],
                "Status": status,
                "Extracted Name": row["Name"],
                "Company Name": row.get("Company Name"),
                "Name Similarity %": round(name_sim, 1) if has_company_row else None,
                "Extracted Address": row["Address"],
                "Company Address": row.get("Company Address"),
                "Address Similarity %": round(addr_sim, 1) if has_company_row else None,
                "Extracted Cheque No": row["Cheque No"],
                "Source File": row["Source File"],
            }
        )

    report = pd.DataFrame(results)
    report = report.sort_values(["Agreement No", "Party Number"]).reset_index(drop=True)
    return report


# --------------------------------------------------------------------------
# Step 4: Write output (with a summary sheet)
# --------------------------------------------------------------------------
def write_report(report: pd.DataFrame, path: str) -> None:
    summary = (
        report["Status"]
        .value_counts()
        .rename_axis("Status")
        .reset_index(name="Count")
        .sort_values("Status")
    )

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        report.to_excel(writer, sheet_name="Reconciliation", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

        # Light auto-fit of column widths for readability
        for sheet_name, df in [("Reconciliation", report), ("Summary", summary)]:
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns, start=1):
                max_len = max(
                    [len(str(col))] + [len(str(v)) for v in df[col].astype(str).head(500)]
                )
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(
                    max(12, max_len + 2), 60
                )


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    extracted_path = sys.argv[1] if len(sys.argv) > 1 else EXTRACTED_FILE
    company_path = sys.argv[2] if len(sys.argv) > 2 else COMPANY_FILE
    output_path = sys.argv[3] if len(sys.argv) > 3 else OUTPUT_FILE

    print(f"Loading extracted notices from: {extracted_path}")
    extracted = load_extracted(extracted_path)
    print(f"  -> {len(extracted)} party rows loaded")

    print(f"Loading company master data from: {company_path}")
    company = load_company_data(company_path)
    print(f"  -> {len(company)} party rows unpivoted from wide format")

    print("Reconciling...")
    report = reconcile(extracted, company)

    print(f"Writing report to: {output_path}")
    write_report(report, output_path)

    print("\nSummary:")
    print(report["Status"].value_counts().to_string())
    print("\nDone.")


if __name__ == "__main__":
    main()
