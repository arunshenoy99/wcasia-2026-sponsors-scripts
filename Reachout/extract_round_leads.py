"""
Extract round lead list from 2026 Sponsor Prospects master file.
Filters by status, optional date (e.g. Last Contact Date before a date), resolves email
from Alternate Email v2 / Alternate Email / Email, and writes a round CSV for the send script.
"""
import argparse
import os
import sys
from datetime import datetime
from typing import Optional

import pandas as pd

from config import (
    EMAIL_COLUMN_PRIORITY,
    EXCEL_COLUMNS,
    LAST_CONTACT_DATE_COLUMN,
    STATUS_COLUMN,
    STATUS_TO_TEMPLATE,
)
from email_utils import extract_emails_as_single_string


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace from column names."""
    df = df.rename(columns={c: c.strip() for c in df.columns})
    return df


def _find_column(df: pd.DataFrame, name: str) -> Optional[str]:
    """Return actual column name from df that matches name (strip + case-insensitive)."""
    name_clean = name.strip().lower()
    for col in df.columns:
        if col.strip().lower() == name_clean:
            return col
    return None


def _resolve_email_for_row(row: pd.Series, df: pd.DataFrame) -> str:
    """First non-empty value from EMAIL_COLUMN_PRIORITY, then smart-extract emails as single string."""
    for col_name in EMAIL_COLUMN_PRIORITY:
        col = _find_column(df, col_name)
        if col is None:
            continue
        val = row.get(col)
        if pd.isna(val) or str(val).strip() == "":
            continue
        resolved = extract_emails_as_single_string(str(val))
        if resolved:
            return resolved
    return ""


def _parse_date(val) -> Optional[pd.Timestamp]:
    """Parse cell value to date; return None if missing or invalid."""
    if pd.isna(val) or val == "":
        return None
    try:
        return pd.to_datetime(val)
    except Exception:
        return None


def _row_date_before_or_empty(row: pd.Series, date_col: str, before: pd.Timestamp) -> bool:
    """True if row's date is missing/invalid (include) or is before 'before'."""
    val = row.get(date_col)
    d = _parse_date(val)
    if d is None:
        return True
    return d < before


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract round leads from 2026 Sponsor Prospects master file."
    )
    parser.add_argument(
        "master_path",
        help="Path to master Excel/CSV (e.g. 2026 Sponsor Prospects - Assigning Roles v2.xlsx)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output round CSV path (default: round_YYYYMMDD.csv in same dir as script)",
    )
    parser.add_argument(
        "-s",
        "--status",
        default=None,
        help="Only include this status (e.g. \"First Follow Up\"). Default: all statuses in config.",
    )
    parser.add_argument(
        "--before-date",
        default=None,
        metavar="DATE",
        help="Only include rows where date column is before this date (e.g. 2026-02-01). Use with --date-column.",
    )
    parser.add_argument(
        "--date-column",
        default=LAST_CONTACT_DATE_COLUMN,
        help=f"Column name for date filter (default: {LAST_CONTACT_DATE_COLUMN!r}).",
    )
    args = parser.parse_args()

    if not os.path.exists(args.master_path):
        print(f"Error: File not found: {args.master_path}", file=sys.stderr)
        return 1

    # Read master
    try:
        if args.master_path.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(args.master_path)
        else:
            df = pd.read_csv(args.master_path)
    except Exception as e:
        print(f"Error reading file: {e}", file=sys.stderr)
        return 1

    df = _normalize_columns(df)

    status_col = _find_column(df, STATUS_COLUMN)
    if status_col is None:
        print(f"Error: Status column '{STATUS_COLUMN}' not found. Columns: {list(df.columns)}", file=sys.stderr)
        return 1

    company_col = _find_column(df, EXCEL_COLUMNS["COMPANY_NAME"])
    contact_col = _find_column(df, EXCEL_COLUMNS["CONTACT_PERSON"])
    if company_col is None or contact_col is None:
        print(
            f"Error: Need columns '{EXCEL_COLUMNS['COMPANY_NAME']}' and '{EXCEL_COLUMNS['CONTACT_PERSON']}'. Found: {list(df.columns)}",
            file=sys.stderr,
        )
        return 1

    allowed_statuses = set(STATUS_TO_TEMPLATE.keys())
    if args.status is not None:
        status_filter = str(args.status).strip()
        if status_filter not in allowed_statuses:
            print(f"Error: --status must be one of {sorted(allowed_statuses)}", file=sys.stderr)
            return 1
        allowed_statuses = {status_filter}

    before_ts = None
    date_col = None
    if args.before_date is not None:
        try:
            before_ts = pd.to_datetime(args.before_date)
        except Exception as e:
            print(f"Error: --before-date must be a valid date (e.g. 2026-02-01): {e}", file=sys.stderr)
            return 1
        date_col = _find_column(df, args.date_column.strip())
        if date_col is None:
            print(f"Error: Date column {args.date_column!r} not found. Columns: {list(df.columns)}", file=sys.stderr)
            return 1

    rows_out = []

    # Excel often shows "New - haven't been emailed" in the dropdown but stores empty; treat empty as that status
    default_status = "New - haven't been emailed"
    for _, row in df.iterrows():
        status_val = str(row[status_col]).strip() if not pd.isna(row.get(status_col)) else ""
        if not status_val and default_status in allowed_statuses:
            status_val = default_status
        if status_val not in allowed_statuses:
            continue
        if before_ts is not None and date_col is not None and not _row_date_before_or_empty(row, date_col, before_ts):
            continue
        email_str = _resolve_email_for_row(row, df)
        if not email_str:
            continue
        template_name = STATUS_TO_TEMPLATE.get(status_val, "")
        if not template_name:
            continue
        company = str(row[company_col]).strip() if not pd.isna(row.get(company_col)) else ""
        contact = str(row[contact_col]).strip() if not pd.isna(row.get(contact_col)) else ""

        out_row = {
            "Email": email_str,
            "Company Name": company,
            "Contact Person": contact,
            "Template Name": template_name,
            "Status": status_val,
        }
        if "Sales Strategy" in df.columns:
            out_row["Sales Strategy"] = row.get("Sales Strategy", "")
        if EXCEL_COLUMNS.get("INITIAL_OUTREACH_BY") and _find_column(df, EXCEL_COLUMNS["INITIAL_OUTREACH_BY"]):
            out_row["Assigned Team Member"] = row.get(_find_column(df, EXCEL_COLUMNS["INITIAL_OUTREACH_BY"]), "")
        rows_out.append(out_row)

    if not rows_out:
        print("No rows matched (status and non-empty email).")
        return 0

    out_df = pd.DataFrame(rows_out)
    if args.output:
        out_path = args.output
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        out_path = os.path.join(base_dir, f"round_{datetime.now().strftime('%Y%m%d')}.csv")
    out_df.to_csv(out_path, index=False, encoding="utf-8")
    print(f"Wrote {len(rows_out)} leads to {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
