"""
CLI: build Company Name + Email CSV from a local approval export; flag matches to the public sponsors page.

Standalone (not used by main.py). The published-sponsors list is read from a local
file (default: published_sponsors.txt next to this script; not committed) — see
published_sponsors.example.txt for the format.

Usage and data policy: Reachout/README.md (approval list section).
"""
import argparse
import csv
import os
import re
import sys

DEFAULT_PUBLISHED_SPONSORS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "published_sponsors.txt"
)


def _norm(s):
    if not s or not isinstance(s, str):
        return ""
    return re.sub(r"\s+", " ", s.strip().lower())


def load_published_sponsors(path: str) -> set:
    """Load published sponsor names from a text file (one per line; '#' comments ignored)."""
    sponsors = set()
    if not path or not os.path.isfile(path):
        return sponsors
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            name = _norm(line)
            if name:
                sponsors.add(name)
    return sponsors


def _listed_on_site(company_name: str, published: set) -> bool:
    """Check if company appears in the published sponsors list (case-insensitive, flexible match)."""
    n = _norm(company_name)
    if not n or not published:
        return False
    if n in published:
        return True
    for pub in published:
        if pub and (pub in n or n in pub):
            return True
    return False


def main():
    parser = argparse.ArgumentParser(
        description="Merge company + email from an approval CSV and flag public-page matches."
    )
    parser.add_argument(
        "-i", "--input",
        required=True,
        help="Path to your approval list CSV export (local only; do not commit).",
    )
    parser.add_argument(
        "-o", "--output",
        required=True,
        help="Path to write the output CSV (local only; do not commit if it contains contact data).",
    )
    parser.add_argument(
        "-p", "--published-sponsors",
        default=DEFAULT_PUBLISHED_SPONSORS_FILE,
        help="Path to the published sponsors list (default: published_sponsors.txt next to this script). "
             "See published_sponsors.example.txt.",
    )
    args = parser.parse_args()
    input_path = args.input
    output_path = args.output

    published = load_published_sponsors(args.published_sponsors)
    if not published:
        print(
            f"Warning: no published sponsors loaded from {args.published_sponsors!r}; "
            "'Listed on public sponsors site' will be 'No' for every row. "
            "Copy published_sponsors.example.txt to published_sponsors.txt to enable matching.",
            file=sys.stderr,
        )

    rows_out = []
    with open(input_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames or []
        email_col = None
        company_col = None
        for fn in fieldnames:
            fn_clean = _norm(fn)
            if "contact email" in fn_clean or (fn_clean == "email"):
                email_col = fn
            if "company" in fn_clean and "sponsor" in fn_clean and "name" in fn_clean and "website" not in fn_clean:
                company_col = fn
            elif "company" in fn_clean and "sponsor" in fn_clean and company_col is None:
                company_col = fn
        if not email_col:
            for fn in fieldnames:
                if "email" in _norm(fn):
                    email_col = fn
                    break
        if not company_col:
            company_col = "Company / Sponsor Name"

        for row in reader:
            company = (row.get(company_col) or "").strip()
            email = (row.get(email_col) or "").strip()
            if not company and not email:
                continue
            listed = "Yes" if _listed_on_site(company, published) else "No"
            rows_out.append({"Company Name": company, "Email": email, "Listed on public sponsors site": listed})

    with open(output_path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["Company Name", "Email", "Listed on public sponsors site"])
        w.writeheader()
        w.writerows(rows_out)

    listed_count = sum(1 for r in rows_out if r["Listed on public sponsors site"] == "Yes")
    print(f"Wrote {len(rows_out)} rows to {output_path}")
    print(f"Listed on public sponsors page: {listed_count} of {len(rows_out)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
