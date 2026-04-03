"""
Build a CSV with Company Name and Email from a local approval-list export,
and verify each company against names on the public sponsors page (see README).
"""
import argparse
import csv
import re
import sys

# Normalized names aligned with https://asia.wordcamp.org/2026/sponsors/ (public page; update as needed)
PUBLISHED_SPONSORS = {
    "elementor", "hostinger", "jetpack", "paypal", "pressable", "salesforce", "wordpress.com",
    "google", "woo",
    "astra", "automattic for agencies", "bluehost", "yoast",
    "easy wp powered by spaceship", "kinsta", "knowledge pillars education inc", "rank math", "rtcamp", "wp rocket",
    "hosting.com",
    "flexicloud", "greengeeks web hosting", "rumahweb indonesia",
    "flowmattic", "wpexperts",
}


def _norm(s):
    if not s or not isinstance(s, str):
        return ""
    return re.sub(r"\s+", " ", s.strip().lower())


def _listed_on_site(company_name: str) -> bool:
    """Check if company appears on the published WCAsia 2026 sponsors page (flexible match)."""
    n = _norm(company_name)
    if not n:
        return False
    if n in PUBLISHED_SPONSORS:
        return True
    for pub in PUBLISHED_SPONSORS:
        if pub in n or n in pub:
            return True
    if "spaceship" in n and ("easy" in n or "easywp" in n):
        return True
    if "wordpress.com" in n:
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
    args = parser.parse_args()
    input_path = args.input
    output_path = args.output

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
            listed = "Yes" if _listed_on_site(company) else "No"
            rows_out.append({"Company Name": company, "Email": email, "Listed on WCAsia 2026 site": listed})

    with open(output_path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["Company Name", "Email", "Listed on WCAsia 2026 site"])
        w.writeheader()
        w.writerows(rows_out)

    listed_count = sum(1 for r in rows_out if r["Listed on WCAsia 2026 site"] == "Yes")
    print(f"Wrote {len(rows_out)} rows to {output_path}")
    print(f"Listed on public sponsors page: {listed_count} of {len(rows_out)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
