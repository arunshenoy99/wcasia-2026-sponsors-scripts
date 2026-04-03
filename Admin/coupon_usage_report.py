#!/usr/bin/env python3
"""
Map sponsor tix_coupon rows to Camptix coupon-summary counts.

Inputs/outputs and file patterns: Admin/README.md. Stdlib only; not imported by other repo code.
"""

import csv
from pathlib import Path

ADMIN = Path(__file__).resolve().parent
COUPONS_CSV = ADMIN / "Camptix Coupons Generator for WordCamp Asia 2026 - Sponsors.csv"
OUTPUT_CSV = ADMIN / "coupon_usage_report.csv"
# If the Camptix Coupons sheet is missing, reuse last report for sponsor order + names.
ORDER_FALLBACK_CSV = OUTPUT_CSV


def load_sponsor_coupon_order():
    """(coupon, sponsor_name) in Camptix Coupons sheet order, or fallback CSV order."""
    if COUPONS_CSV.is_file():
        order = []
        with open(COUPONS_CSV, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                coupon = (row.get("tix_coupon") or "").strip()
                name = (row.get("Name") or "").strip()
                if coupon:
                    order.append((coupon, name or coupon))
        return order
    if ORDER_FALLBACK_CSV.is_file():
        order = []
        with open(ORDER_FALLBACK_CSV, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                coupon = (row.get("tix_coupon") or "").strip()
                name = (row.get("sponsor_name") or "").strip()
                if coupon:
                    order.append((coupon, name or coupon))
        return order
    raise FileNotFoundError(
        f"Need either {COUPONS_CSV} or {ORDER_FALLBACK_CSV} for sponsor coupon list."
    )


def get_latest_summary_csv():
    """Pick the newest camptix summary coupon file in Admin."""
    candidates = sorted(ADMIN.glob("camptix-summary-coupon-code-*.csv"))
    if not candidates:
        raise FileNotFoundError(
            f"No summary file found matching {ADMIN / 'camptix-summary-coupon-code-*.csv'}"
        )
    return candidates[-1]


def main():
    # 1. Sponsor coupons in file order (Camptix Coupons sheet, or last report)
    coupon_order = load_sponsor_coupon_order()
    summary_csv = get_latest_summary_csv()

    # 2. Read usage counts from Camptix coupon summary (Coupon code -> Count)
    usage = {}
    with open(summary_csv, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            code = (row.get("Coupon code") or "").strip()
            if code and code.lower() != "none":
                try:
                    usage[code] = int(row.get("Count", 0))
                except ValueError:
                    usage[code] = 0

    # 3. Build report in same order as coupons CSV (top to bottom)
    rows = []
    for coupon, name in coupon_order:
        count = usage.get(coupon, 0)
        rows.append((coupon, count, name))

    # 4. Write CSV: coupon, count, sponsor_name
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["tix_coupon", "claimed_count", "sponsor_name"])
        for coupon, count, name in rows:
            w.writerow([coupon, count, name])

    # 5. Print summary (coupon -> number)
    print("Coupon -> Claimed count\n")
    for coupon, count, name in rows:
        label = f" ({name})" if name else ""
        print(f"{coupon}{label}\t{count}")
    print(f"\nSummary source: {summary_csv}")
    print(f"\nReport saved to: {OUTPUT_CSV}")

if __name__ == "__main__":
    main()
