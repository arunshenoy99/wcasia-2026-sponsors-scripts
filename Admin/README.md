# Admin — Camptix coupon usage report

Small helper to align **sponsor coupon codes** from your local Camptix Coupons sheet with **usage counts** from a Camptix **coupon summary** export.

## What stays local (not in git)

This repo’s `.gitignore` ignores `*.csv` and other data files. Keep these **only on your machine** (they often contain sponsor or attendee-related data):

- Camptix coupon summary exports, e.g. `camptix-summary-coupon-code-YYYY-MM-DD.csv`
- `Camptix Coupons Generator for WordCamp Asia 2026 - Sponsors.csv` (or your equivalent with `tix_coupon` + sponsor rows)
- Generated `coupon_usage_report.csv` (optional; contains sponsor names from your sheet)

Do not commit those files.

## Prerequisites

- Python 3 (stdlib only; no `pip` packages)

## Usage

1. Place the latest `camptix-summary-coupon-code-*.csv` in this `Admin/` folder (export from Camptix).
2. Place your sponsor coupons sheet CSV here (must include a `tix_coupon` column and `Name` for sponsor label). Row order in that file defines the report order.
3. Run:

```bash
cd Admin
python3 coupon_usage_report.py
```

The script picks the **newest** `camptix-summary-coupon-code-*.csv` by filename sort. It writes `coupon_usage_report.csv` and prints which summary file was used.

If the sponsor coupons sheet is missing but a previous `coupon_usage_report.csv` exists, the script reuses that file for coupon order and names (so you can refresh counts from a new summary only).

## Summary CSV format

Expected columns (from Camptix export): **`Coupon code`**, **`Count`**. Rows with coupon code `None` are skipped for counting.

## Output columns

- `tix_coupon` — coupon code  
- `claimed_count` — count from the summary for that code  
- `sponsor_name` — from the `Name` column of the coupons sheet (or prior report)
