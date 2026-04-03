# Admin — Camptix coupon usage report

Small helper to align **sponsor coupon codes** from your local Camptix Coupons sheet with **usage counts** from a Camptix **coupon summary** export.

## What runs here

| Artifact | Role |
|----------|------|
| **`coupon_usage_report.py`** | Only executable in this folder; stdlib only |

There are no Python imports from `Reachout/` or `Agreements/`.

## What stays local (not in git)

This repo’s `.gitignore` ignores `*.csv` and other data files. Keep these **only on your machine** (they often contain sponsor or attendee-related data):

- Camptix coupon summary exports, e.g. `camptix-summary-coupon-code-YYYY-MM-DD.csv`
- `Camptix Coupons Generator for WordCamp Asia 2026 - Sponsors.csv` (or your equivalent with `tix_coupon` + sponsor rows)
- Generated `coupon_usage_report.csv` (optional; contains sponsor names from your sheet)

Do not commit those files.

## Prerequisites

- Python 3 (stdlib only; no `pip` packages)

## Inputs and outputs

| Kind | File pattern | Required columns / notes |
|------|--------------|---------------------------|
| **Coupon summary** (input) | `camptix-summary-coupon-code-*.csv` | **`Coupon code`**, **`Count`** (Camptix export). Newest file by name is chosen automatically. |
| **Sponsor coupons sheet** (input, optional if report exists) | e.g. `Camptix Coupons Generator for WordCamp Asia 2026 - Sponsors.csv` | **`tix_coupon`**, **`Name`** — row order defines report order. |
| **Report** (output) | `coupon_usage_report.csv` | Written each run; see below. |

## Usage

1. Copy the latest Camptix **coupon summary** CSV into `Admin/`.
2. Copy your **sponsor coupons** CSV into `Admin/` (same folder as the script), or keep an existing `coupon_usage_report.csv` to reuse coupon order and names.
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
