# FreeScout Email Automation

Automated email sending tool for WordCamp Asia 2026 sponsor outreach via FreeScout.

## Overview

This tool reads sponsor data from an Excel/CSV file, maps sponsor types to FreeScout email templates, fills placeholders, and automatically sends personalized emails through FreeScout's web interface using browser automation.

## Features

- Reads sponsor data from Excel/CSV files
- Automatically identifies sponsor type column
- Maps sponsor types to appropriate FreeScout templates
- Fills placeholders in email templates
- Extracts and sets subject lines
- Shows email preview and asks for confirmation before sending
- Logs success/failure for each email
- Generates summary report

## Setup

### 1. Create Virtual Environment

```bash
cd "Sponsors/Reachout"
python3 -m venv venv
```

### 2. Activate Virtual Environment

**On macOS/Linux:**
```bash
source venv/bin/activate
```

**On Windows:**
```bash
venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3b. Python version (optional)

If you use [pyenv](https://github.com/pyenv/pyenv) or similar, `Reachout/.python-version` pins a tested Python (e.g. `3.12.11`). It is optional; any Python that satisfies `requirements.txt` is fine.

### 4. Configure FreeScout Credentials

Create a `.env` file in the `Reachout` folder:

```env
FREESCOUT_URL=https://your-freescout-instance.com
FREESCOUT_EMAIL=your-email@example.com
FREESCOUT_PASSWORD=your-password

# Optional settings
HEADLESS_MODE=false
BROWSER_WAIT_TIME=10
# Speed: 1.0=normal, 0.5=faster, 0.4=fast (reduces fixed waits between actions)
BROWSER_DELAY_SCALE=1.0

# Optional: if your sheet uses a different primary email column (default is "Email")
# EXCEL_EMAIL_COLUMN=Alternate Emails

# Optional: only when FILTER_BY_OUTREACH=true in config.py — must match "Assigned Team Member" in the sheet
# OUTREACH_FILTER_VALUE=Your Name
```

Copy `.env.example` to `.env` and edit (never commit `.env`):

```bash
cp .env.example .env
# Then edit .env with your credentials
```

### 5. Input files for `main.py` (two workflows)

**Workflow A — standard spreadsheet (template from Sales Strategy)**  
`main.py` reads `.xlsx` / `.xls` / `.csv`. Required columns:

| Column | Purpose |
|--------|---------|
| **Company Name** | Placeholders, logging |
| **Contact Person** | Placeholders (including first name) |
| Primary email column | Default name **`Email`**; override with `EXCEL_EMAIL_COLUMN` in `.env` if your export uses another header (e.g. `Alternate Emails`) |
| **Sales Strategy** | Must contain sponsor-type values that map to FreeScout templates (see **Template mapping** below) |

If **`Template Name`** is present on every data row, the reader treats the file as a **round CSV** and uses that column instead of looking up template from Sales Strategy.

**Workflow B — round CSV from `extract_round_leads.py`**  
Input: your **master** Excel/CSV (prospects workbook). Output: a CSV with at least **Email**, **Company Name**, **Contact Person**, **Template Name** (and optional **Status**, **Sales Strategy**, **Assigned Team Member**). Feed that file to `main.py` the same way as workflow A.

**Filtering by assignee (optional):**  
By default all rows run (`FILTER_BY_OUTREACH = False` in `config.py`). To limit rows, set `FILTER_BY_OUTREACH = True` and set `OUTREACH_FILTER_VALUE` in `.env` to match **Assigned Team Member** exactly.

## Usage

**Important:** Make sure your virtual environment is activated before running the script.

### Basic Usage

```bash
python main.py path/to/sponsors.xlsx
```

Or run without arguments to be prompted for the file path:

```bash
python main.py
```

### Round workflow (status-based leads)

1. **Extract** from the **master** file (`.xlsx` / `.xls` / `.csv`). The master must include:
   - **`Status`** (configurable via `STATUS_COLUMN` in `config.py`) — empty cells are treated as *New - haven't been emailed* when that status is allowed.
   - **`Company Name`**, **`Contact Person`** (names from `EXCEL_COLUMNS` in `config.py`).
   - At least one of the email columns in `EMAIL_COLUMN_PRIORITY` (default: Alternate Email v2 → Alternate Email → Email).

   ```bash
   python extract_round_leads.py "path/to/master.xlsx" -o round_leads.csv
   # Optional: single status, or date filter on Last Contact Date
   python extract_round_leads.py "path/to/master.xlsx" -s "First Email Sent" -o round_first.csv
   python extract_round_leads.py "path/to/master.xlsx" --before-date 2026-02-01 -o round_stale.csv
   ```

2. **Send** with `main.py` using the round CSV as the only argument (round file includes **Template Name**):

   ```bash
   python main.py round_leads.csv
   ```

Configure **`STATUS_TO_TEMPLATE`**, **`EMAIL_COLUMN_PRIORITY`**, and related keys in `config.py`.  
**`email_utils.py`** is only used by `extract_round_leads.py` (parsing messy email cells); it is not run standalone.

### Approval list → company + email CSV (optional)

**Standalone** script (does not call `main.py` or FreeScout).  
**Input (`-i`):** your CSV export from an internal approval list. Column names are detected flexibly (e.g. company/sponsor name + contact/email columns).  
**Output (`-o`):** CSV with `Company Name`, `Email`, and `Listed on WCAsia 2026 site` (matched against the public sponsors list hardcoded in the script — update `PUBLISHED_SPONSORS` when the public page changes).

```bash
python build_approval_company_email_csv.py -i path/to/approval_export.csv -o path/to/out.csv
```

Do not commit input/output CSVs if they contain contact data.

### FreeScout selector debugging (optional)

**Standalone** helper: logs in with **`freescout_automation.py`**, opens a conversation, prints DOM hints for Reply / editor / Send. Use when your FreeScout skin breaks selectors.

```bash
python explore_freescout_selectors.py [search_email]
```

Requires `.env` with `FREESCOUT_URL`, `FREESCOUT_EMAIL`, `FREESCOUT_PASSWORD`. Browser stays visible; see file docstring for details.

**To deactivate the virtual environment when done:**
```bash
deactivate
```

### Email Sending

**Email sending is now enabled.** The automation will:
- Fill in all the email fields
- Show you a preview
- Ask for confirmation (y/n/skip)
- **Click the Send button** to send the email if you confirm with 'y'

**⚠️ IMPORTANT:** Make sure you review each email preview carefully before confirming, as emails will be sent immediately after confirmation.

### Workflow

1. The script reads your Excel/CSV file
2. It identifies sponsor types and maps them to FreeScout templates
3. A browser window opens and logs into FreeScout
4. For each sponsor:
   - Opens a new conversation
   - Fills the "To" field
   - Selects the appropriate template
   - Fills placeholders in the template
   - Extracts and sets the subject line
   - Shows you a preview of the email
   - **Asks for confirmation before sending** (y/n/skip)
   - Sends the email if confirmed
5. Displays a summary report

### Confirmation Options

When prompted for each email:
- **y**: Send the email
- **n**: Cancel and don't send
- **skip**: Skip this email and move to the next

## Placeholders

The tool currently supports these placeholders in templates:

- `[Company Name]` → Replaced with Company Name from Excel
- `[Customer Company POC][Prospective Sponsor's Name]` → Replaced with Contact Person
- `[Customer Company POC]` → Replaced with Contact Person
- `[Prospective Sponsor's Name]` → Replaced with Contact Person (first name)
- `{%customer.firstName%}` → Replaced with Contact Person's first name (FreeScout-style, used in follow-up template)

Additional placeholders can be added by updating `PLACEHOLDER_MAPPINGS` in `config.py`.

## Template Mapping

Sponsor types are automatically mapped to FreeScout templates:

| Sponsor Type | FreeScout Template |
|-------------|-------------------|
| Past WCAsia Sponsor | Past Sponsors (Flagships & Local) Outreach - 2026 Sponsors |
| Past Flagship Sponsor | Past Sponsors (Flagships & Local) Outreach - 2026 Sponsors |
| Past Sponsor | Past Sponsors (Flagships & Local) Outreach - 2026 Sponsors |
| New WP | New WP Outreach - 2026 Sponsors |
| New Non-WP | Non-WP Outreach - 2026 Sponsors |
| Not WP Related | Non-WP Outreach - 2026 Sponsors |

(Edit `SPONSOR_TYPE_TO_TEMPLATE` in `config.py` if your FreeScout template names differ.)

## Troubleshooting

### Browser Issues

- If ChromeDriver fails, it will be automatically downloaded by `webdriver-manager`
- If you see browser-related errors, try setting `HEADLESS_MODE=false` in `.env` to see what's happening
- Make sure Chrome browser is installed on your system

### Login Issues

- Verify your FreeScout URL, email, and password in `.env`
- Check if FreeScout requires 2FA (not currently supported)
- The tool tries multiple common selectors for login fields - if your FreeScout instance uses different selectors, you may need to update `freescout_automation.py`

### Template Selection Issues

- Ensure template names in FreeScout exactly match the mapping in `config.py`
- The tool looks for templates by visible text - make sure the template name is visible in the dropdown

### Placeholder Issues

- If placeholders aren't being filled, check that the Excel column names match `EXCEL_COLUMNS` in `config.py`
- Placeholders are case-sensitive - ensure they match exactly in your templates

### Subject Line Extraction

- The tool looks for lines starting with "Subject -" in the template
- If your templates use a different format, you may need to update the `extract_template_content()` method

## File structure and how everything is used

### Typical data flow

```
Master prospects file (.xlsx / .csv)
        │
        │  extract_round_leads.py  (optional)
        ▼
Round CSV (Email, Company Name, Contact Person, Template Name, …)
        │
        │  main.py
        ├──────────────────────► sponsor_reader.py  (parse rows, templates)
        │
        └──────────────────────► freescout_automation.py  (Selenium / FreeScout)
```

You can also point **`main.py`** directly at a standard spreadsheet (workflow A above) and skip `extract_round_leads.py`.

### Entry-point scripts (run from CLI)

| Script | Purpose | Typical inputs | Typical outputs |
|--------|---------|----------------|-----------------|
| **`main.py`** | Orchestrates send flow | Path to `.xlsx` / `.csv` (workflows A/B in **§5** above) | `logs/sent_emails_*.log`, emails in FreeScout |
| **`extract_round_leads.py`** | Build round CSV from master | Master workbook path; optional `-o`, `-s`, `--before-date` | `round_YYYYMMDD.csv` or path from `-o` |
| **`build_approval_company_email_csv.py`** | Company + email + public-site flag | `-i` approval CSV, `-o` output CSV | Written CSV only |
| **`explore_freescout_selectors.py`** | Debug UI selectors | Optional search email; `.env` required | Console / browser inspection |

### Shared modules (imported only — do not run directly)

| File | Imported by | Role |
|------|-------------|------|
| **`config.py`** | `main.py`, `sponsor_reader.py`, `freescout_automation.py`, `extract_round_leads.py`, `explore_freescout_selectors.py` | Env-loaded settings, column names, template maps |
| **`sponsor_reader.py`** | `main.py` | Load spreadsheet, resolve sponsor type / template, build row list |
| **`freescout_automation.py`** | `main.py`, `explore_freescout_selectors.py` | Login, compose, templates, send / reply automation |
| **`email_utils.py`** | `extract_round_leads.py` only | Parse multiple emails from one cell into one string |

### Other tracked files

| File | Role |
|------|------|
| **`requirements.txt`** | `pip install -r` dependencies |
| **`.env.example`** | Copy to `.env`; documents optional vars |
| **`.python-version`** | Optional pyenv pin |
| **`README.md`** | This documentation |

```
Reachout/
├── main.py
├── extract_round_leads.py
├── build_approval_company_email_csv.py
├── explore_freescout_selectors.py
├── email_utils.py
├── sponsor_reader.py
├── freescout_automation.py
├── config.py
├── requirements.txt
├── .python-version
├── .env.example
├── .env                 # local only
└── README.md
```

## Notes

- The tool uses Selenium for browser automation, which requires Chrome browser
- Emails are sent one at a time with confirmation prompts
- The browser window will remain open during the process so you can monitor it
- If the process is interrupted (Ctrl+C), the browser will be closed gracefully

## Support

If you encounter issues:
1. Check that all dependencies are installed
2. Verify your `.env` file has correct credentials
3. Ensure your Excel file has the required columns
4. Check the console output for specific error messages
5. Try running with `HEADLESS_MODE=false` to see the browser actions

