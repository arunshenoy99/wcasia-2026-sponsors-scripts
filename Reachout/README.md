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

You can copy `.env.example` as a template:

```bash
cp .env.example .env
# Then edit .env with your credentials
```

### 5. Prepare Excel File

Your Excel/CSV file should contain the following columns:

- **Email** (required): Email address of the sponsor
- **Company Name** (required): Company name for placeholders
- **Contact Person** (required): Contact person name for placeholders
- **Sales Strategy** (required): Dropdown column containing one of the following sponsor type values:
  - "Past WCAsia Sponsor"
  - "Past Flagship Sponsor"
  - "Past Sponsor"
  - "New WP"
  - "New Non-WP"
  - "Not WP Related"

**Notes:**
- By default, **all rows** are processed (`FILTER_BY_OUTREACH = False` in `config.py`). To limit by assignee, set `FILTER_BY_OUTREACH = True` and set `OUTREACH_FILTER_VALUE` in `.env` (or `config.py`) to the exact value in the **Assigned Team Member** column.
- If your file uses a different column for the primary email than **Email**, set `EXCEL_EMAIL_COLUMN` in `.env` (see `.env.example`).

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

For a round targeting specific statuses (e.g. "New - haven't been emailed", "First Email Sent"):

1. **Extract** leads from the master file into a round CSV (filters by status, resolves email from Alternate Email v2 / Alternate Email / Email):

   ```bash
   python extract_round_leads.py "path/to/2026 Sponsor Prospects - Assigning Roles v2.xlsx" [-o round_leads.csv]
   ```

2. **Send** using the round CSV (file already has Template Name; no Sales Strategy lookup):

   ```bash
   python main.py round_leads.csv
   ```

Configure status-to-template mapping and email column priority in `config.py` (`STATUS_TO_TEMPLATE`, `EMAIL_COLUMN_PRIORITY`).

### Approval list → company + email CSV (optional)

Build a small CSV from **your local** approval export (paths are required; do not commit inputs/outputs that contain contact data):

```bash
python build_approval_company_email_csv.py -i path/to/approval_export.csv -o path/to/out.csv
```

The script adds a column that checks company names against the public WordCamp Asia 2026 sponsors page list (maintain `PUBLISHED_SPONSORS` in the script as the public page changes).

### FreeScout selector debugging (optional)

If reply/send selectors break on your instance, run (visible browser):

```bash
python explore_freescout_selectors.py
```

Requires `.env` with FreeScout credentials. See docstring in that file.

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

## File Structure

```
Reachout/
├── main.py                       # Main orchestration script
├── extract_round_leads.py        # Extract round CSV from master (status filter, email resolution)
├── build_approval_company_email_csv.py  # Optional: company+email from approval CSV
├── explore_freescout_selectors.py     # Optional: debug UI selectors
├── email_utils.py                # Shared email extraction from cell values
├── sponsor_reader.py             # Excel/CSV reading and processing (supports round CSV)
├── freescout_automation.py       # Browser automation logic
├── config.py                     # Configuration settings
├── requirements.txt              # Python dependencies
├── .python-version               # Optional: pyenv / local Python pin
├── .env                          # Your credentials (not in git)
├── .env.example                  # Example environment file
└── README.md                     # This file
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

