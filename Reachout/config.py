"""
Configuration settings for FreeScout email automation.

Imported by main.py, sponsor_reader.py, freescout_automation.py,
extract_round_leads.py, and explore_freescout_selectors.py.
See Reachout/README.md for script map, input files, and env vars.
"""
import os
from dotenv import load_dotenv

load_dotenv()

# Excel: primary email column for master/bounced lists (default "Email"). Example: Alternate Emails
_EXCEL_EMAIL_COLUMN = os.getenv("EXCEL_EMAIL_COLUMN", "").strip()

# FreeScout Configuration
FREESCOUT_URL = os.getenv('FREESCOUT_URL', '')
FREESCOUT_EMAIL = os.getenv('FREESCOUT_EMAIL', '')
FREESCOUT_PASSWORD = os.getenv('FREESCOUT_PASSWORD', '')

# Browser Settings
HEADLESS_MODE = os.getenv('HEADLESS_MODE', 'false').lower() == 'true'
BROWSER_WAIT_TIME = int(os.getenv('BROWSER_WAIT_TIME', '10'))
# Scale for fixed delays (1.0 = normal; 0.4 = faster, 0.5 = slightly faster). Lower = quicker but may be flaky on slow networks.
BROWSER_DELAY_SCALE = float(os.getenv('BROWSER_DELAY_SCALE', '1.0'))

# Sponsor Type to FreeScout Template Mapping
SPONSOR_TYPE_TO_TEMPLATE = {
    "Past WCAsia Sponsor": "Past Sponsors (Flagships & Local) Outreach - 2026 Sponsors",
    "Past Flagship Sponsor": "Past Sponsors (Flagships & Local) Outreach - 2026 Sponsors",
    "Past Sponsor": "Past Sponsors (Flagships & Local) Outreach - 2026 Sponsors",
    "New WP": "New WP Outreach - 2026 Sponsors",
    "New Non-WP": "Non-WP Outreach - 2026 Sponsors",
    "Not WP Related": "Non-WP Outreach - 2026 Sponsors"
}

# Placeholder Mappings
# These map Excel column names to placeholder names in templates
PLACEHOLDER_MAPPINGS = {
    "[Company Name]": "Company Name",
    "[Customer Company POC][Prospective Sponsor's Name]": "Contact Person",
    "{%customer.firstName%}": "Contact Person (first name)",
    # Add more mappings as placeholders are identified
}

# Excel Column Names (expected)
EXCEL_COLUMNS = {
    "COMPANY_NAME": "Company Name",
    "EMAIL": _EXCEL_EMAIL_COLUMN or "Email",
    "CONTACT_PERSON": "Contact Person",
    "SPONSOR_TYPE": "Sales Strategy",  # This column contains the sponsor type dropdown values
    "INITIAL_OUTREACH_BY": "Assigned Team Member",  # Used when FILTER_BY_OUTREACH is True
}

# Filter settings
FILTER_BY_OUTREACH = False  # Set to True to only process rows where Assigned Team Member matches OUTREACH_FILTER_VALUE
OUTREACH_FILTER_VALUE = os.getenv("OUTREACH_FILTER_VALUE", "").strip()  # e.g. your name as it appears in the sheet

# Round lead list (extract_round_leads.py and round CSV)
STATUS_COLUMN = "Status"
STATUS_TO_TEMPLATE = {
    "New - haven't been emailed": "New WP Outreach - 2026 Sponsors",
    "First Email Sent": "2026 Sponsor Email - 1st Sponsor Follow up Email",
    "First Follow Up": "2026 Sponsors - 2nd Sponsor Follow up Email",
}
EMAIL_COLUMN_PRIORITY = ["Alternate Email v2", "Alternate Email", "Email"]
TEMPLATE_NAME_COLUMN = "Template Name"
LAST_CONTACT_DATE_COLUMN = "Last Contact Date"
# When using round CSV, file should have: Email, Company Name, Contact Person, Template Name (and optionally Status)

# Reply-to-thread workflow: templates whose name contains this string will reply in existing conversation instead of new
REPLY_TO_THREAD_TEMPLATE_PATTERN = "Follow up"