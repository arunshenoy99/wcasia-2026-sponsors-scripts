"""
Configuration settings for FreeScout email automation
"""
import os
from dotenv import load_dotenv

load_dotenv()

# FreeScout Configuration
FREESCOUT_URL = os.getenv('FREESCOUT_URL', '')
FREESCOUT_EMAIL = os.getenv('FREESCOUT_EMAIL', '')
FREESCOUT_PASSWORD = os.getenv('FREESCOUT_PASSWORD', '')

# Browser Settings
HEADLESS_MODE = os.getenv('HEADLESS_MODE', 'false').lower() == 'true'
BROWSER_WAIT_TIME = int(os.getenv('BROWSER_WAIT_TIME', '10'))

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
    # Add more mappings as placeholders are identified
}

# Excel Column Names (expected)
EXCEL_COLUMNS = {
    "COMPANY_NAME": "Company Name",
    "EMAIL": "Email",
    "CONTACT_PERSON": "Contact Person",
    "SPONSOR_TYPE": "Sales Strategy",  # This column contains the sponsor type dropdown values
    "INITIAL_OUTREACH_BY": "Assigned Team Member"  # Filter column - only process rows where this equals "Arun"
}

# Filter settings
FILTER_BY_OUTREACH = True  # Set to False to process all rows regardless of "Assigned Team Member"
OUTREACH_FILTER_VALUE = "Yash"  # Only process rows where "Assigned Team Member" equals this value

