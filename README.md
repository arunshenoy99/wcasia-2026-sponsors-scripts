# WordCamp Asia 2026 - Sponsors Management

This repository contains tools and scripts for managing sponsor agreements and outreach for WordCamp Asia 2026.

## Repository Structure

```
Sponsors/
├── Agreements/          # Sponsor agreement template generation
│   ├── generate_templates.py    # Main script to generate Word document templates
│   ├── check_numbering.py       # Utility to check document numbering
│   ├── compare_docs.py          # Utility to compare documents
│   └── inspect_doc.py           # Utility to inspect document structure
│
├── Admin/               # Camptix coupon usage helper (see Admin/README.md)
│   ├── coupon_usage_report.py   # Build sponsor coupon vs summary counts
│   └── README.md
│
├── Reachout/           # FreeScout email automation (see Reachout/README.md)
│   ├── main.py                  # CLI: send flow
│   ├── extract_round_leads.py   # CLI: master → round CSV
│   ├── build_approval_company_email_csv.py  # CLI: optional approval CSV helper
│   ├── explore_freescout_selectors.py       # CLI: optional selector debugging
│   ├── email_utils.py           # Used by extract_round_leads only
│   ├── sponsor_reader.py        # Used by main only
│   ├── freescout_automation.py  # Used by main + explore script
│   ├── config.py                # Shared settings (imported by scripts above)
│   ├── requirements.txt
│   ├── .env.example
│   ├── .python-version          # Optional pyenv pin
│   └── README.md                # Scripts reference, inputs/outputs, data flow
│
├── .gitignore          # Git ignore rules
└── README.md           # This file
```

## Quick Start

### Agreements Module

The Agreements module generates Word document templates for different sponsor tiers.

**Setup:**
```bash
cd Agreements
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install python-docx
```

**Usage:**
```bash
python generate_templates.py
```

This generates Word document templates for all sponsor tiers (Super Admin, Admin, Editor, Author, Contributor, Subscriber, Viewer) and Addon agreements.

### Reachout Module

The Reachout module automates email sending to sponsors via FreeScout.

**Setup:**
```bash
cd Reachout
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

**Configuration:**
Create a `.env` file with your FreeScout credentials:
```env
FREESCOUT_URL=https://your-freescout-instance.com
FREESCOUT_EMAIL=your-email@example.com
FREESCOUT_PASSWORD=your-password
```

**Usage:**
```bash
python main.py path/to/sponsors.xlsx
```

See **`Reachout/README.md`** for the full **scripts reference** (which files import which, input columns, round vs standard workflow, and optional tools).

### Admin — Camptix coupon report

Sponsor coupon codes vs Camptix summary counts (local CSVs only; see `Admin/README.md`):

```bash
cd Admin
python3 coupon_usage_report.py
```

## Requirements

- Python 3.8+
- Chrome browser (for Reachout automation)
- Required Python packages (see individual module requirements.txt files)

## Notes

- Virtual environments (`venv/`) are excluded from git
- Environment files (`.env`) are excluded from git
- Log files and temporary files are excluded from git
- Excel/CSV data files and Word documents are excluded from git (add them manually if needed)

## Contributing

When adding new features:
1. Follow existing code structure
2. Update relevant README files
3. Test thoroughly before committing
4. Use meaningful commit messages

