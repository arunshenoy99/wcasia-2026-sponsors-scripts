# Sponsor Agreement Template Generator

This module generates Word document templates for WordCamp Asia 2026 sponsor agreements.

## Overview

The `generate_templates.py` script creates Word document templates for different sponsor tiers:
- Super Admin
- Admin
- Editor
- Author
- Contributor
- Subscriber
- Viewer
- Addon Agreement

## Setup

1. Create a virtual environment:
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the main script:
```bash
python generate_templates.py
```

This will generate all sponsor agreement templates as Word documents in the current directory.

## Requirements

- Python 3.8+
- python-docx library

## Files

- `generate_templates.py` - Main script to generate templates
- `check_numbering.py` - Utility to check document numbering structure
- `compare_docs.py` - Utility to compare original and generated documents
- `inspect_doc.py` - Utility to inspect document structure

## Notes

- The script uses an original 2025 agreement document as a reference for formatting
- Generated templates preserve numbering, formatting, and structure from the original
- Placeholders like `[SPONSOR_NAME]` and `[DATE]` are included in templates for manual filling

