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

## Reference document (required for correct numbering)

Every script in this folder reads one **2025 reference agreement** — a signed Super Admin
sponsorship agreement from the previous year, used purely as a formatting/numbering source. It is
**not committed** (`*.docx` is git-ignored, and it contains sponsor data), so you must place your
own copy in the `Agreements/` folder before running anything here.

The reference file is discovered automatically (`reference_doc.py`), in this order:

1. A file named exactly **`2025-super-admin-sponsorship-agreement.docx`** (rename your copy to
   this for an unambiguous match), **or**
2. any `*Super Admin*Sponsorship Agreement*.docx` in this folder that is not a generated 2026
   template — so a file like `<Company> WordCamp Asia 2025 Super Admin Sponsorship Agreement.docx`
   is picked up as-is.

Run the scripts from inside `Agreements/` (`cd Agreements` first) so the folder is searched.

- `generate_templates.py` copies the Word numbering definitions out of this file so the
  generated templates match the original numbering. If no reference file is found, it still runs
  but prints `Warning: Could not copy numbering definitions: ...` and the output numbering may be wrong.
- `check_numbering.py`, `compare_docs.py`, and `inspect_doc.py` **require** the reference file and
  exit with a clear message if it is absent.

## Usage

All scripts are run with **no arguments** — input/output filenames are hard-coded — and must be
run from the `Agreements/` directory.

Generate the templates:
```bash
cd Agreements
python generate_templates.py
```

This writes one `.docx` per tier plus the addon template into the current directory:

```
WordCamp Asia 2026 Super Admin Sponsorship Agreement Template.docx
WordCamp Asia 2026 Admin Sponsorship Agreement Template.docx
... (Editor, Author, Contributor, Subscriber, Viewer)
WordCamp Asia 2026 Addon Agreement Template.docx
```

`compare_docs.py` additionally reads the generated
`WordCamp Asia 2026 Super Admin Sponsorship Agreement Template.docx`, so run
`generate_templates.py` first before comparing.

## Requirements

- Python 3.8+
- python-docx library (`pip install -r requirements.txt`)

## Files

| Script | Run with | Requires reference `.docx`? | What it does |
|--------|----------|------------------------------|--------------|
| `generate_templates.py` | no args | Optional (warns if missing) | Generates all tier templates + addon template (saved to current dir) |
| `check_numbering.py` | no args | **Yes** | Prints the numbering definitions found in the reference doc |
| `compare_docs.py` | no args | **Yes** (+ generated Super Admin template) | Prints a side-by-side of original vs generated paragraph styles/numbering |
| `inspect_doc.py` | no args | **Yes** | Prints the paragraph structure (style, bold, numbering) of the reference doc |
| `reference_doc.py` | imported | — | Helper that locates the 2025 reference `.docx` (not run directly) |
| `requirements.txt` | — | — | `pip install -r` dependency (`python-docx`) |

`check_numbering.py`, `compare_docs.py`, and `inspect_doc.py` are debugging/inspection helpers
used while tuning `generate_templates.py`; they only print to the console and produce no files.

## Notes

- The script uses the original 2025 agreement document (see above) as a reference for formatting
- Generated templates preserve numbering, formatting, and structure from the original
- Tier amounts and benefits are defined in the `tiers` dict in `generate_templates.py`; edit there to change them
- Placeholders like `[SPONSOR_NAME]`, `[DATE]`, `[BOOTH_SIZE_SQM]`, and `[RAFFLE_TIME]` are included in templates for manual filling
- Generated `.docx` files are git-ignored (they are deliverables, not source)

