"""Locate the 2025 reference sponsorship agreement used as a formatting source.

`generate_templates.py` copies the Word numbering definitions out of a signed
2025 *Super Admin* sponsorship agreement so the generated 2026 templates keep the
same numbering. That source file is not committed (it contains sponsor data), so
it is discovered at runtime instead of being hard-coded. See README.md.

Resolution order (within this folder):
  1. A file named exactly `2025-super-admin-sponsorship-agreement.docx`.
  2. Otherwise, any `*Super Admin*Sponsorship Agreement*.docx` that is not a
     generated 2026 template.
"""
from pathlib import Path
from typing import Optional

# Preferred filename — rename your 2025 reference agreement to this for an exact match.
DEFAULT_REFERENCE_NAME = "2025-super-admin-sponsorship-agreement.docx"
# Fallback pattern: any Super Admin sponsorship agreement .docx in this folder.
REFERENCE_GLOB = "*Super Admin*Sponsorship Agreement*.docx"


def find_reference_doc(base_dir: Optional[str] = None) -> Optional[Path]:
    """Return the Path to the reference agreement, or None if none is present."""
    base = Path(base_dir) if base_dir else Path(__file__).resolve().parent
    explicit = base / DEFAULT_REFERENCE_NAME
    if explicit.is_file():
        return explicit
    for path in sorted(base.glob(REFERENCE_GLOB)):
        # Skip the generated 2026 templates this project produces.
        if "Template" in path.name or "2026" in path.name:
            continue
        return path
    return None
