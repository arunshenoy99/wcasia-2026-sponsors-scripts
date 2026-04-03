"""
Shared utilities for extracting email addresses from spreadsheet cell values.
"""
import re
from typing import List


def extract_emails_from_string(value: str) -> List[str]:
    """
    Smartly extract all email addresses from a string that may contain multiple
    emails separated by comma, semicolon, newline, " and ", or other characters.

    Args:
        value: Cell value (e.g. "a@x.com, b@y.com" or "a@x.com; b@y.com")

    Returns:
        List of email strings (trimmed, no empty strings).
    """
    if value is None or (isinstance(value, float) and str(value) == "nan"):
        return []
    s = str(value).strip()
    if not s:
        return []
    # Normalize separators: semicolon, newline, " and " -> comma
    s = s.replace(";", ",").replace("\n", ",").replace("\r", ",")
    s = re.sub(r"\s+and\s+", ",", s, flags=re.IGNORECASE)
    parts = re.split(r"[,]+", s)
    emails = []
    for part in parts:
        part = part.strip().strip('"\'()[]')
        if part and "@" in part:
            # Simple validation: has local part and domain (something@something)
            if re.search(r".+@.+\..+", part):
                emails.append(part)
    return emails


def extract_emails_as_single_string(value: str, separator: str = ", ") -> str:
    """
    Extract emails from a string and return as a single comma-separated string
    (for use as the "Email" column in round CSV).

    Args:
        value: Cell value.
        separator: String to join multiple emails (default ", ").

    Returns:
        Single string of one or more emails, or empty string if none found.
    """
    emails = extract_emails_from_string(value)
    return separator.join(emails) if emails else ""
