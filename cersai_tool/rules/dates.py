"""Dates: reading the LMS formats and writing CERSAI's YYYY-MM-DD.

The LMS writes some dates with a 2-digit year (01-JUL-26), so the century has
to be worked out:
  - date of birth: 20YY only if that makes the person at least 18 today,
    otherwise 19YY (11-JAN-51 -> 1951-01-11);
  - any other past date (incorporation, disbursal): 20YY unless that would be
    in the future, otherwise 19YY.
"""
import re
from datetime import date, datetime

from .text import as_text

MONTHS = {m: i for i, m in enumerate(
    ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"], start=1)}

BIRTH = "birth"
PAST = "past"


def iso(d: date) -> str:
    return d.strftime("%Y-%m-%d")


def _safe_date(y, m, d):
    try:
        return date(y, m, d)
    except ValueError:
        return None


def _century(yy: int, month: int, day: int, kind: str, today: date):
    modern = _safe_date(2000 + yy, month, day)
    old = _safe_date(1900 + yy, month, day)
    if modern is None or old is None:
        return None
    if kind == BIRTH:
        return modern if 2000 + yy <= today.year - 18 else old
    return modern if modern <= today else old


def from_lms(value, kind: str = PAST, today: date = None):
    """Reads DD-MM-YYYY, DD-MON-YY, DD-MON-YYYY or a real Excel date.
    Returns a date, or None if the text is not a date."""
    today = today or date.today()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = as_text(value).upper()
    if not text:
        return None
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", text)
    if m:
        return _safe_date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?", text)
    if m:
        return _safe_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.fullmatch(r"(\d{1,2})[- ]([A-Z]{3})[A-Z]*[- ](\d{2}|\d{4})", text)
    if m and m.group(2) in MONTHS:
        day, month, year = int(m.group(1)), MONTHS[m.group(2)], m.group(3)
        if len(year) == 4:
            return _safe_date(int(year), month, day)
        return _century(int(year), month, day, kind, today)
    return None


def to_cersai(value, kind: str = PAST, today: date = None) -> str:
    """Any supported date text -> 'YYYY-MM-DD', or '' if it can't be read."""
    d = from_lms(value, kind, today)
    return iso(d) if d else ""


def check_cersai_date(text: str, today: date = None) -> str:
    """Returns a problem description, or '' if text is a valid CERSAI date
    (YYYY-MM-DD, not in the future, not before 1900)."""
    today = today or date.today()
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", text or "")
    if not m:
        return "not in YYYY-MM-DD format"
    d = _safe_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    if d is None:
        return "not a real date"
    if d > today:
        return "in the future"
    if d.year < 1900:
        return "before 1900"
    return ""
