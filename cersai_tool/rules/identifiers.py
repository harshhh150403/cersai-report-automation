"""PAN, CKYC, CIN, LLPIN, mobile, email and PIN: cleaning and format rules."""
import re

from .text import as_text

PAN_RE = re.compile(r"^[A-Z]{5}[0-9]{4}[A-Z]$")
CKYC_RE = re.compile(r"^[0-9]{14}$")
CIN_RE = re.compile(r"^[A-Z][0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6}$")
LLPIN_RE = re.compile(r"^[A-Z0-9]{7}$")
MOBILE_RE = re.compile(r"^[0-9]{10}$")
EMAIL_RE = re.compile(r"^[A-Za-z0-9_.]+@[A-Za-z0-9]+(\.[A-Za-z0-9]+)+$")
PIN_RE = re.compile(r"^[1-9][0-9]{5}$")

# 4th letter of a PAN tells the holder type (a mismatch is only a warning)
PAN_TYPE_LETTER = {"IND": "P", "SPR": "P", "PRF": "F", "LLP": "F", "COM": "C",
                   "HUF": "H", "TRS": "T", "AOP": "A", "BOI": "B", "GOV": "G"}


def clean_pan(value) -> str:
    text = re.sub(r"\s+", "", as_text(value)).upper()
    return text if PAN_RE.match(text) else ""


def clean_ckyc(value) -> str:
    text = re.sub(r"\D", "", as_text(value))
    return text if CKYC_RE.match(text) else ""


def clean_cin(value) -> str:
    text = re.sub(r"\s+", "", as_text(value)).upper()
    return text if CIN_RE.match(text) else ""


def clean_llpin(value) -> str:
    """LLPINs are written like AAB-1234; CERSAI wants 7 characters, so the
    hyphen is removed (open question Q16)."""
    text = re.sub(r"[\s\-]+", "", as_text(value)).upper()
    return text if LLPIN_RE.match(text) else ""


def clean_mobile(value) -> str:
    """'+91--9820302743' -> '9820302743'. Anything that is not 10 digits after
    removing the country code becomes ''."""
    digits = re.sub(r"\D", "", as_text(value))
    if len(digits) == 12 and digits.startswith("91"):
        digits = digits[2:]
    if len(digits) == 11 and digits.startswith("0"):
        digits = digits[1:]
    return digits if MOBILE_RE.match(digits) else ""


def clean_email(value) -> str:
    """Trims spaces; returns '' if the email breaks CERSAI's rule (letters,
    digits, '_' and '.', one '@', a dot after the '@', max 50 characters)."""
    text = as_text(value).replace(" ", "")
    if len(text) > 50 or not EMAIL_RE.match(text):
        return ""
    return text


def clean_pin(value) -> str:
    text = re.sub(r"\s+", "", as_text(value))
    return text if PIN_RE.match(text) else ""
