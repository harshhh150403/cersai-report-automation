"""Text cleaning shared by every row.

CERSAI text fields allow letters, digits and a few special characters (which
ones depends on the field), never two special characters side by side (a space
counts as one), and no special character at the start or end. The helpers
here turn any source text into a value that follows those rules.
"""
import math
import re
from datetime import date, datetime

# Special characters each kind of field may contain (a space is always allowed)
ADDRESS = "-_./"            # address lines, locality, pocket, plot number
DESCRIPTION = "-_.,/"       # asset description, "others" descriptions, narration
ORG_NAME = "-_."            # non-individual borrower name, project, building
STREET = "-/"               # street name
HOUSE = "-/"                # house / flat / unit number
SURVEY = "-/^"              # survey numbers, several separated by ^

_SPECIALS = "-_./,^'"


def as_text(value) -> str:
    """Turns any spreadsheet value into plain text: empty for blanks, whole
    numbers without ".0", dates as YYYY-MM-DD, single spaces."""
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return f"{value:.10f}".rstrip("0").rstrip(".")
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    text = str(value)
    for ch in ("\xa0", "\t", "\r", "\n", " ", " ", "​"):
        text = text.replace(ch, " ")
    text = re.sub(r"\s+", " ", text).strip()
    if text.lower() in ("nan", "none", "null"):
        return ""
    return text


def straighten(text: str) -> str:
    """Curly quotes and dashes to plain ones; rejoins words split by a line-end
    hyphen ("Co- op" -> "Co-op")."""
    for a, b in (("‘", "'"), ("’", "'"), ("“", '"'), ("”", '"'),
                 ("–", "-"), ("—", "-"), ("…", "..."), ("\xad", "-")):
        text = text.replace(a, b)
    return re.sub(r"([A-Za-z])- ([a-z])", r"\1-\2", text)


def collapse_special_runs(text: str) -> str:
    """Replaces two or more special characters in a row (spaces included) with
    one: the first non-space one if there is one, otherwise a space."""
    def pick(match):
        run = match.group(0)
        for ch in run:
            if ch != " ":
                return ch
        return " "
    return re.sub(r"[^A-Za-z0-9]{2,}", pick, text)


_DANGLING = {"of", "at", "in", "on", "near", "and", "the", "to", "opp", "behind", "beside", "nr", "with"}


def cut_at_word(text: str, max_len: int) -> str:
    """Shortens text to max_len characters at a word boundary, never mid-word
    and never by deleting spaces. A cut text does not end with a joining word
    such as "of" or "near"."""
    if len(text) <= max_len:
        return text
    head = text[: max_len + 1]
    space = head.rfind(" ")
    head = head[:space] if space > 0 else text[:max_len]
    words = head.rstrip(" " + _SPECIALS).split(" ")
    while len(words) > 1 and words[-1].lower() in _DANGLING:
        words.pop()
    return " ".join(words).rstrip(" " + _SPECIALS)


def clean(value, specials: str, max_len: int, *, letters_only=False,
          ampersand_to_and=False, upper=False) -> str:
    """The standard cleaning for a CERSAI text field.

    1. straighten quotes, rejoin broken words, turn '&' into 'AND' (names);
    2. drop dots, except a dot between two digits when the field allows dots;
    3. replace every character the field does not allow with a space;
    4. collapse runs of special characters and tighten ' / ' to '/';
    5. strip special characters from both ends and cut at a word boundary.
    """
    text = straighten(as_text(value))
    if ampersand_to_and:
        text = text.replace("&", " AND ")
    if "." in specials:
        text = re.sub(r"(?<!\d)\.|\.(?!\d)", " ", text)
    else:
        text = text.replace(".", " ")
    out = []
    for ch in text:
        if ch.isascii() and ch.isalpha():
            out.append(ch)
        elif ch.isascii() and ch.isdigit():
            out.append(" " if letters_only else ch)
        elif ch == " " or (ch in specials and not letters_only):
            out.append(ch)
        else:
            out.append(" ")
    text = re.sub(r"\s+", " ", "".join(out)).strip()
    text = re.sub(r"\s*([\-_./,^])\s*", r"\1", text)
    text = collapse_special_runs(text)
    text = text.strip(" " + _SPECIALS)
    if upper:
        text = text.upper()
    return cut_at_word(text, max_len)


def break_rules(value: str, specials: str, max_len: int, letters_only=False) -> list:
    """Lists the CERSAI character rules a value breaks (empty list = fine).
    Used by the checks before the .DAT file is written."""
    problems = []
    if len(value) > max_len:
        problems.append(f"longer than {max_len} characters")
    allowed_specials = set(specials) | {" "}
    for ch in value:
        if ch.isascii() and ch.isalpha():
            continue
        if ch.isascii() and ch.isdigit() and not letters_only:
            continue
        if ch in allowed_specials and not (letters_only and ch != " "):
            continue
        problems.append(f"character '{ch}' is not allowed")
        break
    if re.search(r"[^A-Za-z0-9]{2,}", value):
        problems.append("two special characters (or spaces) side by side")
    if value and (not value[0].isalnum() or not value[-1].isalnum()):
        problems.append("starts or ends with a special character")
    return problems


def normalise_for_compare(text: str) -> str:
    """Upper case, letters and digits only, single spaces: for comparing values
    from different sources."""
    text = straighten(as_text(text)).upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()
