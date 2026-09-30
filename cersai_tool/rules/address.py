"""Reading free-text addresses.

LMS addresses end with ", <City>, <State>- <PIN>, India". The part before that
is free text. These helpers pull out the pieces CERSAI asks for separately
(flat number, floor, survey and plot numbers, building, project, street,
locality). A piece counts as found only when its pattern matches cleanly;
otherwise it is left empty so the title report or a person can supply it.
The same helpers read the property description in the title reports.
"""
import re

from .text import as_text, straighten

TAIL_RE = re.compile(r",\s*([^,]*?),\s*([^,]*?)\s*[-–]\s*(\d{6})\s*,\s*India\s*$", re.I)

# Words that name the unit being mortgaged, and what kind of structure it is
UNIT_KINDS = {
    "flat": "without_land", "apartment": "without_land", "office": "without_land",
    "unit": "without_land", "shop": "without_land", "gala": "without_land",
    "galla": "without_land", "premises": "without_land",
    "house": "with_land", "h": "with_land", "row house": "with_land",
    "bungalow": "with_land", "villa": "with_land", "duplex": "with_land",
    "property": "",   # "Property bearing No. ...": a number, but the structure is not stated
}
_UNIT_WORDS = (r"row\s*house|flat|apartment|office|unit|shop|gall?a|premises|house|h|bungalow|villa|duplex|"
               r"door|property\s+bearing|property")
_NUMBER = r"[A-Z]{0,3}\s?[-/]?\s?\d+[A-Z]?(?:\s?[-/]\s?[A-Z0-9]+)*"
UNIT_RE = re.compile(
    rf"\b({_UNIT_WORDS})\.?\s*(?:bearing\s+)?(?:no\.?s?|nos\.?|number|#)\s*[:.\-]?\s*({_NUMBER})"
    rf"((?:\s*(?:,|&|and|/)\s*(?:no\.?\s*)?(?:{_NUMBER})(?!\s*(?:st|nd|rd|th)\b)(?![\w-]))*)",
    re.I)

FLOOR_WORDS = {"ground": 0, "upper ground": 0, "lower ground": -1, "basement": -1, "stilt": 0,
               "first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5, "sixth": 6,
               "seventh": 7, "eighth": 8, "ninth": 9, "tenth": 10}
FLOOR_NUM_RE = re.compile(r"\b(\d{1,3})\s*(?:st|nd|rd|th)?\s*floor\b|\bfloor\s*no\.?\s*(\d{1,3})\b", re.I)
FLOOR_WORD_RE = re.compile(r"\b(upper ground|lower ground|ground|basement|stilt|first|second|third|"
                           r"fourth|fifth|sixth|seventh|eighth|ninth|tenth)\s+floor\b", re.I)

_NUM_TOKEN = r"\d+[A-Z]?(?:\s*/\s*[A-Z0-9]+)*(?:\s*\((?:pt|p|part)\))?"
SURVEY_KEYS = (r"c\.?\s?t\.?\s?s\.?\s*(?:no\.?s?|nos\.?|number)?|survey\s*(?:no\.?s?|nos\.?|number)?"
               r"|(?<![A-Za-z]\.)(?<![A-Za-z])s\.\s?no\.?|(?<![A-Za-z]\.)s\s?no\.?(?=\s*\d)|sy\.?\s*no\.?"
               r"|r\.?\s?s\.?\s*no\.?|(?<![A-Za-z]\.)c\.?\s?s\.?\s*no\.?"
               r"|gat\s*no\.?|gut\s*no\.?|khasra\s*(?:no\.?)?|kh\.?\s*no\.?|khata\s*no\.?|khewat\s*no\.?"
               r"|mustatil\s*no\.?|killa\s*no\.?|new\s*s\.?\s*no\.?")
SURVEY_RE = re.compile(
    rf"\b(?:{SURVEY_KEYS})\s*[:.\-]?\s*({_NUM_TOKEN}(?:\s*(?:,|&|and)\s*{_NUM_TOKEN}(?![A-Za-z]))*)", re.I)
HISSA_RE = re.compile(r"\bhissa\s*no\.?\s*([0-9A-Z/]+)", re.I)
OLD_SURVEY_RE = re.compile(r"\bold\s*s(?:urvey)?\.?\s*no\.?\s*" + _NUM_TOKEN, re.I)
PLOT_RE = re.compile(
    rf"(?:\b(?:plot|p\.?\s?f|f\.?\s?p|final\s*plot|site)\s*(?:no\.?s?|nos\.?|number)?|(?<![A-Za-z]\.)\bp\.\s?no\.?)"
    rf"\s*[:.\-]?\s*"
    rf"({_NUMBER}(?:\s*(?:,|&|and)\s*{_NUMBER}(?![A-Za-z]))*)", re.I)

PROJECT_WORDS = re.compile(
    r"\b(chsl?|csl|c\.h\.s|co-?op\.?|co-?operative|society|soc|apartments?|apts?|residency|residences?|heights|"
    r"towers?|complex|enclave|arcade|chambers|plaza|mall|park|estate|gardens?|niwas|nivas|sadan|bhavan|"
    r"bhawan|mansion|palace|avenue|homes|township|city|hub|square|court|kunj|vatika|vihar|dham|"
    r"industrial\s+estate|business\s+(?:centre|center|park)|centre|center|building|bldg)\b", re.I)
BUILDING_WORDS = re.compile(r"\b(wing|tower|building|bldg|block)\b", re.I)
STREET_WORDS = re.compile(r"\b(road|rd|marg|street|lane|path|highway|hwy|chowk)\b", re.I)
LANDMARK_START = re.compile(r"^\s*(near|nr|opp|opposite|behind|beside|next to|off|adjacent|above|below)\b", re.I)
LEAD_WORDS = re.compile(r"^\s*(situated\s+(?:at|in|on)|located\s+(?:at|in|on)|lying\s+at|at|on|in)\s+", re.I)
NOT_A_NAME = re.compile(r"\b(flat|shop|office|unit|gala|floor|plot|survey|khasra|cts|house|no\.?|ward|"
                        r"taluka?|tal|tehsil|district|dist|division|pin|measuring|admeasuring|area|sq|mtrs?|"
                        r"yds?|ft|feet|hobli|mandal|circle|municipal|corporation|property|bearing|"
                        r"portion|part|roof|rights|india)\b", re.I)


BOUNDARIES_RE = re.compile(r"\b(?:butted\s+and\s+)?bounded\b|\bboundar(?:y|ies)\b|\bfour\s+sides\b", re.I)


def without_boundaries(text: str) -> str:
    """The text before a description of the property's boundaries ("bounded on
    the West by Plot No. 16 ..."): the numbers there belong to the neighbours."""
    m = BOUNDARIES_RE.search(text or "")
    return (text or "")[: m.start()] if m else (text or "")


def split_tail(address: str):
    """'..., Goregaon (E), Mumbai 400063, Mumbai, Maharashtra- 400063, India'
    -> (body, 'Mumbai', 'Maharashtra', '400063'). Without the usual ending the
    whole text is the body and the other parts are empty."""
    text = straighten(as_text(address))
    m = TAIL_RE.search(text)
    if not m:
        return text.strip(" ,"), "", "", ""
    return text[: m.start()].strip(" ,"), m.group(1).strip(), m.group(2).strip(), m.group(3)


def segments(body: str) -> list:
    return [s.strip() for s in re.split(r",|;", body) if s.strip()]


def _compact_number(text: str) -> str:
    text = re.sub(r"\s*([-/])\s*", r"\1", text.strip())
    text = re.sub(r"\((?:pt|p|part)\)", "pt", text, flags=re.I)
    return re.sub(r"\s+", "", text).upper()


def find_unit(text: str):
    """The mortgaged unit: ('flat', ['1301', '1302']) from 'Flat No. 1301 & 1302'.
    Returns (None, []) when no unit keyword with a number is found. A number
    marked "Old" ("Old Property No. 121, New Property No. 1777") is used only
    when there is no other."""
    matches = list(UNIT_RE.finditer(text or ""))
    if not matches:
        return None, []
    current = [m for m in matches if not re.search(r"\bold\s*$", text[: m.start()], re.I)]
    m = (current or matches)[0]
    kind = re.sub(r"\s+", " ", m.group(1).lower())
    if kind == "h":
        kind = "house"
    numbers = [_compact_number(m.group(2))]
    for extra in re.findall(rf"(?:,|&|and|/)\s*(?:no\.?\s*)?({_NUMBER})", m.group(3) or "", re.I):
        numbers.append(_compact_number(extra))
    return kind, [n for n in numbers if n]


def unit_structure(kind: str) -> str:
    """'without_land', 'with_land' or '' for a unit keyword."""
    return UNIT_KINDS.get((kind or "").lower(), "")


_FLOOR_NAME = r"(?:upper ground|lower ground|ground|basement|stilt|first|second|third|fourth|fifth|\d{1,3}(?:st|nd|rd|th))"
SEVERAL_FLOORS_RE = re.compile(rf"\b{_FLOOR_NAME}\s+(?:and|&|to)\s+{_FLOOR_NAME}\s+floors?\b", re.I)


def find_floor(text: str):
    """'11th Floor' -> 11, 'Ground Floor' -> 0. None if absent or if the
    address names more than one floor ("Ground and First Floor")."""
    if SEVERAL_FLOORS_RE.search(text or ""):
        return None
    found = set()
    for m in FLOOR_NUM_RE.finditer(text or ""):
        found.add(int(m.group(1) or m.group(2)))
    for m in FLOOR_WORD_RE.finditer(text or ""):
        found.add(FLOOR_WORDS[m.group(1).lower()])
    if len(found) == 1:
        value = found.pop()
        return value if -3 <= value <= 100 else None
    return None


def find_survey_numbers(text: str) -> list:
    """All survey-type numbers (CTS, survey, S. No., Sy. No., khasra, gat...).
    When both old and new survey numbers are given, only the new ones are kept,
    with the Hissa number appended: 'Old S. No. 260(Pt) and New S. No. 12 and
    Hissa No. 2' -> ['12/2']."""
    text = text or ""
    if re.search(r"\bnew\s*s(?:urvey)?\.?\s*no", text, re.I):
        text = OLD_SURVEY_RE.sub(" ", text)
    numbers = []
    for m in SURVEY_RE.finditer(text):
        for part in re.split(r"\s*(?:,|&|\band\b)\s*", m.group(1)):
            part = _compact_number(part)
            if part and part not in numbers:
                numbers.append(part)
    hissa = HISSA_RE.search(text)
    if hissa and len(numbers) == 1 and "/" not in numbers[0]:
        numbers[0] = f"{numbers[0]}/{_compact_number(hissa.group(1))}"
    return numbers


def find_plot_numbers(text: str) -> list:
    numbers = []
    for m in PLOT_RE.finditer(text or ""):
        for part in re.split(r"\s*(?:,|&|\band\b)\s*", m.group(1)):
            part = _compact_number(part)
            if part and part not in numbers:
                numbers.append(part)
    return numbers


def _name_segment(seg: str) -> bool:
    return not LANDMARK_START.search(seg) and len(re.sub(r"[^A-Za-z]", "", seg)) >= 3


BUILDING_ID = re.compile(r"^\s*(wing|tower|building|bldg|block)\s*(no\.?)?\s*[-:]?\s*[A-Z0-9][A-Z0-9 -]{0,6}$",
                         re.I)
# words that only ever name a building or society (never a locality)
STRONG_PROJECT_WORDS = re.compile(
    r"\b(chsl?|csl|c\.h\.s|co-?op\.?|co-?operative|society|soc|apartments?|apts?|residency|heights|towers?|"
    r"complex|arcade|chambers|plaza|mall|building|bldg)\b", re.I)


def find_project(parts: list) -> str:
    """The first segment that names a society, project or complex and is not a
    unit, floor, landmark, survey description or a bare building/wing number
    ("Tower C-9" is a building, not a project)."""
    for seg in parts:
        seg = LEAD_WORDS.sub("", seg).strip()
        if PROJECT_WORDS.search(seg) and _name_segment(seg) and not UNIT_RE.search(seg) \
                and not FLOOR_NUM_RE.search(seg) and not SURVEY_RE.search(seg) \
                and not STREET_WORDS.search(seg) and not BUILDING_ID.match(seg):
            return re.sub(r"^(in|at|of)?\s*(the\s+)?(society|building|project)?\s*(known as|named)\s+", "",
                          seg, flags=re.I).strip(" '\"")
    return ""


def find_building(parts: list, project: str = "") -> str:
    for seg in parts:
        seg = LEAD_WORDS.sub("", seg).strip()
        if seg == project:
            continue
        if BUILDING_WORDS.search(seg) and _name_segment(seg) and not UNIT_RE.search(seg) \
                and not FLOOR_NUM_RE.search(seg) and not STREET_WORDS.search(seg):
            return seg
    return ""


def find_street(parts: list) -> str:
    for seg in parts:
        if STREET_WORDS.search(seg) and not LANDMARK_START.search(seg) and not UNIT_RE.search(seg) \
                and not NOT_A_NAME.search(LEAD_WORDS.sub("", seg).split(" road")[0]):
            return LEAD_WORDS.sub("", seg).strip()
    return ""


def _place_key(text: str) -> str:
    return re.sub(r"[^a-z]", "", (text or "").lower())


def find_locality(parts: list, lms_city: str = "", lms_state: str = "") -> str:
    """The area name just before the city/state ending: the last segment that is
    a place name. Segments that only repeat the city or state, PINs, landmarks,
    survey/plot/area descriptions and administrative words are skipped."""
    skip = [k for k in (_place_key(lms_city), _place_key(lms_state)) if k]
    for seg in reversed(parts):
        plain = LEAD_WORDS.sub("", re.sub(r"\b\d{6}\b", "", seg)).strip(" -.")
        plain = re.sub(r"^(known\s+as|named)\s+", "", plain, flags=re.I)
        plain = re.sub(r"^(village|vill|mouje|moje|mauja)\s*[-.:]?\s*", "Village ", plain, flags=re.I)
        key = _place_key(plain)
        # skip a segment that only repeats the city or state ("Delhi" when the
        # city is "WEST DELHI"); a locality that merely ends with the city name
        # is kept, with the city name trimmed off below
        if not key or any(key == s or (len(key) >= 4 and key in s) for s in skip):
            continue
        if re.fullmatch(r"[\d\s/-]*", plain) or NOT_A_NAME.search(plain) or LANDMARK_START.search(plain):
            continue
        if STREET_WORDS.search(plain) or STRONG_PROJECT_WORDS.search(plain) or BUILDING_ID.match(plain) \
                or len(key) < 3:
            continue
        for s in (lms_city, lms_state):
            if s and plain.lower().endswith(" " + s.lower()):
                plain = plain[: -len(s)].strip(" -")
        return plain
    return ""


def parse_property_address(address: str) -> dict:
    """Everything the IMM row can take from an LMS property address."""
    body, city, state, pin = split_tail(address)
    core = without_boundaries(body)
    parts = segments(core)
    kind, numbers = find_unit(core)
    project = find_project(parts)
    return {
        "body": body, "lms_city": city, "lms_state": state, "tail_pin": pin,
        "unit_kind": kind, "unit_numbers": numbers,
        "floor": find_floor(core),
        "survey_numbers": find_survey_numbers(core),
        "plot_numbers": find_plot_numbers(core),
        "project": project,
        "building": find_building(parts, project),
        "street": find_street(parts),
        "locality": find_locality(parts, city, state),
    }


def split_lines(text: str, max_len: int = 100, max_lines: int = 3) -> list:
    """Packs comma-separated pieces into lines of at most max_len characters,
    breaking a piece at a word boundary only when it is too long on its own.
    Returns at most max_lines lines when the text fits; when it does not fit,
    all lines are returned so the caller can report the overflow."""
    pieces = []
    for piece in segments(text):
        words, current = piece.split(), ""
        for word in words:
            if len(current) + 1 + len(word) > max_len and current:
                pieces.append(current)
                current = word
            else:
                current = f"{current} {word}".strip()
        if current:
            pieces.append(current)
    lines, current = [], ""
    for piece in pieces:
        candidate = f"{current} {piece}".strip()
        if len(candidate) <= max_len:
            current = candidate
        else:
            lines.append(current)
            current = piece
    if current:
        lines.append(current)
    return lines[:max_lines] if len(lines) <= max_lines else lines
