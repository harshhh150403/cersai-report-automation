"""Reading the advocates' title reports (legal PDFs).

File names: <Application number>_<Collateral ID>.pdf, one PDF per collateral.

The reports follow ARGFL's standard template: a header (Status, dates), a box
with "Loan applicant" and "Property Owner", and a numbered table. Items are
found by their label, not their number (one law firm's reports have no item 6).
Words that are struck through are ignored: the text layer still contains them,
so strike lines drawn across words are detected and those words are dropped.
"""
import re
from pathlib import Path

import pdfplumber
import pymupdf

from ..rules import address as addr
from ..rules.owners import split_owner_list
from ..rules.text import as_text, straighten

NAME_RE = re.compile(r"^(APPL\d{8})_(COLL\d{8})\.pdf$", re.I)

# item labels (lower case, spaces collapsed) -> item number
ITEM_LABELS = [
    ("full description", 1), ("nature of immovable", 2), ("type of transaction", 3),
    ("name of the owner", 4), ("list of documents", 5), ("flow of title", 6), ("land history", 6),
    ("title deeds", 7), ("nature of mortgage", 8), ("non- encumbrance", 9), ("non-encumbrance", 9),
    ("search", 9), ("views on encumbrance", 10), ("duly stamped", 11), ("certificate of title", 12),
    ("any other suggestion", 13),
]


def index_folder(folder: Path, active_ids: dict):
    """Finds the title reports for this batch.

    active_ids: {collateral_id: application_number} of the batch.
    Returns (found, problems): found = {collateral_id: path}; problems = list of
    (file name, message) for names that break the convention or point to the
    wrong application. PDFs for collaterals that are not in this batch (older
    days) are ignored silently."""
    found, problems = {}, []
    if not folder.exists():
        return found, [("", f"Legal folder not found: {folder}")]
    batch_apps = set(active_ids.values())
    for path in sorted(folder.glob("*")):
        if path.suffix.lower() != ".pdf":
            continue
        m = NAME_RE.match(path.name)
        if not m:
            head = re.match(r"(APPL\d+)", path.name, re.I)
            if head and head.group(1).upper() in batch_apps:
                problems.append((path.name, "file name does not follow <APPL + 8 digits>_<COLL + 8 digits>.pdf"))
            elif not head:
                problems.append((path.name, "file name does not follow <APPL + 8 digits>_<COLL + 8 digits>.pdf"))
            continue
        app, coll = m.group(1).upper(), m.group(2).upper()
        if coll in active_ids:
            if active_ids[coll] != app:
                problems.append((path.name, f"collateral {coll} belongs to {active_ids[coll]}, not {app}: misfiled"))
            elif coll in found:
                problems.append((path.name, f"a second PDF for {coll}; only {found[coll].name} is used"))
            else:
                found[coll] = path
        elif app in batch_apps:
            problems.append((path.name, f"no active collateral {coll} for {app} in this batch"))
    return found, problems


# --------------------------------------------------------------------------
# Low-level PDF reading
# --------------------------------------------------------------------------
def _struck_word_boxes(page) -> list:
    """Bounding boxes of words with a horizontal line through their middle."""
    lines = []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] == "l":
                a, b = item[1], item[2]
                if abs(a.y - b.y) < 0.8 and abs(b.x - a.x) > 4:
                    lines.append((min(a.x, b.x), max(a.x, b.x), (a.y + b.y) / 2))
            elif item[0] == "re":
                r = item[1]
                if r.height < 1.6 and r.width > 4:
                    lines.append((r.x0, r.x1, (r.y0 + r.y1) / 2))
    struck = []
    for w in page.get_text("words"):
        x0, y0, x1, y1 = w[:4]
        h = y1 - y0
        for lx0, lx1, ly in lines:
            if lx0 <= x0 + 1 and lx1 >= x1 - 1 and (y0 + 0.30 * h) < ly < (y1 - 0.25 * h):
                struck.append((x0, y0, x1, y1))
                break
    return struck


def _words_in(words, box, struck) -> str:
    x0, top, x1, bottom = box
    picked = []
    for w in words:
        cx, cy = (w[0] + w[2]) / 2, (w[1] + w[3]) / 2
        if x0 <= cx <= x1 and top <= cy <= bottom and tuple(w[:4]) not in struck:
            picked.append(w)
    picked.sort(key=lambda w: (round(w[3] / 3), w[0]))
    return " ".join(w[4] for w in picked)


def _item_number(label: str, number_text: str):
    label = re.sub(r"\s+", " ", label.lower())
    for key, number in ITEM_LABELS:
        if key in label:
            return number
    digits = re.match(r"^\s*(\d{1,2})\.?\s*$", number_text or "")
    return int(digits.group(1)) if digits else None


def read_pdf(path: Path) -> dict:
    """Low-level read: {'items': {n: text}, 'header': text, 'struck': bool, 'text_chars': n}."""
    doc = pymupdf.open(str(path))
    text_chars = sum(len(p.get_text().strip()) for p in doc)
    result = {"items": {}, "header": "", "struck_text": False, "text_chars": text_chars, "pages": doc.page_count}
    if text_chars < 200:
        return result
    page_words, page_struck = [], []
    for page in doc:
        words = [tuple(w[:5]) for w in page.get_text("words")]
        struck = set(_struck_word_boxes(page))
        page_words.append(words)
        page_struck.append(struck)
        if struck:
            result["struck_text"] = True
    # header: page 1 words above the first table, without struck words
    items, current = {}, None
    with pdfplumber.open(str(path)) as pdf:
        first_table_top = None
        for pno, page in enumerate(pdf.pages):
            for table in page.find_tables():
                if pno == 0 and (first_table_top is None or table.bbox[1] < first_table_top):
                    first_table_top = table.bbox[1]
                for row in table.rows:
                    cells = [c for c in row.cells]
                    if len(cells) != 3 or any(c is None for c in cells):
                        continue
                    number, label, value = (_words_in(page_words[pno], c, page_struck[pno]) for c in cells)
                    item = None
                    if number.strip() or label.strip():
                        item = _item_number(label, number)
                    if item is not None and item != current:
                        current = item
                        items.setdefault(item, {"label": label, "value": value})
                        if items[item]["value"] != value:
                            items[item]["value"] = (items[item]["value"] + " " + value).strip()
                    elif current is not None:
                        items[current]["label"] = (items[current]["label"] + " " + label).strip()
                        items[current]["value"] = (items[current]["value"] + " " + value).strip()
        header_bottom = first_table_top or 842
        result["header"] = _words_in(page_words[0], (0, 0, 10000, header_bottom), page_struck[0])
    result["items"] = {k: straighten(v["value"]) for k, v in items.items()}
    return result


# --------------------------------------------------------------------------
# From raw text to the values the rows need
# --------------------------------------------------------------------------
AREA_RE = re.compile(
    r"(?:adm\.?|admeasuring|area\s+(?:of|measuring))\s*(?:about\s*)?([\d,]+(?:\.\d+)?)\s*"
    r"(sq\.?\s*m(?:trs?|tr|eters?|etres?)?\.?|sq\.?\s*f(?:ee)?t?s?\.?|sq\.?\s*y(?:ar)?ds?\.?|"
    r"acres?|gunthas?|guntas?|hectares?|ares?|cents?)\s*,?\s*"
    r"((?:rera\s+)?(?:super\s*)?(?:carpet|built\s*-?\s*up|plot)\s*area)?", re.I)
UNIT_IDS = [(r"^sq\.?\s*m", "7"), (r"^sq\.?\s*f", "6"), (r"^sq\.?\s*y", "8"), (r"^acre", "1"),
            (r"^gunt", "4"), (r"^hect", "5"), (r"^are", "2"), (r"^cent", "3")]
MORTGAGE_TYPES = [(r"equitable|deposit\s+of\s+title\s+deeds", "17"), (r"registered|simple|english", "18"),
                  (r"agreement\s+to\s+create", "19"), (r"conditional\s+sale", "20")]


def _parse_description(text: str) -> dict:
    """Item 1 'Full Description of the property' -> separate pieces. Numbers
    are read only before any description of the boundaries (which name the
    neighbouring plots)."""
    out = {}
    full, text = text, addr.without_boundaries(text)
    kind, numbers = addr.find_unit(text)
    out["unit_kind"], out["unit_numbers"] = kind, numbers
    out["floor"] = addr.find_floor(text)
    m = AREA_RE.search(text)
    if m:
        out["area"] = m.group(1).replace(",", "")
        unit = m.group(2).lower()
        out["area_unit"] = next((uid for pattern, uid in UNIT_IDS if re.match(pattern, unit)), "")
        out["area_type"] = (m.group(3) or "").strip()
    land = re.search(r"land\s+bearing\s+(.+?)(?:\s+situated|\s+of\s+village|\s+lying|,\s*village|$)", text, re.I)
    land_text = land.group(1) if land else text
    out["survey_numbers"] = addr.find_survey_numbers(land_text)
    out["plot_numbers"] = addr.find_plot_numbers(land_text)
    project = re.search(r"project\s+known\s+as\s+['\"]?([^',\"]+?)['\"]?(?:,|\s+constructed|\s+on\s+the|$)", text, re.I)
    society = re.search(r"society\s+known\s+as\s+['\"]?([^'\"]+?)['\"]?(?:,?\s+constructed|,\s+situated|$)", text, re.I)
    if project:
        out["project"] = project.group(1).strip()
    elif society:
        out["project"] = society.group(1).strip()
    building = re.search(r"building\s+no\.?\s*(\w+)\s+known\s+as\s+['\"]?([^'\",]+)", text, re.I)
    tower = re.search(r"(tower\s*-?\s*\w+)\s*\(([^)]+)\).*?building\s+known\s+as\s+([^,]+?)\s+of\s+the\s+project",
                      text, re.I)
    if building:
        out["building"] = f"{building.group(2).strip()} Building No {building.group(1)}"
    elif tower:
        out["building"] = f"{tower.group(3).strip()} {tower.group(1).strip()} {tower.group(2).strip()}"
    village = re.search(r"\b(?i:village)\s+([A-Z][A-Za-z]+)", text)
    if village:
        out["village"] = village.group(1)
    district = re.search(r"\b(?i:district)\s+([A-Z][A-Za-z]+)|\b(?i:dist)\.?\s+([A-Z][A-Za-z]+)", text)
    if district:
        out["district_name"] = district.group(1) or district.group(2)
    road = re.search(r"((?:[A-Z][A-Za-z]*\s+){1,4}(?:Road|Marg|Street|Lane|Path|Highway)(?:\s+No\.?\s*\d+)?)", text)
    if road:
        out["street"] = road.group(1).strip()
        after = re.search(re.escape(road.group(1)) + r",\s*([A-Z][A-Za-z ]+?)\s*,", text)
        if after:
            out["locality"] = after.group(1).strip()
    if "locality" not in out and "village" in out:
        out["locality"] = f"Village {out['village']}"
    pin = re.findall(r"\b([1-9]\d{2})\s?(\d{3})\b", full)
    if pin:
        out["pin"] = "".join(pin[-1])
    return out


def _owners_from_certificate(text: str) -> list:
    m = re.search(r"certify\s+that\s+(.+?)\s+(?:is|are|were|was)\s+the\s+owner", text or "", re.I)
    return split_owner_list(m.group(1)) if m else []


def read_title_report(path: Path) -> dict:
    """Everything the tool uses from one title report, as a plain dictionary
    (saved with the batch). 'status' is 'ok' or 'no_text'."""
    raw = read_pdf(path)
    report = {"file": path.name, "status": "ok", "pages": raw["pages"], "struck_text": raw["struck_text"]}
    if raw["text_chars"] < 200:
        report["status"] = "no_text"
        return report
    header = as_text(raw["header"])
    items = raw["items"]
    m = re.search(r"Status\s*:?\s*(.+?)\s+Date of Initiation", header, re.I)
    report["report_status"] = m.group(1).strip() if m else ""
    m = re.search(r"Loan applicant\s*[–\-:]+\s*(.+?)\s+Property Owner", header, re.I)
    report["loan_applicant"] = m.group(1).strip(" .") if m else ""
    m = re.search(r"Property Owner\s*[–\-:]+\s*(.+?)$", header, re.I)
    report["owners_header"] = split_owner_list(m.group(1)) if m else []
    m = re.search(r"Date of Issuance\s*[–\-:]+\s*([\d/;\s]+\d)", header, re.I)
    report["date_of_issuance"] = m.group(1).strip() if m else ""

    report["description"] = items.get(1, "")
    report["nature"] = items.get(2, "").strip(" .")
    report["transaction"] = items.get(3, "").strip(" .")
    report["owners_item4"] = split_owner_list(items.get(4, ""))
    report["mortgage_text"] = re.sub(r"\s+", " ", items.get(8, "")).strip(" .")
    report["owners_item12"] = _owners_from_certificate(items.get(12, ""))
    report["items_found"] = sorted(items)
    report["details"] = _parse_description(report["description"])
    return report


def mortgage_type_id(text: str) -> str:
    """'Equitable Mortgage' -> '17'. Unknown wording -> '21' (Other)."""
    for pattern, code in MORTGAGE_TYPES:
        if re.search(pattern, text or "", re.I):
            return code
    return "21" if text else ""
