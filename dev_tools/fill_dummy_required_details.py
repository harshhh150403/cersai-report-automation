"""Types the dummy answers into a dummy batch's required-details file.

    python dev_tools/fill_dummy_required_details.py [batch name]

For the dummy data only. Step 1 (run1 --dummy) asks for the values the
hand-made test cases leave out of the extracts on purpose (a gender CKYC
leaves blank, a PAN, a registration number...) and puts questions in the
Check and correct sheet. make_dummy_data.py writes what a person would type in
"Dummy input data/Dummy required details answers.xlsx"; this script types it
into "2 Required details_<batch>.xlsx" of the latest dummy batch (or the
batch named). Then run step 2 and step 3 with --dummy to get a complete .DAT
file.

Nothing else is changed: only the cells the answers name, and the "Your
value" column of the Check and correct sheet. Needed (yellow) cells that are
still empty afterwards are listed.
"""
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from cersai_tool import batch as batches          # noqa: E402  (after the path is set)
from cersai_tool.settings import make_paths       # noqa: E402
from cersai_tool.sources.required_details import BORROWER_KEYS, COLLATERAL_KEYS     # noqa: E402

ANSWERS = ROOT / "Dummy input data" / "Dummy required details answers.xlsx"
NEEDED = "FFFFFF00"                                # the fill of a needed (yellow) cell


def answer_rows(book, sheet_name):
    """The rows of one sheet of the answers file as dicts, without empty values."""
    rows = list(book[sheet_name].iter_rows(values_only=True))
    header = rows[0]
    return [{k: str(v).strip() for k, v in zip(header, row) if v not in (None, "")} for row in rows[1:]]


def fill_sheet(sheet, answers, key_columns, fields, first_row):
    """Types each answer into the row whose key columns match it (only the
    input fields, never the identifying columns). Returns (values typed,
    answers with no row in this sheet)."""
    header = [c.value for c in sheet[1]]
    column = {name: i for i, name in enumerate(header) if name}
    rows = {}
    for row in sheet.iter_rows(min_row=first_row):
        key = tuple(str(row[column[k]].value or "").strip() for k in key_columns)
        if all(key):
            rows[key] = row
    typed, unused = 0, 0
    for answer in answers:
        row = rows.get(tuple(answer.get(k, "") for k in key_columns))
        if row is None:
            unused += 1                            # nothing is needed for this borrower or collateral
            continue
        for field, value in answer.items():
            if field not in fields or field not in column:
                continue
            row[column[field]].value = value
            typed += 1
    return typed, unused


def fill_checks(sheet, answers):
    """Types the answers in the Check and correct sheet: in the row the tool
    made for that collateral and field, or else in an empty row at the bottom."""
    header = [c.value for c in sheet[1]]
    column = {name: i + 1 for i, name in enumerate(header) if name}
    typed = 0
    for answer in answers:
        target = None
        for r in range(2, sheet.max_row + 1):
            if (sheet.cell(r, column["Collateral ID"]).value == answer["Collateral ID"]
                    and sheet.cell(r, column["Field"]).value == answer["Field"]):
                target = r
                break
        if target is None:
            target = next((r for r in range(2, sheet.max_row + 2)
                           if not sheet.cell(r, column["Collateral ID"]).value), sheet.max_row + 1)
            sheet.cell(target, column["Collateral ID"]).value = answer["Collateral ID"]
            sheet.cell(target, column["Row"]).value = answer["Row"]
            sheet.cell(target, column["Field"]).value = answer["Field"]
        sheet.cell(target, column["Your value"]).value = answer["Your value"]
        typed += 1
    return typed


def still_needed(book):
    """Needed (yellow) cells that are still empty: (sheet, row label, field).
    A borrower needs a PAN or a CKYC number: either one typed is enough."""
    left = []
    for name in ("Borrowers", "Collaterals"):
        sheet = book[name]
        header = [c.value for c in sheet[1]]
        for row in sheet.iter_rows(min_row=3):
            typed = {header[i] for i, cell in enumerate(row) if cell.value not in (None, "")}
            for i, cell in enumerate(row):
                alternative = {"PAN": "CKYC", "CKYC": "PAN"}.get(header[i]) if name == "Borrowers" else None
                if alternative in typed:
                    continue
                if cell.fill.fgColor.rgb == NEEDED and cell.value in (None, ""):
                    label = " / ".join(str(row[k].value) for k in ((0, 1) if name == "Borrowers" else (0,)))
                    left.append((name, label, header[i]))
    return left


def main(argv):
    paths = make_paths(None, dummy=True)
    if not ANSWERS.exists():
        print(f"{ANSWERS} is missing: run dev_tools/make_dummy_data.py first.")
        return 1
    folder = batches.pick_batch(paths.output_folder, argv[0] if argv else None, "required", True, print)
    path = batches.batch_files(folder)["required"]
    answers = load_workbook(ANSWERS, read_only=True)
    book = load_workbook(path)
    people, people_unused = fill_sheet(book["Borrowers"], answer_rows(answers, "Borrowers"),
                                       ("Application number", "Neo CIF ID"), BORROWER_KEYS, 3)
    places, places_unused = fill_sheet(book["Collaterals"], answer_rows(answers, "Collaterals"),
                                       ("Collateral ID",), COLLATERAL_KEYS, 3)
    checks = fill_checks(book["Check and correct"], answer_rows(answers, "Check and correct"))
    left = still_needed(book)
    try:
        book.save(path)
    except PermissionError:
        print(f"Could not save {path.name}: close it in Excel and run this again.")
        return 1
    print(f"Typed into {path.name}: {people} borrower value(s), {places} collateral value(s), "
          f"{checks} Check and correct answer(s)")
    if people_unused or places_unused:
        print(f"  {people_unused + places_unused} answer row(s) not needed in this batch (nothing asked for them)")
    if left:
        print(f"  {len(left)} needed cell(s) have no answer:")
        for sheet, label, field in left[:20]:
            print(f"    {sheet}: {label} {field}")
        return 1
    print("Every needed value is filled. Next: run step 2 and step 3 with --dummy.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
