"""One file per row type of the CERSAI batch file (FH, SI, IMM, BOR, CHG).

Inside each file there is one function per column, in CERSAI's column order,
with a plain-English comment above it explaining where the value comes from,
the rules applied, and what happens when it cannot be found.

Order used for every value (see the mapping notes):
  1. a value typed in the required-details file always wins;
  2. otherwise CKYC / COLLATERAL (the LMS extracts);
  3. otherwise the legal title report;
  4. otherwise a default, or the value is reported as missing.
Property owners are the exception: they are always checked against the title
report (see NUMBER_OF_ASSET_OWNERS in si_row.py).
"""
from ..model import CLEAR_WORD, blank


def build_row(fields, column_functions, *args, cleared=()):
    """Runs the column functions in order. Each gets the arguments plus the
    cells already built for this row (some columns depend on earlier ones).
    Fields in 'cleared' (BLANK typed in the required-details file) are left
    empty whatever the sources say."""
    row = {}
    for column in column_functions:
        cell = column(*args, row)
        if cell.field in cleared and cell.field not in ("ROW_TYPE", "SERIAL_NUMBER"):
            cell = blank(cell.field, f"left empty: {CLEAR_WORD} was typed in the required-details file")
        row[cell.field] = cell
    return [row[name] for name in fields]
