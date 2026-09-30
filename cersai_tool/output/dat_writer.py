"""Writing the .DAT file(s) for CERSAI.

  - name: crs_<entity code>_012039_<YYYYMMDDHHMISS>.DAT (settings.py);
  - one FH line, then per SI its SI, IMM, BOR... and CHG lines;
  - fields separated by "|", no "|" after the last field, CRLF line ends,
    plain ASCII;
  - at most 1000 SIs per file: a bigger batch is split into several files,
    each with its own FH line and a name one second later than the previous.
When the file is written, the SIs are numbered 1, 2, 3... (and the borrowers
of each SI 1, 2...), the FH row counts the SIs and carries the day's date, and
BATCH_REF_NUMBER is recalculated with that date.
"""
from datetime import timedelta

from .. import settings
from ..rows import bor_row, si_row

SI_SERIAL = si_row.FIELDS.index("SERIAL_NUMBER")
SI_BATCH_REF = si_row.FIELDS.index("BATCH_REF_NUMBER")
BOR_SERIAL = bor_row.FIELDS.index("SERIAL_NUMBER")


def _line(values) -> str:
    text = "|".join(values)
    return text + "|" if settings.TRAILING_PIPE else text


def _digits(text, count=8):
    return "".join(ch for ch in text if ch.isdigit())[-count:].rjust(count, "0")


def dat_lines(blocks, file_date) -> list:
    """blocks: the final-Excel blocks to include, in order."""
    lines = [_line(["FH", "SI_REGN", str(len(blocks)), file_date.isoformat()])]
    for serial, block in enumerate(blocks, start=1):
        borrower_number = 0
        for row in block["rows"]:
            values = list(row["values"])
            if row["type"] == "SI":
                values[SI_SERIAL] = str(serial)
                if settings.GENERATE_BATCH_REF_NUMBER:
                    values[SI_BATCH_REF] = (f"{settings.BATCH_REF_PREFIX}{file_date.strftime('%Y%m%d')}"
                                            f"{_digits(block['collateral_id'])}")
            elif row["type"] == "BOR":
                borrower_number += 1
                values[BOR_SERIAL] = str(borrower_number)
            lines.append(_line(values))
    return lines


def file_name(moment) -> str:
    return f"crs_{settings.ENTITY_CODE}_{settings.FILE_NAME_MIDDLE}_{moment.strftime('%Y%m%d%H%M%S')}.DAT"


def write_dat_files(folder, blocks, moment) -> list:
    """Writes the .DAT file(s) into the batch folder. Returns
    [(path, [collateral IDs in that file])]."""
    written = []
    size = settings.MAX_SI_PER_FILE
    for number, start in enumerate(range(0, len(blocks), size)):
        chunk = blocks[start:start + size]
        when = moment + timedelta(seconds=number)
        path = folder / file_name(when)
        while path.exists():
            when += timedelta(seconds=1)
            path = folder / file_name(when)
        text = settings.LINE_ENDING.join(dat_lines(chunk, when.date())) + settings.LINE_ENDING
        path.write_bytes(text.encode("ascii"))
        written.append((path, [b["collateral_id"] for b in chunk]))
    return written
