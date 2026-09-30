"""Batch folders, the frozen batch data, backups and the filed register.

Each run of step 1 creates a new batch folder under the output folder, named
after the day (2026-09-30; a second batch that day is 2026-09-30_2). Every
file of the batch lives in that folder, and nothing is ever overwritten: when
a file has to be replaced, the old one is renamed with a "(backup ...)" tag.
"""
import hashlib
import json
import re
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

BATCH_NAME_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})(?:_(\d+))?$")
DATA_FILE = "batch data (do not edit).json"


class BatchError(Exception):
    """A problem with the batch folders that stops the run."""


def batch_files(folder: Path) -> dict:
    tag = folder.name
    return {
        "working": folder / f"1 Working_{tag}.xlsx",
        "required": folder / f"2 Required details_{tag}.xlsx",
        "final": folder / f"3 Final_{tag}.xlsx",
        "checks": folder / f"4 Checks_{tag}.xlsx",
        "data": folder / DATA_FILE,
    }


def list_batches(output_folder: Path) -> list:
    if not output_folder.exists():
        return []
    batches = []
    for folder in output_folder.iterdir():
        m = BATCH_NAME_RE.match(folder.name)
        if folder.is_dir() and m:
            batches.append((m.group(1), int(m.group(2) or 1), folder))
    return [folder for _, _, folder in sorted(batches)]


def new_batch(output_folder: Path, today: date) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    name = today.isoformat()
    number = 1
    while (output_folder / name).exists():
        number += 1
        name = f"{today.isoformat()}_{number}"
    folder = output_folder / name
    folder.mkdir()
    return folder


def pick_batch(output_folder: Path, name, needs: str, assume_yes: bool, log=print) -> Path:
    """Chooses the batch to work on: the one named, or else the latest batch
    that has the file this step needs. Asks for confirmation unless
    assume_yes."""
    if name:
        folder = output_folder / name
        if not folder.is_dir():
            raise BatchError(f"No batch folder named '{name}' in {output_folder}")
    else:
        candidates = [b for b in list_batches(output_folder) if batch_files(b)[needs].exists()]
        if not candidates:
            raise BatchError(f"No batch in {output_folder} has a '{batch_files(Path('x'))[needs].name.split('_')[0]}' "
                             "file yet. Run the previous step first.")
        folder = candidates[-1]
        if not assume_yes:
            answer = input(f"Use batch {folder.name}? Press Enter for yes, or type another batch name: ").strip()
            if answer:
                folder = output_folder / answer
                if not folder.is_dir():
                    raise BatchError(f"No batch folder named '{answer}' in {output_folder}")
    if not batch_files(folder)[needs].exists():
        raise BatchError(f"Batch {folder.name} has no file {batch_files(folder)[needs].name}")
    log(f"Batch: {folder.name}")
    return folder


def save_data(folder: Path, data: dict):
    path = batch_files(folder)["data"]
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=1, ensure_ascii=True, default=str), encoding="utf-8")
    tmp.replace(path)


def load_data(folder: Path) -> dict:
    path = batch_files(folder)["data"]
    if not path.exists():
        raise BatchError(f"{path.name} is missing from batch {folder.name}; run step 1 again")
    return json.loads(path.read_text(encoding="utf-8"))


def file_hash(path: Path) -> str:
    return hashlib.sha1(path.read_bytes()).hexdigest() if path.exists() else ""


def backup(path: Path) -> Path:
    """Renames an existing file to '<name> (backup YYYY-MM-DD HHMMSS)<ext>'."""
    stamp = datetime.now().strftime("%Y-%m-%d %H%M%S")
    target = path.with_name(f"{path.stem} (backup {stamp}){path.suffix}")
    path.rename(target)
    return target


# --------------------------------------------------------------------------
# The filed register: which collaterals have gone into a .DAT file
# --------------------------------------------------------------------------
REGISTER_HEADERS = ["Collateral ID", "Application number", "Batch", "Status", "DAT file", "Date", "Reason"]
FILED = "Filed"
PENDING = "Pending"


def load_register(path: Path) -> dict:
    if not path.exists():
        return {}
    book = openpyxl.load_workbook(path, read_only=True)
    sheet = book.active
    entries = {}
    for i, values in enumerate(sheet.iter_rows(values_only=True)):
        if i == 0 or not values or not values[0]:
            continue
        entry = {h: ("" if v is None else str(v)) for h, v in zip(REGISTER_HEADERS, values)}
        entries[entry["Collateral ID"]] = entry
    book.close()
    return entries


def save_register(path: Path, entries: dict):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Filed register"
    sheet.append(REGISTER_HEADERS)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="FFD9D9D9")
    for entry in sorted(entries.values(), key=lambda e: (e["Date"], e["Collateral ID"])):
        sheet.append([entry.get(h, "") for h in REGISTER_HEADERS])
    for column, width in zip("ABCDEFG", (16, 18, 14, 10, 40, 12, 60)):
        sheet.column_dimensions[column].width = width
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)
