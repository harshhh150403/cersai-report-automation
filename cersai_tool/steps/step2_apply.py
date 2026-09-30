"""Step 2: apply the required details and make the final Excel.

1. Picks the batch (the latest one with a required-details file, unless a
   batch is named) and loads the data step 1 saved.
2. Reads the required-details file: a value counts as typed when it differs
   from what the tool finds on its own.
3. Rebuilds every row with those values and writes "3 Final_<batch>.xlsx".
   If a final file already exists and was edited by hand since step 2 made
   it, the edited file is kept as a backup and the user is told.
4. Records what was typed, which SIs are held, and the final file's
   fingerprint (to notice later hand edits).
Step 2 can be run as often as needed: each run reads the required-details
file again.
"""
from datetime import date, datetime

from .. import batch as batches
from ..model import CONFLICT, NEEDS_INPUT, Issue, Overrides
from ..output.excel_layout import cell_problems, write_batch_excel
from ..sources.masters import Masters
from ..sources.required_details import read_required_details
from .common import batch_issues, build_blocks, file_header


def notices(blocks) -> dict:
    """Defaults, LMS / title report differences and warnings still open, per
    collateral: step 3 repeats them in its report so they are seen before
    upload."""
    out = {}
    for block in blocks:
        items = [[row[2] or "SI", row[4] or "", f"{row[6]}: {row[7]}"] for row in cell_problems(block)
                 if row[5] == "check"]
        items += [["", "", f"{i.topic}: {i.message}"] for i in block.issues() if i.level == "warning"]
        if items:
            out[block.collateral_id] = items
    return out


def run(paths, batch_name=None, assume_yes=False, log=print, today=None):
    today = today or date.today()
    folder = batches.pick_batch(paths.output_folder, batch_name, "required", assume_yes, log)
    files = batches.batch_files(folder)
    data = batches.load_data(folder)
    log("Reading the CERSAI masters...")
    masters = Masters(paths.masters_file)

    base = build_blocks(data, masters, Overrides(), today)
    log(f"Reading {files['required'].name}...")
    typed, notes = read_required_details(files["required"], base, masters)
    count = len(typed.collateral) + len(typed.borrower) + len(typed.per_owner)
    log(f"  {count} value(s) typed")
    for note in notes:
        log(f"  note: {note}")

    blocks = build_blocks(data, masters, typed, today)
    final = files["final"]
    if final.exists():
        if batches.file_hash(final) != data.get("final_hash"):
            kept = batches.backup(final)
            log(f"  {final.name} had been changed by hand since step 2 made it. The changed file is kept as")
            log(f"  '{kept.name}'. Copy any correction you still need into the new final file, or better,")
            log("  into the required-details file so it is not lost again.")
        else:
            final.unlink()
    issues = batch_issues(data) + [Issue("", "warning", "Required details", note) for note in notes]
    write_batch_excel(final, file_header(blocks, today), blocks, issues, final=True, batch_name=folder.name)

    data["typed"] = typed.to_json()
    data["final_hash"] = batches.file_hash(final)
    data["held"] = {b.collateral_id: "; ".join(b.hold_reasons) for b in blocks if b.held}
    data["notices"] = notices(blocks)
    data["step2_at"] = datetime.now().isoformat(timespec="seconds")
    batches.save_data(folder, data)

    log(f"Wrote {final.name}")
    ready, waiting = [], []
    for block in blocks:
        problems = list(dict.fromkeys(f"{kind} {c.field}" for kind, cells, _ in block.rows for c in cells
                                      if c.status in NEEDS_INPUT and c.status != "default"))
        if problems or block.held or block.tpm:
            reason = f"held: {'; '.join(block.hold_reasons)}" if block.held else \
                "third-party mortgage" if block.tpm else \
                f"missing/invalid: {', '.join(problems[:4])}{' ...' if len(problems) > 4 else ''}"
            waiting.append(f"{block.collateral_id}: {reason}")
        else:
            ready.append(block.collateral_id)
    conflicts = sum(1 for b in blocks for _, cells, _ in b.rows for c in cells if c.status == CONFLICT)
    defaults = sum(1 for b in blocks for _, cells, _ in b.rows for c in cells if c.status == "default")
    log(f"  ready for the .DAT file: {len(ready)} of {len(blocks)} SI(s)")
    if conflicts or defaults:
        log(f"  still to look at: {conflicts} LMS / title report difference(s), {defaults} default value(s) "
            "(sheet Issues)")
    for line in waiting[:15]:
        log(f"  not ready - {line}")
    if len(waiting) > 15:
        log(f"  ... and {len(waiting) - 15} more (sheet Issues of the final Excel)")
    if waiting:
        log("Fill the missing values in the required-details file and run step 2 again, or run step 3 now:")
        log("SIs that are not ready are left out of the .DAT file and stay pending.")
    else:
        log("Next: review the final Excel, then run step 3.")
    return folder
