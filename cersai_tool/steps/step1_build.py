"""Step 1: build the working Excel and the required-details file.

1. Reads the CERSAI masters and the CKYC and COLLATERAL extracts (active
   loans only).
2. Leaves out collaterals already filed in an earlier .DAT file (Filed
   register).
3. Creates a new batch folder: <output folder>/<today>, or <today>_2 when a
   batch was already made today. Nothing earlier is ever overwritten.
4. Finds and reads each collateral's title report in the Legal folder.
5. Carries over values typed for the same collaterals and borrowers in the
   latest earlier batch that had them, so nobody types them twice.
6. Saves everything it read in the batch folder ("batch data"), so steps 2
   and 3 work from exactly the same data even if the input files change.
7. Writes "1 Working_<batch>.xlsx" and "2 Required details_<batch>.xlsx".
"""
from datetime import date, datetime
from pathlib import Path

from .. import batch as batches
from .. import settings
from ..model import CONFLICT, NEEDS_INPUT, Overrides
from ..output.excel_layout import write_batch_excel
from ..sources import lms, title_report
from ..sources.masters import Masters
from ..sources.required_details import RequiredDetailsError, read_required_details, write_required_details
from .common import batch_issues, build_blocks, file_header, short_list


def typed_in_batch(folder, masters) -> Overrides:
    """The values typed for a batch: those step 2 applied, or, if step 2 was
    not run, those in its required-details file."""
    data = batches.load_data(folder)
    if data.get("typed") is not None:
        return Overrides.from_json(data["typed"])
    path = batches.batch_files(folder)["required"]
    if not path.exists():
        return Overrides()
    base = build_blocks(data, masters, Overrides(), date.today())
    typed, _ = read_required_details(path, base, masters)
    return typed


def carry_over(output_folder, this_folder, bundles, masters, log):
    """For each collateral and borrower of this batch, the values typed in the
    latest earlier batch that contained it (that batch's values replace any
    older ones entirely)."""
    collaterals = {b["collateral_id"] for b in bundles}
    borrowers = {(b["application_number"], x["neo_cif"]) for b in bundles for x in b["borrowers"]}
    carried, sources = Overrides(), []
    for folder in reversed(batches.list_batches(output_folder)):
        if folder == this_folder or not (collaterals or borrowers):
            continue
        try:
            data = batches.load_data(folder)
        except (batches.BatchError, ValueError):
            continue
        in_folder = {b["collateral_id"] for b in data.get("bundles", [])} & collaterals
        people = {(b["application_number"], x["neo_cif"]) for b in data.get("bundles", [])
                  for x in b["borrowers"]} & borrowers
        if not in_folder and not people:
            continue
        try:
            typed = typed_in_batch(folder, masters)
        except (RequiredDetailsError, OSError, KeyError) as error:
            log(f"  could not read the typed values of batch {folder.name}: {error}")
            typed = Overrides()
        found = Overrides(
            {k: v for k, v in typed.collateral.items() if k[0] in in_folder},
            {k: v for k, v in typed.borrower.items() if (k[0], k[1]) in people},
            {k: v for k, v in typed.per_owner.items() if k[0] in in_folder})
        if not found.is_empty():
            carried.update(found)
            sources.append(folder.name)
        collaterals -= in_folder
        borrowers -= people
    return carried, sources


def summarise(blocks, log):
    needed = sum(1 for b in blocks for _, cells, _ in b.rows for c in cells if c.status in NEEDS_INPUT)
    conflicts = sum(1 for b in blocks for _, cells, _ in b.rows for c in cells if c.status == CONFLICT)
    held = [b.collateral_id for b in blocks if b.held]
    tpm = [b.collateral_id for b in blocks if b.tpm]
    log(f"  values still needed or to confirm: {needed}")
    log(f"  LMS / title report differences:    {conflicts}")
    if held:
        log(f"  held (title report may be for another property): {short_list(held)}")
    if tpm:
        log(f"  third-party mortgages (left out of the .DAT file): {short_list(tpm)}")


def run(paths, log=print, today=None):
    today = today or date.today()
    log("Reading the CERSAI masters...")
    masters = Masters(paths.masters_file)
    log("Reading CKYC and COLLATERAL...")
    bundles, notes, info, known = lms.load_bundles(paths)
    applications = len({b["application_number"] for b in bundles})
    log(f"  {len(bundles)} active collateral(s) of {applications} application(s) in "
        f"{Path(info['collateral_file']).name}")
    for note in notes:
        log(f"  note: {note}")

    register = batches.load_register(paths.register_file)
    skipped, kept = [], []
    for bundle in bundles:
        entry = register.get(bundle["collateral_id"])
        if entry and entry["Status"] == batches.FILED:
            skipped.append({"collateral_id": bundle["collateral_id"], "dat_file": entry["DAT file"],
                            "date": entry["Date"]})
        else:
            kept.append(bundle)
    if skipped:
        log(f"  {len(skipped)} already filed in an earlier .DAT file: left out")
    if not kept:
        log("Nothing to do: every active collateral in the input has already been filed.")
        return None
    if len(kept) > settings.MAX_SI_PER_FILE:
        log(f"  {len(kept)} collaterals: the .DAT output will be split into files of "
            f"{settings.MAX_SI_PER_FILE}")

    folder = batches.new_batch(paths.output_folder, today)
    log(f"New batch: {folder}")
    active_ids = {b["collateral_id"]: b["application_number"] for b in kept}
    found, pdf_problems = title_report.index_folder(paths.legal_folder, active_ids, known)
    log(f"Reading {len(found)} title report(s)...")
    reports = {}
    for coll, path in found.items():
        try:
            reports[coll] = title_report.read_title_report(path)
        except Exception as error:          # a damaged PDF must not stop the batch
            pdf_problems.append((path.name, f"could not be read ({error.__class__.__name__}: {error})"))
    without = [c for c in active_ids if c not in reports]
    if without:
        log(f"  {len(without)} collateral(s) have no title report in {paths.legal_folder.name}")

    carried, carried_from = carry_over(paths.output_folder, folder, kept, masters, log)
    if carried_from:
        log(f"  values typed earlier carried over from batch(es) {', '.join(carried_from)}")

    data = {
        "tool_version": 1,
        "created": datetime.now().isoformat(timespec="seconds"),
        "batch_date": today.isoformat(),
        "input": info,
        "masters_version": masters.version,
        "bundles": kept,
        "title_reports": reports,
        "pdf_problems": pdf_problems,
        "notes": notes,
        "skipped": skipped,
        "carried": carried.to_json(),
        "carried_from": carried_from,
    }
    batches.save_data(folder, data)

    blocks = build_blocks(data, masters, carried, today)
    files = batches.batch_files(folder)
    write_batch_excel(files["working"], file_header(blocks, today), blocks, batch_issues(data),
                      final=False, batch_name=folder.name)
    counts = write_required_details(files["required"], blocks, carried, masters, folder.name)
    log(f"Wrote {files['working'].name}")
    log(f"Wrote {files['required'].name}: {counts['borrowers']} borrower row(s), {counts['collaterals']} "
        f"collateral row(s), {counts['checks']} row(s) to check")
    summarise(blocks, log)
    log("Next: fill the required-details file, save it, then run step 2.")
    return folder
