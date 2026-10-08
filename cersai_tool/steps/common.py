"""Shared by steps 1 and 2: turning a batch's frozen data into rows."""
from dataclasses import dataclass, field

from ..context import CollateralContext
from ..model import Issue
from ..rows import bor_row, chg_row, fh_row, imm_row, si_row


@dataclass
class Block:
    """One SI block: the SI, IMM, BOR... and CHG rows of one collateral."""
    context: CollateralContext
    rows: list = field(default_factory=list)       # [(row type, [Cell], borrower dict or None)]

    @property
    def collateral_id(self):
        return self.context.collateral_id

    @property
    def application_number(self):
        return self.context.application_number

    @property
    def held(self) -> bool:
        """Waiting for an answer in the required-details file (PDF_CHECK)."""
        return bool(self.context.hold_reasons)

    @property
    def hold_reasons(self) -> list:
        return self.context.hold_reasons

    def cells(self, row_type):
        return [cells for kind, cells, _ in self.rows if kind == row_type]

    @property
    def tpm(self) -> bool:
        """A third-party mortgage: an owner is not a borrower."""
        return self.cells("SI")[0][si_row.FIELDS.index("IS_TPM")].value == "1"

    def issues(self) -> list:
        return self.context.issues


def build_blocks(data: dict, masters, overrides, run_date) -> list:
    reports = data.get("title_reports", {})
    by_application = {}
    for bundle in data["bundles"]:
        by_application.setdefault(bundle["application_number"], []).append(bundle)
    blocks = []
    for serial, bundle in enumerate(data["bundles"], start=1):
        siblings = [b for b in by_application[bundle["application_number"]] if b is not bundle]
        c = CollateralContext(bundle, reports.get(bundle["collateral_id"]), masters, overrides, serial, run_date,
                              siblings)
        block = Block(c)
        block.rows.append(("SI", si_row.build(c), None))
        block.rows.append(("IMM", imm_row.build(c), None))
        for number, borrower in enumerate(c.borrowers, start=1):
            block.rows.append(("BOR", bor_row.build(c, borrower, number), borrower))
        block.rows.append(("CHG", chg_row.build(), None))
        blocks.append(block)
    return blocks


def file_header(blocks, run_date):
    return fh_row.build(len(blocks), run_date.isoformat())


def batch_issues(data: dict) -> list:
    """Problems found while reading the inputs (not tied to one collateral's rows)."""
    issues = []
    for name, message in data.get("pdf_problems", []):
        issues.append(Issue("", "warning", "Title report file", f"{name}: {message}"))
    for note in data.get("notes", []):
        issues.append(Issue("", "info", "Input data", note))
    for skipped in data.get("skipped", []):
        issues.append(Issue(skipped["collateral_id"], "info", "Already filed",
                            f"left out: already in {skipped['dat_file']} ({skipped['date']})"))
    return issues


def short_list(items, limit=10) -> str:
    """'a, b, c' - or the first ten and 'and N more' for long lists."""
    items = list(items)
    text = ", ".join(items[:limit])
    return text + (f" and {len(items) - limit} more" if len(items) > limit else "")
