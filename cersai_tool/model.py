"""Small data holders shared by all parts of the tool.

A Cell is one value in one row of the batch file, together with where it came
from and whether anything is wrong with it. The Excel files colour each cell
from its status, and the required-details file lists every cell that is
missing, uses a default, or has a conflict.
"""
from dataclasses import dataclass, field

# Where a value came from (shown in cell comments and in the Issues sheet)
SRC_LMS = "LMS"
SRC_PDF = "Title report"
SRC_MANUAL = "Required details"
SRC_DEFAULT = "Default"
SRC_FIXED = "Fixed"
SRC_DERIVED = "Derived"

# Status of a value
OK = "ok"                # found and valid
BLANK = "blank"          # optional and not found: left empty on purpose
NA = "n/a"               # does not apply to this row (e.g. gender for a company)
MISSING = "missing"      # required but not found anywhere: must be filled by hand
DEFAULT = "default"      # a default was used; someone should confirm it
INVALID = "invalid"      # a value was found but breaks a CERSAI rule
CONFLICT = "conflict"    # LMS and title report disagree; see lms_value / pdf_value

NEEDS_INPUT = (MISSING, DEFAULT, INVALID)

# Typing this word in the required-details file leaves an optional value empty
CLEAR_WORD = "BLANK"

# Short names for borrower values typed in the required-details file, and the
# BOR fields each one fills (the borrower type decides which one is used).
BORROWER_KEY_FIELDS = {
    "PAN": ("BORROWER_PAN", "INDIVIDUAL_PAN"),
    "CKYC": ("BORROWER_CKYC_NUMBER", "INDIVIDUAL_CKYC_NUMBER"),
    "NAME": ("BORROWER_NAME", "INDIVIDUAL_NAME"),
    "REG_DATE": ("BORROWER_REGISTRATION_DATE",),
    "REG_NUMBER": ("BORROWER_REGISTRATION_NUMBER",),
    "ADDRESS": ("ADDRESS_LINE_1", "ADDRESS_LINE_2", "ADDRESS_LINE_3"),
}
BORROWER_KEY_OF_FIELD = {f: key for key, fields in BORROWER_KEY_FIELDS.items() if key != "ADDRESS"
                         for f in fields}


def is_clear(value) -> bool:
    return isinstance(value, str) and value.strip().upper() == CLEAR_WORD


@dataclass
class Cell:
    field: str
    value: str = ""
    source: str = ""
    status: str = OK
    note: str = ""
    lms_value: str = ""
    pdf_value: str = ""


def ok(field_name, value, source, note="") -> Cell:
    return Cell(field_name, str(value), source, OK, note)


def fixed(field_name, value) -> Cell:
    return Cell(field_name, str(value), SRC_FIXED, OK)


def derived(field_name, value, note="") -> Cell:
    return Cell(field_name, str(value), SRC_DERIVED, OK, note)


def manual(field_name, value) -> Cell:
    return Cell(field_name, str(value), SRC_MANUAL, OK, "Typed in the required-details file")


def blank(field_name, note="") -> Cell:
    return Cell(field_name, "", "", BLANK, note)


def not_applicable(field_name, note="") -> Cell:
    return Cell(field_name, "", "", NA, note)


def missing(field_name, note) -> Cell:
    return Cell(field_name, "", "", MISSING, note)


def default(field_name, value, note) -> Cell:
    return Cell(field_name, str(value), SRC_DEFAULT, DEFAULT, note)


def invalid(field_name, bad_value, note) -> Cell:
    return Cell(field_name, "", "", INVALID, f"{note} (found: {bad_value})")


def conflict(field_name, value, source, lms_value, pdf_value, note) -> Cell:
    return Cell(field_name, str(value), source, CONFLICT, note, str(lms_value), str(pdf_value))


@dataclass
class Issue:
    """Anything worth telling the reviewer that is not tied to one cell."""
    collateral_id: str
    level: str          # "hold", "warning", "info"
    topic: str
    message: str


@dataclass
class Overrides:
    """Values typed by a person in the required-details file.

    collateral:  (collateral_id, FIELD)            -> value
    borrower:    (application_number, neo_cif, FIELD) -> value
    per_owner:   (collateral_id, neo_cif, FIELD)   -> value  (owner matches, IS_ASSET_OWNER)
    """
    collateral: dict = field(default_factory=dict)
    borrower: dict = field(default_factory=dict)
    per_owner: dict = field(default_factory=dict)

    def is_empty(self) -> bool:
        return not (self.collateral or self.borrower or self.per_owner)

    def update(self, other: "Overrides"):
        """Adds other's values; other wins where both have a value."""
        self.collateral.update(other.collateral)
        self.borrower.update(other.borrower)
        self.per_owner.update(other.per_owner)

    def to_json(self) -> dict:
        return {"collateral": [[*key, value] for key, value in self.collateral.items()],
                "borrower": [[*key, value] for key, value in self.borrower.items()],
                "per_owner": [[*key, value] for key, value in self.per_owner.items()]}

    @classmethod
    def from_json(cls, data) -> "Overrides":
        data = data or {}
        return cls({tuple(item[:-1]): item[-1] for item in data.get("collateral", [])},
                   {tuple(item[:-1]): item[-1] for item in data.get("borrower", [])},
                   {tuple(item[:-1]): item[-1] for item in data.get("per_owner", [])})
