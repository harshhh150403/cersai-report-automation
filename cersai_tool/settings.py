"""Settings for the CERSAI batch tool.

Everything that can change without the mapping logic changing lives here: file
locations, CERSAI identity codes, fixed values, and the answers to the open
questions in the mapping notes (Q1-Q32). Change a value here, not in the row
files.
"""
from dataclasses import dataclass
from pathlib import Path

TOOL_FOLDER = Path(__file__).resolve().parent
PROJECT_ROOT = TOOL_FOLDER.parent


# --------------------------------------------------------------------------
# File locations
# --------------------------------------------------------------------------
@dataclass(frozen=True)
class Paths:
    root: Path
    ckyc_folder: Path
    collateral_folder: Path
    legal_folder: Path
    masters_file: Path
    output_folder: Path

    @property
    def register_file(self) -> Path:
        return self.output_folder / "Filed register.xlsx"


def make_paths(root=None, dummy: bool = False) -> Paths:
    """Real runs read "Input data" and write "Output data".
    Development runs (--dummy) read "Dummy input data" and write
    "Dummy output data", so test batches never mix with real ones."""
    root = Path(root).resolve() if root else PROJECT_ROOT
    masters = root / "Master data" / "CERSAI2.0_Masters.xls"
    if dummy:
        base = root / "Dummy input data"
        return Paths(root, base / "Dummy CKYC", base / "Dummy COLLATERAL",
                     base / "Dummy Legal", masters, root / "Dummy output data")
    base = root / "Input data"
    return Paths(root, base / "CKYC", base / "COLLATERAL", base / "Legal",
                 masters, root / "Output data")


# --------------------------------------------------------------------------
# CERSAI file format (open questions Q1, Q3, Q22 in the mapping notes)
# --------------------------------------------------------------------------
ENTITY_CODE = "J8084"            # ARGFL's CERSAI entity code (file name)
FILE_NAME_MIDDLE = "012039"      # literal part of crs_NNNNN_012039_... (Q3)
LINE_ENDING = "\r\n"             # Windows line endings (Q22)
TRAILING_PIPE = False            # no "|" at the end of a line (Q1)
MAX_SI_PER_FILE = 1000           # CERSAI limit per .DAT file

# --------------------------------------------------------------------------
# Fixed and default values
# --------------------------------------------------------------------------
IMM_COUNTRY = "IND"              # IMM row: the batch format description says "as per master", and the
                                 # COUNTRY master's code for India is IND
BOR_CHG_COUNTRY = "IN"           # BOR and CHG rows: the batch format description gives the value "IN"
                                 # (a constant for CHG); these rows do not use the COUNTRY master
FINANCING_TYPE_ID = "1"          # Sole
CONSORTIUM_MEMBERS = "0"
DEFAULT_SI_TYPE_ID = "17"        # Equitable Mortgage, used when no title report
TYPE_OF_CHARGE = "FIRST CHARGE"
INCLUDE_LAN_IN_NARRATION = False     # True -> NARRATION = "LAN <loan account>"
USE_GLOBAL_COLLATERAL_NO_AS_ASSET_UNIQUE_ID = True
GENERATE_BATCH_REF_NUMBER = True     # "ARGF" + file date + 8 collateral digits (Q9)
BATCH_REF_PREFIX = "ARGF"
SURVEY_SEPARATOR = "^"           # between several survey numbers (Q15)
LMS_AREA_UNIT_ID = "6"           # LMS "Carpert Area" assumed to be sq ft (Q19, to confirm)
INCLUDE_GUARANTORS = False       # guarantors are not borrowers (Q10)

# --------------------------------------------------------------------------
# Matching thresholds (0-100 name similarity)
# --------------------------------------------------------------------------
SAME_PERSON_SCORE = 90           # at or above: the same person
PROBABLE_PERSON_SCORE = 75       # between this and the above: similar names, decided and listed for review
RIGHT_PDF_HOLD_FAILURES = 2      # this many failed "right PDF" checks holds the SI
AREA_TOLERANCE = 0.10            # LMS vs title-report area: flag if more than 10% apart

# --------------------------------------------------------------------------
# Excel output
# --------------------------------------------------------------------------
WRITE_FIELD_NAME_COMMENTS = True     # the demo layout names each cell in a comment
