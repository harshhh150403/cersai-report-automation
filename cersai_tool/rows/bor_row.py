"""BOR row: one per borrower of the SI. 27 columns (sheet "borrower").

This file holds columns 1-19 (who the borrower is). Columns 20-27 (address and
location codes) are in bor_address.py.

Borrowers come from CKYC: one row per customer (Neo CIF ID) of the
application, primary borrower first, guarantors left out. CKYC lists a
customer once per ID document, so its rows are grouped per customer first.

Which columns are filled depends on the borrower type (column 3):
  - not IND (companies, firms, sole proprietorships...): columns 7-11;
  - IND (individuals): columns 12-19.
The other group stays blank.
"""
import re

from ..model import (SRC_LMS, blank, derived, fixed, invalid, manual, missing, not_applicable, ok)
from ..rules import dates
from ..rules.identifiers import (PAN_TYPE_LETTER, clean_ckyc, clean_cin, clean_email, clean_llpin,
                                 clean_mobile, clean_pan)
from ..rules.text import ORG_NAME, as_text, clean

FIELDS = ["ROW_TYPE", "SERIAL_NUMBER", "BORROWER_TYPE", "IS_ASSET_OWNER", "BORROWER_UID_TYPE",
          "BORROWER_UID_VALUE", "BORROWER_PAN", "BORROWER_CKYC_NUMBER", "BORROWER_NAME",
          "BORROWER_REGISTRATION_DATE", "BORROWER_REGISTRATION_NUMBER", "INDIVIDUAL_PAN",
          "INDIVIDUAL_CKYC_NUMBER", "GENDER", "INDIVIDUAL_NAME", "FATHER_MOTHER_NAME", "DATE_OF_BIRTH",
          "MOBILE_NUMBER", "EMAIL", "ADDRESS_LINE_1", "ADDRESS_LINE_2", "ADDRESS_LINE_3", "CITY", "DISTRICT",
          "STATE", "PIN_CODE", "COUNTRY"]

# CKYC "CONSTITUTION" -> CERSAI borrower type (BORROWER_ASSET_OWNER_TYPE master)
CONSTITUTION_CODES = {
    "individual": "IND", "no pan": "IND",
    "sole proprietorship firm": "SPR", "sole proprietorship": "SPR", "proprietorship": "SPR",
    "partnership firm": "PRF", "partnership": "PRF",
    "limited liability partnership": "LLP", "llp": "LLP",
    "private limited company": "COM", "public limited company": "COM", "company": "COM",
    "one person company": "COM",
    "hindu undivided company": "HUF", "hindu undivided family": "HUF", "huf": "HUF",
    "private trust": "TRS", "public trust": "TRS", "trust": "TRS",
    "society": "SOC", "association of persons": "AOP", "body of individuals": "BOI",
    "government body": "GOV", "foreign company": "FRC",
}
GENDER_WORDS = {"M": "M", "MALE": "M", "F": "F", "FEMALE": "F", "T": "T", "TRANSGENDER": "T",
                "THIRD GENDER": "T"}


def is_individual(row, b) -> bool:
    code = row["BORROWER_TYPE"].value
    if code:
        return code == "IND"
    return b["applicant_type"].lower() == "individual"


# Column 1 - ROW_TYPE
# Always the text "BOR".
def col_01_row_type(c, b, serial, row):
    return fixed("ROW_TYPE", "BOR")


# Column 2 - SERIAL_NUMBER
# 1, 2, 3... for the borrowers of this SI, primary borrower first.
def col_02_serial_number(c, b, serial, row):
    return derived("SERIAL_NUMBER", serial, "borrower number within the SI")


# Column 3 - BORROWER_TYPE (a code from the BORROWER_ASSET_OWNER_TYPE master)
# Where it comes from: CKYC "CONSTITUTION", through a fixed list because the
#   LMS wording differs from CERSAI's: Individual / NO PAN -> IND,
#   Sole Proprietorship Firm -> SPR, Partnership Firm -> PRF,
#   Limited Liability Partnership -> LLP, Private/Public Limited Company -> COM,
#   Hindu Undivided Company -> HUF, Private Trust -> TRS.
# Rules: CKYC "Applicant Type" is not used to decide (it is sometimes wrong);
#   a disagreement is only noted.
# If it can't be found: an unknown constitution is reported as missing.
def col_03_borrower_type(c, b, serial, row):
    typed = c.bov(b, "BORROWER_TYPE")
    if typed is not None:
        code = as_text(typed).upper()
        if code not in c.masters.borrower_types:
            return invalid("BORROWER_TYPE", code, "not a BORROWER_TYPE code")
        return manual("BORROWER_TYPE", code)
    constitution = b["constitution"]
    code = CONSTITUTION_CODES.get(constitution.strip().lower())
    if not code:
        return missing("BORROWER_TYPE", f"unknown CKYC CONSTITUTION '{constitution}'")
    note = f"CKYC CONSTITUTION '{constitution}'"
    applicant = b["applicant_type"].lower()
    if (code == "IND") != (applicant == "individual"):
        note += f"; CKYC Applicant Type says '{b['applicant_type']}'"
    return ok("BORROWER_TYPE", code, SRC_LMS, note)


# Column 4 - IS_ASSET_OWNER (1 or 0)
# Where it comes from: the property owners decided for the SI row (COLLATERAL
#   "Owner Name" first; the title report when COLLATERAL has none - see
#   NUMBER_OF_ASSET_OWNERS). 1 when one of the owners is this borrower (names
#   compared ignoring titles, spelling, initials and a relative's name after
#   S/O, D/O or W/O), otherwise 0.
# Rules: similar names count as the same person and names whose given names
#   differ (relatives) do not; such decisions are listed in the Check and
#   correct sheet, where a Y or N changes them. With several collaterals on one
#   application, this is worked out per collateral: the same borrower can own
#   one property and not another.
def col_04_is_asset_owner(c, b, serial, row):
    typed = c.pov(b, "IS_ASSET_OWNER")
    if typed is not None:
        return manual("IS_ASSET_OWNER", "1" if as_text(typed) in ("1", "Y", "YES") else "0")
    links = [link for link in c.owner_links if link.neo_cif == b["neo_cif"]]
    if not links:
        return derived("IS_ASSET_OWNER", "0", "not one of the property owners")
    note = "owner: " + "; ".join(link.owner for link in links)
    if any(link.uncertain for link in links):
        note += " (similar names: listed in Check and correct)"
    return derived("IS_ASSET_OWNER", "1", note)


# Columns 5 and 6 - BORROWER_UID_TYPE and BORROWER_UID_VALUE
# Not used in this version of the batch format: always empty.
def col_05_borrower_uid_type(c, b, serial, row):
    return not_applicable("BORROWER_UID_TYPE", "must be empty")


def col_06_borrower_uid_value(c, b, serial, row):
    return not_applicable("BORROWER_UID_VALUE", "must be empty")


def _pan(c, b):
    """PAN from CKYC (the row whose Identification Type is PAN), else the
    'Identification Number' of the customer's COLLATERAL rows."""
    for raw, label in ((b["pan"], "CKYC PAN"), (b["collateral_pan"], "COLLATERAL Identification Number")):
        pan = clean_pan(raw)
        if pan:
            return pan, label
    return "", ""


def _pan_note(code, pan):
    expected = PAN_TYPE_LETTER.get(code)
    if expected and pan[3] != expected:
        return f"; note: 4th PAN letter '{pan[3]}' is unusual for {code}"
    return ""


# Column 7 - BORROWER_PAN (non-individuals only)
# Where it comes from: CKYC, the customer's row whose Identification Type is
#   PAN; else COLLATERAL "Identification Number". Format: 5 letters, 4 digits,
#   1 letter. A 4th letter that does not fit the borrower type is noted.
# If it can't be found: blank - then a CKYC number (column 8) is required.
def col_07_borrower_pan(c, b, serial, row):
    if is_individual(row, b):
        return not_applicable("BORROWER_PAN", "individuals use column 12")
    typed = c.bov(b, "PAN", "BORROWER_PAN")
    if typed is not None:
        pan = clean_pan(typed)
        return manual("BORROWER_PAN", pan) if pan else invalid("BORROWER_PAN", typed, "not a valid PAN")
    pan, label = _pan(c, b)
    if pan:
        return ok("BORROWER_PAN", pan, SRC_LMS, label + _pan_note(row["BORROWER_TYPE"].value, pan))
    return blank("BORROWER_PAN", "no PAN in CKYC or COLLATERAL")


# Column 8 - BORROWER_CKYC_NUMBER (non-individuals only; 14 digits)
# Where it comes from: CKYC, a row whose Identification Type is "KYC
#   Identification Number" - only if it really is 14 digits (the extract's
#   values are usually GST-style numbers, which do not count).
# If it can't be found: required when there is no PAN -> reported as missing;
#   otherwise blank.
def col_08_borrower_ckyc_number(c, b, serial, row):
    if is_individual(row, b):
        return not_applicable("BORROWER_CKYC_NUMBER", "individuals use column 13")
    typed = c.bov(b, "CKYC", "BORROWER_CKYC_NUMBER")
    if typed is not None:
        number = clean_ckyc(typed)
        return manual("BORROWER_CKYC_NUMBER", number) if number else invalid(
            "BORROWER_CKYC_NUMBER", typed, "a CKYC number has 14 digits")
    number = clean_ckyc(b["ckyc_number"])
    if number:
        return ok("BORROWER_CKYC_NUMBER", number, SRC_LMS, "CKYC KYC Identification Number")
    if not row["BORROWER_PAN"].value:
        return missing("BORROWER_CKYC_NUMBER", "no PAN, so a 14-digit CKYC number is required")
    return blank("BORROWER_CKYC_NUMBER", "PAN is present")


# Column 9 - BORROWER_NAME (non-individuals only; max 100 characters)
# Where it comes from: CKYC "Customer Name".
# Rules: "&" becomes "AND"; brackets and dots are removed ("JOGARAM PUROHIT (HUF)"
#   -> "JOGARAM PUROHIT HUF"); only letters, digits, "-", "_" and spaces stay.
# If it can't be found: reported as missing.
def col_09_borrower_name(c, b, serial, row):
    if is_individual(row, b):
        return not_applicable("BORROWER_NAME", "individuals use column 15")
    typed = c.bov(b, "NAME", "BORROWER_NAME")
    if typed is not None:
        return manual("BORROWER_NAME", clean(typed, ORG_NAME, 100, ampersand_to_and=True))
    name = clean(b["name"], ORG_NAME, 100, ampersand_to_and=True)
    if name:
        return ok("BORROWER_NAME", name, SRC_LMS, "CKYC Customer Name")
    return missing("BORROWER_NAME", "no customer name in CKYC")


# Column 10 - BORROWER_REGISTRATION_DATE (non-individuals only)
# Where it comes from: CKYC "Incorporation Date" (DD-MON-YY), turned into
#   YYYY-MM-DD. The 2-digit year means 20YY unless that would be in the future,
#   in which case 19YY ("01-JUL-04" -> 2004-07-01).
# If it can't be found: reported as missing.
def col_10_borrower_registration_date(c, b, serial, row):
    if is_individual(row, b):
        return not_applicable("BORROWER_REGISTRATION_DATE", "individuals use column 17")
    typed = c.bov(b, "REG_DATE", "BORROWER_REGISTRATION_DATE")
    if typed is not None:
        value = dates.to_cersai(typed, dates.PAST)
        return manual("BORROWER_REGISTRATION_DATE", value) if value else invalid(
            "BORROWER_REGISTRATION_DATE", typed, "not a date")
    value = dates.to_cersai(b["incorporation_date"], dates.PAST)
    if value:
        return ok("BORROWER_REGISTRATION_DATE", value, SRC_LMS, f"CKYC Incorporation Date '{b['incorporation_date']}'")
    return missing("BORROWER_REGISTRATION_DATE", "no incorporation date in CKYC")


# Column 11 - BORROWER_REGISTRATION_NUMBER
# Only for companies (COM: the 21-character CIN) and LLPs (LLP: the LLPIN,
#   written without its hyphen, e.g. AAB1234). Blank for everyone else.
# Where it comes from: CKYC, a row whose Identification Type is "CIN No" (or
#   an LLPIN row). The CKYC extract rarely has one.
# If it can't be found: for COM and LLP, reported as missing.
def col_11_borrower_registration_number(c, b, serial, row):
    code = row["BORROWER_TYPE"].value
    if is_individual(row, b):
        return not_applicable("BORROWER_REGISTRATION_NUMBER", "not for individuals")
    if code not in ("COM", "LLP"):
        return not_applicable("BORROWER_REGISTRATION_NUMBER", "only needed for companies and LLPs")
    tidy = clean_cin if code == "COM" else clean_llpin
    kind = "CIN" if code == "COM" else "LLPIN"
    typed = c.bov(b, "REG_NUMBER", "BORROWER_REGISTRATION_NUMBER")
    if typed is not None:
        number = tidy(typed)
        return manual("BORROWER_REGISTRATION_NUMBER", number) if number else invalid(
            "BORROWER_REGISTRATION_NUMBER", typed, f"not a valid {kind}")
    number = tidy(b["cin"] if code == "COM" else b["llpin"])
    if number:
        return ok("BORROWER_REGISTRATION_NUMBER", number, SRC_LMS, f"CKYC {kind}")
    return missing("BORROWER_REGISTRATION_NUMBER", f"{kind} required for {code}; not in CKYC")


# Column 12 - INDIVIDUAL_PAN (individuals only)
# Where it comes from: CKYC, the customer's PAN row; else COLLATERAL
#   "Identification Number".
# If it can't be found: blank - then a CKYC number (column 13) is required.
def col_12_individual_pan(c, b, serial, row):
    if not is_individual(row, b):
        return not_applicable("INDIVIDUAL_PAN", "non-individuals use column 7")
    typed = c.bov(b, "PAN", "INDIVIDUAL_PAN")
    if typed is not None:
        pan = clean_pan(typed)
        return manual("INDIVIDUAL_PAN", pan) if pan else invalid("INDIVIDUAL_PAN", typed, "not a valid PAN")
    pan, label = _pan(c, b)
    if pan:
        return ok("INDIVIDUAL_PAN", pan, SRC_LMS, label + _pan_note("IND", pan))
    return blank("INDIVIDUAL_PAN", "no PAN in CKYC or COLLATERAL")


# Column 13 - INDIVIDUAL_CKYC_NUMBER (individuals only; 14 digits)
# Where it comes from: a CKYC "KYC Identification Number" row, if it is 14 digits.
# If it can't be found: required when there is no PAN -> reported as missing.
def col_13_individual_ckyc_number(c, b, serial, row):
    if not is_individual(row, b):
        return not_applicable("INDIVIDUAL_CKYC_NUMBER", "non-individuals use column 8")
    typed = c.bov(b, "CKYC", "INDIVIDUAL_CKYC_NUMBER")
    if typed is not None:
        number = clean_ckyc(typed)
        return manual("INDIVIDUAL_CKYC_NUMBER", number) if number else invalid(
            "INDIVIDUAL_CKYC_NUMBER", typed, "a CKYC number has 14 digits")
    number = clean_ckyc(b["ckyc_number"])
    if number:
        return ok("INDIVIDUAL_CKYC_NUMBER", number, SRC_LMS, "CKYC KYC Identification Number")
    if not row["INDIVIDUAL_PAN"].value:
        return missing("INDIVIDUAL_CKYC_NUMBER", "no PAN, so a 14-digit CKYC number is required")
    return blank("INDIVIDUAL_CKYC_NUMBER", "PAN is present")


# Column 14 - GENDER (individuals only: M, F or T)
# Where it comes from: CKYC "Gender" ("Male" -> M, "Female" -> F; M, F, T and
#   "Transgender" are understood too).
# If it can't be found: typed in the required-details file. (Older CKYC
#   extracts have no Gender column; then every individual's gender is typed.)
def col_14_gender(c, b, serial, row):
    if not is_individual(row, b):
        return not_applicable("GENDER", "only for individuals")
    typed = c.bov(b, "GENDER")
    if typed is not None:
        code = GENDER_WORDS.get(as_text(typed).upper())
        return manual("GENDER", code) if code else invalid("GENDER", typed, "must be M, F or T")
    found = as_text(b.get("gender", ""))       # batches made before the Gender column have no "gender"
    code = GENDER_WORDS.get(found.upper())
    if code:
        return ok("GENDER", code, SRC_LMS, f"CKYC Gender '{found}'")
    if found:
        return missing("GENDER", f"CKYC Gender '{found}' is not Male or Female: type M, F or T")
    return missing("GENDER", "not in CKYC: type M, F or T")


# Column 15 - INDIVIDUAL_NAME (individuals only; letters and spaces, max 50)
# Where it comes from: CKYC "Customer Name" (dots, digits and other characters
#   removed: "SARVES ." -> "SARVES"); if empty, First + Middle + Last Name.
# If it can't be found: reported as missing.
def col_15_individual_name(c, b, serial, row):
    if not is_individual(row, b):
        return not_applicable("INDIVIDUAL_NAME", "non-individuals use column 9")
    typed = c.bov(b, "NAME", "INDIVIDUAL_NAME")
    if typed is not None:
        return manual("INDIVIDUAL_NAME", clean(typed, "", 50, letters_only=True))
    name = clean(b["name"], "", 50, letters_only=True)
    if not name:
        name = clean(" ".join([b["first_name"], b["middle_name"], b["last_name"]]), "", 50, letters_only=True)
    if name:
        return ok("INDIVIDUAL_NAME", name, SRC_LMS, "CKYC Customer Name")
    return missing("INDIVIDUAL_NAME", "no name in CKYC")


# Column 16 - FATHER_MOTHER_NAME (individuals only; letters and spaces, max 50)
# Where it comes from: COLLATERAL "Father Name", a column filled by hand, when
#   the extract has it: the borrower's name on any row of the application.
#   (Not in CKYC or the title reports. The middle name is not used: it is
#   sometimes "WO", i.e. "wife of".)
# If it can't be found: typed in the required-details file. When the
#   borrower's rows give different names, none is used and it is asked for.
def col_16_father_mother_name(c, b, serial, row):
    if not is_individual(row, b):
        return not_applicable("FATHER_MOTHER_NAME", "only for individuals")
    typed = c.bov(b, "FATHER_MOTHER_NAME")
    if typed is not None:
        name = clean(typed, "", 50, letters_only=True)
        return manual("FATHER_MOTHER_NAME", name) if name else invalid("FATHER_MOTHER_NAME", typed, "letters only")
    column, problem = c.father_name(b)
    if column:
        name = clean(column, "", 50, letters_only=True)
        if name:
            return ok("FATHER_MOTHER_NAME", name, SRC_LMS, "COLLATERAL Father Name")
        problem = f"COLLATERAL Father Name '{column}' has no letters"
    return missing("FATHER_MOTHER_NAME", "not in CKYC or COLLATERAL (Father Name): type the father's or mother's "
                                         "name" + (f". {problem}" if problem else ""))


# Column 17 - DATE_OF_BIRTH (individuals only)
# Where it comes from: CKYC "Date Of Birth" (DD-MON-YY), turned into
#   YYYY-MM-DD. The 2-digit year means 20YY only if that makes the person at
#   least 18 today, otherwise 19YY ("11-JAN-51" -> 1951-01-11).
# If it can't be found: reported as missing.
def col_17_date_of_birth(c, b, serial, row):
    if not is_individual(row, b):
        return not_applicable("DATE_OF_BIRTH", "only for individuals")
    typed = c.bov(b, "DATE_OF_BIRTH")
    if typed is not None:
        value = dates.to_cersai(typed, dates.BIRTH)
        return manual("DATE_OF_BIRTH", value) if value else invalid("DATE_OF_BIRTH", typed, "not a date")
    value = dates.to_cersai(b["date_of_birth"], dates.BIRTH)
    if value:
        return ok("DATE_OF_BIRTH", value, SRC_LMS, f"CKYC Date Of Birth '{b['date_of_birth']}'")
    return missing("DATE_OF_BIRTH", "no date of birth in CKYC")


# Column 18 - MOBILE_NUMBER (individuals only; optional; 10 digits)
# Where it comes from: CKYC "Mobile No" with the "+91--" prefix removed; if
#   that is not 10 digits, COLLATERAL "PHONE_NUMBER".
# If it can't be found: blank.
def col_18_mobile_number(c, b, serial, row):
    if not is_individual(row, b):
        return not_applicable("MOBILE_NUMBER", "only for individuals")
    typed = c.bov(b, "MOBILE_NUMBER")
    if typed is not None:
        number = clean_mobile(typed)
        return manual("MOBILE_NUMBER", number) if number else invalid("MOBILE_NUMBER", typed, "not 10 digits")
    for raw, label in ((b["mobile"], "CKYC Mobile No"), (b["collateral_phone"], "COLLATERAL PHONE_NUMBER")):
        number = clean_mobile(raw)
        if number:
            return ok("MOBILE_NUMBER", number, SRC_LMS, label)
    return blank("MOBILE_NUMBER", "optional; no valid 10-digit number")


# Column 19 - EMAIL (individuals only; optional)
# Where it comes from: CKYC "Email ID" (spaces trimmed); if missing or not
#   valid, COLLATERAL "EMAIL_ADDRESS".
# Rules: CERSAI allows letters, digits, "_" and "." with one "@" and a dot
#   after it (no hyphens), max 50 characters. An invalid email is left out.
# If it can't be found: blank.
def col_19_email(c, b, serial, row):
    if not is_individual(row, b):
        return not_applicable("EMAIL", "only for individuals")
    typed = c.bov(b, "EMAIL")
    if typed is not None:
        email = clean_email(typed)
        return manual("EMAIL", email) if email else invalid("EMAIL", typed, "not a valid email for CERSAI")
    for raw, label in ((b["email"], "CKYC Email ID"), (b["collateral_email"], "COLLATERAL EMAIL_ADDRESS")):
        email = clean_email(raw)
        if email:
            return ok("EMAIL", email, SRC_LMS, label)
    return blank("EMAIL", "optional; none valid in CKYC or COLLATERAL")


def build(c, b, serial):
    from . import build_row
    from .bor_address import ADDRESS_COLUMNS
    columns = [col_01_row_type, col_02_serial_number, col_03_borrower_type, col_04_is_asset_owner,
               col_05_borrower_uid_type, col_06_borrower_uid_value, col_07_borrower_pan,
               col_08_borrower_ckyc_number, col_09_borrower_name, col_10_borrower_registration_date,
               col_11_borrower_registration_number, col_12_individual_pan, col_13_individual_ckyc_number,
               col_14_gender, col_15_individual_name, col_16_father_mother_name, col_17_date_of_birth,
               col_18_mobile_number, col_19_email] + ADDRESS_COLUMNS
    return build_row(FIELDS, columns, c, b, serial, cleared=c.cleared_for(b))
