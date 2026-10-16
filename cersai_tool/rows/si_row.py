"""SI row: one per security interest, i.e. one per collateral. 15 columns
(sheet "si" of the format description)."""
import re

from .. import settings
from ..model import (SRC_LMS, SRC_PDF, Cell, blank, conflict, derived, fixed, invalid,
                     manual, missing, not_applicable, ok)
from ..rules import dates
from ..rules.text import DESCRIPTION, as_text, clean
from ..sources.title_report import mortgage_type_id

FIELDS = ["ROW_TYPE", "SERIAL_NUMBER", "NUMBER_OF_BORROWERS", "NUMBER_OF_ASSET_OWNERS",
          "NUMBER_OF_CONSORTIUM_MEMBERS", "SI_TYPE_ID", "SI_TYPE_OTHERS", "FINANCING_TYPE_ID",
          "SI_CREATION_DATE", "TOTAL_SECURED_AMOUNT", "ENTITY_MIS_TOKEN", "NARRATION", "TYPE_OF_CHARGE",
          "IS_TPM", "BATCH_REF_NUMBER"]


def _digits(text, count=8):
    return re.sub(r"\D", "", text or "")[-count:].rjust(count, "0")


# Column 1 - ROW_TYPE
# Always the text "SI".
def col_01_row_type(c, row):
    return fixed("ROW_TYPE", "SI")


# Column 2 - SERIAL_NUMBER
# The position of this collateral in the batch: 1 for the first, 2 for the
# second, and so on. Unique within the file. When step 3 leaves some SIs out of
# the .DAT file, it renumbers the rest 1, 2, 3...
def col_02_serial_number(c, row):
    return derived("SERIAL_NUMBER", c.serial, "position in the batch")


# Column 3 - NUMBER_OF_BORROWERS
# Where it comes from: CKYC. The number of different customers (Neo CIF ID)
#   on this collateral's application, not counting guarantors (a guarantor is
#   not a borrower).
# Rules: it always equals the number of BOR rows written below this SI.
# If it can't be found: an application with no CKYC rows has no borrowers;
#   the SI is reported as missing its borrowers and cannot be filed.
def col_03_number_of_borrowers(c, row):
    count = len(c.borrowers)
    if count == 0:
        return missing("NUMBER_OF_BORROWERS", "no borrowers found in CKYC for this application")
    return derived("NUMBER_OF_BORROWERS", count, "borrowers in CKYC, guarantors excluded")


# Column 4 - NUMBER_OF_ASSET_OWNERS
# Where it comes from: the property owners, compared with the borrowers.
#   The owners come from COLLATERAL "Owner Name" (each person once). Their
#   "Percentage Share" should add up to 100% and "Ownership Status" should be
#   Single for one owner and Joint for more: if not, the collateral is listed
#   in the Check and correct sheet. A name cell holding two people ("KALAISELVI
#   AND RADHAKRISHNAN SANKARANARAYANAN") is taken as two owners when each is a
#   borrower, and is otherwise listed there too (company names contain "and").
#   The title report (item 4) is the main source only when COLLATERAL has no
#   owners. Otherwise it checks them: when it names other owners, both lists
#   are shown in the Check and correct sheet and the COLLATERAL owners are used.
#   Only when the COLLATERAL shares show an owner is missing (they add up to
#   less than 100%, or Joint with one owner) and the title report names every
#   COLLATERAL owner plus others, those others are added.
# Rules: each owner is matched to a borrower by name (titles like Mr/Mrs/Smt,
#   a father's or husband's name after S/O, D/O or W/O, and spelling
#   differences are ignored; initials match whole names). This
#   column counts the owners who are NOT borrowers; 0 when every owner is a
#   borrower. Each borrower is linked to one owner of the property at most
#   (owners listed separately are different people). Similar names (a spelling
#   difference, a missing surname) are taken to be the same person; names whose
#   given names differ are taken to be different people, even when the surname
#   and the father's or husband's name are the same ("DEEPA RAVIRAJ HIREMATH",
#   a wife, is not "RAVIRAJ HIREMATH"). These uncertain decisions are listed in
#   the Check and correct sheet, where a Y or N changes them.
#   With several collaterals on one application, all this is worked out for
#   each collateral with its own owners.
# If it can't be found: no owners in COLLATERAL or a usable title report ->
#   reported as missing; owner names can be typed in the required-details file.
def col_04_number_of_asset_owners(c, row):
    typed = c.ov("NUMBER_OF_ASSET_OWNERS")
    if typed is not None:
        return manual("NUMBER_OF_ASSET_OWNERS", as_text(typed))
    decision = c.owners
    if decision.status == "missing":
        return missing("NUMBER_OF_ASSET_OWNERS", decision.note)
    count = len(c.non_borrower_owners())
    note = f"owners: {'; '.join(decision.owners)} ({decision.note})"
    uncertain = [link for link in c.owner_links if link.uncertain]
    if uncertain:
        note += "; decided on similar names: " + "; ".join(
            f"'{l.owner}' = '{l.borrower_name}'" if l.neo_cif else f"'{l.owner}' is not '{l.near_name}'"
            for l in uncertain)
    for problem in (decision.count_problem, decision.split_note):
        if problem:
            note += f"; {problem}"
    if decision.status == "differs":
        return conflict("NUMBER_OF_ASSET_OWNERS", count, decision.source,
                        "; ".join(decision.lms_owners), "; ".join(decision.pdf_owners), note)
    source = decision.source or SRC_LMS
    return Cell("NUMBER_OF_ASSET_OWNERS", str(count), source, "ok", note)


# Column 5 - NUMBER_OF_CONSORTIUM_MEMBERS
# Always 0: ARGFL lends alone (financing type "Sole").
def col_05_number_of_consortium_members(c, row):
    return fixed("NUMBER_OF_CONSORTIUM_MEMBERS", settings.CONSORTIUM_MEMBERS)


# Column 6 - SI_TYPE_ID (type of mortgage)
# Where it comes from: the title report, item 8 "Nature of Mortgage". The LMS
#   files do not hold the type of mortgage, so the title report is always used.
# Rules: the words are turned into the CERSAI code (SI_SUB_TYPE_ID master,
#   immovable rows only): Equitable -> 17, Registered -> 18, Agreement to create
#   mortgage -> 19, Conditional sale -> 20, anything else -> 21 (Other).
#   Never read from the report's "Final Certificate", which one law firm
#   strikes out while a search is pending.
# If it can't be found: no usable title report or no item 8 -> left blank and
#   asked for in the required-details file (it is mandatory, so the SI stays out
#   of the .DAT file until it is typed). No type of mortgage is assumed.
def col_06_si_type_id(c, row):
    typed = c.ov("SI_TYPE_ID")
    if typed is not None:
        code = as_text(typed)
        if code not in c.masters.immovable_si_types:
            return invalid("SI_TYPE_ID", code, "not an immovable SI type (17-21)")
        return manual("SI_TYPE_ID", code)
    report = c.report
    text = report.get("mortgage_text", "") if report else ""
    code = mortgage_type_id(text)
    if code:
        return ok("SI_TYPE_ID", code, SRC_PDF, f"title report item 8: '{text}'")
    why = "no usable title report" if not report else "title report item 8 not found"
    return missing("SI_TYPE_ID", f"{why}: type the type of mortgage (17 Equitable, 18 Registered, ...)")


# Column 7 - SI_TYPE_OTHERS
# Only used when column 6 is 21 (Other): a short description of the mortgage.
# Where it comes from: the wording of title report item 8.
# If it can't be found: reported as missing (it is mandatory when column 6 is 21).
def col_07_si_type_others(c, row):
    if row["SI_TYPE_ID"].value != "21":
        return not_applicable("SI_TYPE_OTHERS", "only when the SI type is 21 (Other)")
    typed = c.ov("SI_TYPE_OTHERS")
    if typed is not None:
        return manual("SI_TYPE_OTHERS", clean(typed, DESCRIPTION, 100))
    text = c.report.get("mortgage_text", "") if c.report else ""
    if text:
        return ok("SI_TYPE_OTHERS", clean(text, DESCRIPTION, 100), SRC_PDF, "title report item 8")
    return missing("SI_TYPE_OTHERS", "describe the type of mortgage (SI type is 21 Other)")


# Column 8 - FINANCING_TYPE_ID
# Always 1 (Sole).
def col_08_financing_type_id(c, row):
    return fixed("FINANCING_TYPE_ID", settings.FINANCING_TYPE_ID)


# Column 9 - SI_CREATION_DATE
# Where it comes from: CKYC "First Disbursal Date" (DD-MM-YYYY), turned into
#   YYYY-MM-DD. If that is empty, COLLATERAL "Disbursal Date" (DD-MON-YY).
# Rules: must not be in the future. The title reports do not contain this date.
# If it can't be found: reported as missing.
def col_09_si_creation_date(c, row):
    typed = c.ov("SI_CREATION_DATE")
    if typed is not None:
        value = dates.to_cersai(typed, dates.PAST)
        return manual("SI_CREATION_DATE", value) if value else invalid("SI_CREATION_DATE", typed, "not a date")
    for raw, label in ((c.bundle["first_disbursal_date"], "CKYC First Disbursal Date"),
                       (c.bundle["collateral_disbursal_date"], "COLLATERAL Disbursal Date")):
        value = dates.to_cersai(raw, dates.PAST)
        if value:
            return ok("SI_CREATION_DATE", value, SRC_LMS, label)
    return missing("SI_CREATION_DATE", "no disbursal date in CKYC or COLLATERAL")


# Column 10 - TOTAL_SECURED_AMOUNT
# Where it comes from: COLLATERAL "Sanctioned Loan Amount" of the active loan.
# Rules: digits only (commas removed), at most 2 decimals, greater than 0.
#   Every active row of a collateral should show the same amount.
# If it can't be found: no amount, or active rows with different amounts ->
#   left blank and asked for in the required-details file (no amount is
#   assumed).
def col_10_total_secured_amount(c, row):
    typed = c.ov("TOTAL_SECURED_AMOUNT")
    amounts = [_amount(a) for a in c.bundle["sanctioned_amounts"]]
    amounts = [a for a in amounts if a]
    if typed is not None:
        value = _amount(typed)
        return manual("TOTAL_SECURED_AMOUNT", value) if value else invalid(
            "TOTAL_SECURED_AMOUNT", typed, "not an amount greater than 0")
    if not amounts:
        return missing("TOTAL_SECURED_AMOUNT", "no sanctioned amount in COLLATERAL")
    if len(set(amounts)) > 1:
        return missing("TOTAL_SECURED_AMOUNT",
                       f"active rows show different sanctioned amounts ({', '.join(amounts)}): type the secured amount")
    return ok("TOTAL_SECURED_AMOUNT", amounts[0], SRC_LMS, "COLLATERAL Sanctioned Loan Amount")


def _amount(value) -> str:
    text = as_text(value).replace(",", "")
    if not re.fullmatch(r"\d+(\.\d+)?", text):
        return ""
    number = round(float(text), 2)
    if number <= 0:
        return ""
    result = f"{number:.2f}".rstrip("0").rstrip(".")
    return result if len(result) <= 13 else ""


# Column 11 - ENTITY_MIS_TOKEN
# Where it comes from: the application number and the Collateral ID.
# Rules: CERSAI allows at most 20 letters/digits, so the two are shortened to
#   "A" + last 8 digits of the application + "C" + last 8 digits of the
#   collateral, e.g. APPL00027326 + COLL00266688 -> A00027326C00266688.
def col_11_entity_mis_token(c, row):
    token = f"A{_digits(c.application_number)}C{_digits(c.collateral_id)}"
    return derived("ENTITY_MIS_TOKEN", token, "A + application digits + C + collateral digits")


# Column 12 - NARRATION
# Optional free text. Left blank, unless the setting INCLUDE_LAN_IN_NARRATION
# is switched on: then "LAN <loan account number>" from COLLATERAL.
def col_12_narration(c, row):
    typed = c.ov("NARRATION")
    if typed is not None:
        return manual("NARRATION", clean(typed, DESCRIPTION, 100))
    if settings.INCLUDE_LAN_IN_NARRATION and c.bundle["loan_account"]:
        return ok("NARRATION", clean(f"LAN {c.bundle['loan_account']}", DESCRIPTION, 100), SRC_LMS,
                  "COLLATERAL Loan Account #")
    return blank("NARRATION", "optional")


# Column 13 - TYPE_OF_CHARGE
# "FIRST CHARGE" (settings.TYPE_OF_CHARGE), unless a different charge is typed
# in the required-details file.
# Rules: letters and spaces only, max 20.
# If it can't be found: when the same property (same COLLATERAL "Global
#   collateral No") also secures another active loan, the charge is not known ->
#   left blank and asked for (first, second, pari passu...). No charge is assumed.
def col_13_type_of_charge(c, row):
    typed = c.ov("TYPE_OF_CHARGE")
    if typed is not None:
        return manual("TYPE_OF_CHARGE", clean(typed, "", 20, letters_only=True, upper=True))
    shared = c.bundle.get("shared_with_applications") or []
    if shared:
        return missing("TYPE_OF_CHARGE", f"this property also secures application(s) {', '.join(shared)}: type "
                                         "the type of charge (FIRST CHARGE, SECOND CHARGE, PARI PASSU...)")
    return fixed("TYPE_OF_CHARGE", settings.TYPE_OF_CHARGE)


# Column 14 - IS_TPM (third-party mortgage)
# Where it comes from: column 4. 1 when at least one owner is not a borrower,
#   otherwise 0.
# Rules: a third-party mortgage also needs asset-owner (ASO) rows, which this
#   tool does not write yet, so such an SI is left out of the .DAT file and
#   listed for manual filing.
def col_14_is_tpm(c, row):
    owners_cell = row["NUMBER_OF_ASSET_OWNERS"]
    count = int(owners_cell.value) if owners_cell.value.isdigit() else 0
    if count > 0:
        c.issue("warning", "Third-party mortgage",
                f"{count} owner(s) are not borrowers: {', '.join(c.non_borrower_owners())}. "
                "Asset-owner rows are needed, so this SI will be left out of the .DAT file.")
    return derived("IS_TPM", "1" if count > 0 else "0", "1 when an owner is not a borrower")


# Column 15 - BATCH_REF_NUMBER
# Optional. Exactly 20 letters/digits: "ARGF" + file date (YYYYMMDD) + last 8
# digits of the Collateral ID. Step 3 recalculates it with the date the .DAT
# file is written. Blank if the setting GENERATE_BATCH_REF_NUMBER is off.
def col_15_batch_ref_number(c, row):
    if not settings.GENERATE_BATCH_REF_NUMBER:
        return blank("BATCH_REF_NUMBER", "optional")
    ref = f"{settings.BATCH_REF_PREFIX}{c.run_date.strftime('%Y%m%d')}{_digits(c.collateral_id)}"
    return derived("BATCH_REF_NUMBER", ref, "ARGF + file date + collateral digits")


COLUMNS = [col_01_row_type, col_02_serial_number, col_03_number_of_borrowers,
           col_04_number_of_asset_owners, col_05_number_of_consortium_members, col_06_si_type_id,
           col_07_si_type_others, col_08_financing_type_id, col_09_si_creation_date,
           col_10_total_secured_amount, col_11_entity_mis_token, col_12_narration,
           col_13_type_of_charge, col_14_is_tpm, col_15_batch_ref_number]


def build(c):
    from . import build_row
    return build_row(FIELDS, COLUMNS, c, cleared=c.cleared())
