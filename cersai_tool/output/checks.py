"""Step 3's checks: is each SI in the final Excel ready for CERSAI?

Every value is checked against the CERSAI batch format description (the
field's characters, length, date or code list), then the fields and rows of
each SI are checked against each other (for example, NUMBER_OF_BORROWERS must
equal the number of BOR rows). An SI with any error is left out of the .DAT
file and listed with its reasons; warnings are listed but do not stop an SI.

SIs are also left out when they are a third-party mortgage (asset-owner rows
are not supported yet), when they wait for an answer (held in step 2: a title
report that may belong to another property), or when they are already in an
earlier .DAT file.
"""
import re
from dataclasses import dataclass, field

from .. import settings
from ..rows import bor_row, chg_row, imm_row, si_row
from ..rules.dates import check_cersai_date
from ..rules.identifiers import CIN_RE, CKYC_RE, EMAIL_RE, LLPIN_RE, MOBILE_RE, PAN_RE
from ..rules.text import ADDRESS, DESCRIPTION, HOUSE, ORG_NAME, STREET, break_rules

STRUCTURE_SUB_TYPES = imm_row.STRUCTURE_SUB_TYPES
INDIVIDUAL_FIELDS = bor_row.FIELDS[11:19]          # fields 12-19
NON_INDIVIDUAL_FIELDS = bor_row.FIELDS[6:11]       # fields 7-11


@dataclass
class BlockResult:
    collateral_id: str
    application_number: str
    block: dict
    errors: list = field(default_factory=list)       # [(row label, field, problem)]
    warnings: list = field(default_factory=list)
    left_out_reason: str = ""                        # set when left out for a reason other than errors

    @property
    def included(self) -> bool:
        return not self.errors and not self.left_out_reason


# --------------------------------------------------------------------------
# One value
# --------------------------------------------------------------------------
def _text(specials, max_len, letters_only=False):
    def rule(value):
        return "; ".join(break_rules(value, specials, max_len, letters_only))
    return rule


def _regex(pattern, what):
    def rule(value):
        return "" if re.fullmatch(pattern, value) else f"must be {what}"
    return rule


def _one_of(*allowed):
    def rule(value):
        return "" if value in allowed else f"must be {' or '.join(allowed)}"
    return rule


def _code(get_codes, master):
    def rule(value, masters):
        return "" if value in get_codes(masters) else f"'{value}' is not in the CERSAI {master} master"
    return rule


def _whole(low, high):
    def rule(value):
        if not re.fullmatch(r"-?\d+", value) or not low <= int(value) <= high:
            return f"must be a whole number from {low} to {high}"
        return ""
    return rule


def _date(value):
    return check_cersai_date(value)


def _empty(value):
    return "must be empty"


def _amount(value):
    if not re.fullmatch(r"\d+(\.\d{1,2})?", value):
        return "must be a number with at most 2 decimals (no commas or spaces)"
    if float(value) <= 0:
        return "must be greater than zero"
    if len(value) > 13:
        return "longer than 13 characters"
    return ""


def _survey(value):
    if len(value) > 1000:
        return "longer than 1000 characters"
    return ""


# field -> (mandatory?, rule). Rules taking masters are marked with a tuple.
CODES = {
    "SI_TYPE_ID": _code(lambda m: m.immovable_si_types, "SI type (immovable)"),
    "FINANCING_TYPE_ID": _code(lambda m: m.financing_types, "FINANCING_TYPE"),
    "ASSET_SUB_TYPE_ID": _code(lambda m: m.asset_sub_types, "ASSET_SUB_TYPE"),
    "ASSET_AREA_UNIT": _code(lambda m: m.area_units, "AREA_UNIT"),
    "BORROWER_TYPE": _code(lambda m: m.borrower_types, "BORROWER_ASSET_OWNER_TYPE"),
    "GENDER": _code(lambda m: m.genders, "GENDER"),
    "CITY": _code(lambda m: m.cities, "CITY"),
    "DISTRICT": _code(lambda m: m.districts, "DISTRICT"),
    "STATE": _code(lambda m: m.states, "STATE"),
    "PIN_CODE": _code(lambda m: m.pin_districts, "PIN_CODE"),
    "COUNTRY": _code(lambda m: m.countries, "COUNTRY"),
}

SI_RULES = {
    "ROW_TYPE": (True, _one_of("SI")),
    "NUMBER_OF_BORROWERS": (True, _whole(1, 999)),
    "NUMBER_OF_ASSET_OWNERS": (True, _whole(0, 999)),
    "NUMBER_OF_CONSORTIUM_MEMBERS": (True, _whole(0, 999)),
    "SI_TYPE_ID": (True, "code"),
    "SI_TYPE_OTHERS": (False, _text(DESCRIPTION, 100)),
    "FINANCING_TYPE_ID": (True, "code"),
    "SI_CREATION_DATE": (True, _date),
    "TOTAL_SECURED_AMOUNT": (True, _amount),
    "ENTITY_MIS_TOKEN": (True, _regex(r"[A-Za-z0-9]{1,20}", "letters and digits only, at most 20")),
    "NARRATION": (False, _text(DESCRIPTION, 100)),
    "TYPE_OF_CHARGE": (True, _text("", 20, letters_only=True)),
    "IS_TPM": (True, _one_of("0", "1")),
    "BATCH_REF_NUMBER": (False, _regex(r"[A-Za-z0-9]{20}", "exactly 20 letters and digits")),
}
IMM_RULES = {
    "ROW_TYPE": (True, _one_of("IMM")),
    "ASSET_CATEGORY_ID": (True, _one_of("1")),
    "ASSET_TYPE_ID": (True, _one_of("1", "2", "3")),
    "ASSET_TYPE_OTHERS": (False, _text(DESCRIPTION, 50)),
    "ASSET_SUB_TYPE_ID": (True, "code"),
    "ASSET_DESCRIPTION": (True, _text(DESCRIPTION, 100)),
    "SURVEY_NUMBER": (False, _survey),
    "PLOT_NUMBER": (False, _text(ADDRESS, 20)),
    "ASSET_AREA": (False, _regex(r"\d{1,5}(\.\d{1,2})?", "a number up to 99999.99")),
    "ASSET_AREA_UNIT": (False, "code"),
    "HOUSE_NUMBER": (False, _text(HOUSE, 20)),
    "FLOOR_NUMBER": (False, _whole(-3, 100)),
    "BUILDING_NAME": (False, _text(ORG_NAME, 50)),
    "PROJECT_NAME": (True, _text(ORG_NAME, 50)),
    "STREET_NAME": (False, _text(STREET, 50)),
    "POCKET": (False, _text(ADDRESS, 20)),
    "LOCALITY": (True, _text(ADDRESS, 50)),
    "CITY": (True, "code"), "DISTRICT": (True, "code"), "STATE": (True, "code"),
    "PIN_CODE": (True, "code"), "COUNTRY": (True, "code"),
}
BOR_RULES = {
    "ROW_TYPE": (True, _one_of("BOR")),
    "BORROWER_TYPE": (True, "code"),
    "IS_ASSET_OWNER": (True, _one_of("0", "1")),
    "BORROWER_UID_TYPE": (False, _empty),
    "BORROWER_UID_VALUE": (False, _empty),
    "BORROWER_PAN": (False, _regex(PAN_RE.pattern.strip("^$"), "a PAN (5 letters, 4 digits, 1 letter)")),
    "BORROWER_CKYC_NUMBER": (False, _regex(CKYC_RE.pattern.strip("^$"), "14 digits")),
    "BORROWER_NAME": (False, _text(ORG_NAME, 100)),
    "BORROWER_REGISTRATION_DATE": (False, _date),
    "INDIVIDUAL_PAN": (False, _regex(PAN_RE.pattern.strip("^$"), "a PAN (5 letters, 4 digits, 1 letter)")),
    "INDIVIDUAL_CKYC_NUMBER": (False, _regex(CKYC_RE.pattern.strip("^$"), "14 digits")),
    "GENDER": (False, "code"),
    "INDIVIDUAL_NAME": (False, _text("", 50, letters_only=True)),
    "FATHER_MOTHER_NAME": (False, _text("", 50, letters_only=True)),
    "DATE_OF_BIRTH": (False, _date),
    "MOBILE_NUMBER": (False, _regex(MOBILE_RE.pattern.strip("^$"), "10 digits")),
    "EMAIL": (False, lambda v: "" if len(v) <= 50 and EMAIL_RE.match(v) else
              "not a valid email for CERSAI (letters, digits, _ and . only, one @, max 50)"),
    "ADDRESS_LINE_1": (True, _text(ADDRESS, 100)),
    "ADDRESS_LINE_2": (False, _text(ADDRESS, 100)),
    "ADDRESS_LINE_3": (False, _text(ADDRESS, 100)),
    "CITY": (True, "code"), "DISTRICT": (True, "code"), "STATE": (True, "code"),
    # the batch format description gives "IN" here, not a COUNTRY master code (the IMM row uses IND)
    "PIN_CODE": (True, "code"), "COUNTRY": (True, _one_of(settings.BOR_CHG_COUNTRY)),
}
CHG_RULES = {
    "ROW_TYPE": (True, _one_of("CHG")),
    "OFFICE_CODE": (False, _empty),
    "OFFICE_NAME": (True, _text("", 100, letters_only=True)),
    "ADDRESS_LINE_1": (True, _text(ADDRESS, 100)),
    "ADDRESS_LINE_2": (False, _text(ADDRESS, 100)),
    "ADDRESS_LINE_3": (False, _text(ADDRESS, 100)),
    "CITY": (True, "code"), "DISTRICT": (True, "code"), "STATE": (True, "code"),
    "PIN_CODE": (True, "code"), "COUNTRY": (True, _one_of(settings.BOR_CHG_COUNTRY)),
}
RULES = {"SI": (si_row.FIELDS, SI_RULES), "IMM": (imm_row.FIELDS, IMM_RULES), "BOR": (bor_row.FIELDS, BOR_RULES),
         "CHG": (chg_row.FIELDS, CHG_RULES)}


# what to do when these mandatory values are empty
EMPTY_MESSAGES = {
    ("SI", "NUMBER_OF_ASSET_OWNERS"): "is empty: no owners are known for this collateral. Type the owners (OWNERS) in "
                                      "the required-details file and run step 2 again",
    ("SI", "NUMBER_OF_BORROWERS"): "is empty: CKYC has no borrowers for this application",
    ("SI", "SI_TYPE_ID"): "is empty: no usable title report gives the type of mortgage. Type it (SI_TYPE_ID) in "
                          "the required-details file and run step 2 again",
    ("SI", "TOTAL_SECURED_AMOUNT"): "is empty: no sanctioned amount, or the active loan rows show different ones. "
                                    "Type it in the required-details file and run step 2 again",
    ("SI", "TYPE_OF_CHARGE"): "is empty: the property also secures another loan, so the type of charge must be "
                              "typed in the required-details file. Run step 2 again after",
    ("IMM", "ASSET_TYPE_ID"): "is empty: the asset type could not be worked out. Type it (ASSET_TYPE_ID) in the "
                              "required-details file and run step 2 again",
    ("IMM", "ASSET_SUB_TYPE_ID"): "is empty: fill COLLATERAL Structure Type, or type the sub-type (ASSET_SUB_TYPE_ID) "
                                  "in the required-details file, and run step 2 again",
}


def check_values(kind, values, masters, label, errors):
    fields, rules = RULES[kind]
    row = dict(zip(fields, values))
    for name, value in row.items():
        if "|" in value:
            errors.append((label, name, "contains '|', which separates the fields of the .DAT file"))
            continue
        if not value.isascii():
            errors.append((label, name, "contains a character that is not plain English text (non-ASCII)"))
            continue
        mandatory, rule = rules.get(name, (False, None))
        if not value:
            if mandatory:
                errors.append((label, name, EMPTY_MESSAGES.get((kind, name), "is empty but mandatory")))
            continue
        if rule == "code":
            problem = CODES[name](value, masters)
        elif rule is None:
            problem = ""
        else:
            problem = rule(value)
        if problem:
            errors.append((label, name, problem))
    return row


def check_geo(row, masters, label, errors):
    """PIN, district, state and city must agree with each other."""
    pin, district, state, city = row["PIN_CODE"], row["DISTRICT"], row["STATE"], row["CITY"]
    if pin in masters.pin_districts and district in masters.districts \
            and district not in masters.pin_districts[pin]:
        names = ", ".join(f"{d} {masters.districts.get(d, ('?', ''))[0]}" for d in masters.pin_districts[pin])
        errors.append((label, "DISTRICT", f"district {district} is not a district of PIN {pin} ({names})"))
    if district in masters.districts and state in masters.states and masters.state_of_district(district) != state:
        errors.append((label, "STATE", f"state {state} is not the state of district {district} "
                                       f"({masters.state_of_district(district)})"))
    if city in masters.cities and district in masters.districts and masters.district_of_city(city) != district:
        errors.append((label, "CITY", f"city {city} ({masters.cities[city][0]}) is in district "
                                      f"{masters.district_of_city(city)}, not {district}"))


# --------------------------------------------------------------------------
# One SI block
# --------------------------------------------------------------------------
def check_block(block, masters) -> BlockResult:
    result = BlockResult(block["collateral_id"], block["application_number"], block)
    errors, warnings = result.errors, result.warnings
    rows = block["rows"]
    kinds = [r["type"] for r in rows]
    if not result.collateral_id:
        errors.append(("SI", "", "no Collateral ID in column AD of the SI row"))
    expected = ["SI", "IMM"] + ["BOR"] * kinds.count("BOR") + ["CHG"]
    if kinds != expected or kinds.count("BOR") == 0:
        errors.append(("SI", "", f"the block's rows are {'-'.join(kinds)}; expected SI, IMM, one or more BOR, "
                                 "then CHG"))
        return result

    si = check_values("SI", rows[0]["values"], masters, "SI", errors)
    imm = check_values("IMM", rows[1]["values"], masters, "IMM", errors)
    bors = []
    for number, row in enumerate(rows[2:-1], start=1):
        label = f"BOR {number}" + (f" ({row['neo_cif']})" if row["neo_cif"] else "")
        bors.append((label, check_values("BOR", row["values"], masters, label, errors)))
    check_values("CHG", rows[-1]["values"], masters, "CHG", errors)
    chg = dict(zip(chg_row.FIELDS, rows[-1]["values"]))

    # SI across fields and rows
    if si["NUMBER_OF_BORROWERS"].isdigit() and int(si["NUMBER_OF_BORROWERS"]) != len(bors):
        errors.append(("SI", "NUMBER_OF_BORROWERS", f"says {si['NUMBER_OF_BORROWERS']} but the block has "
                                                    f"{len(bors)} BOR row(s)"))
    owners = int(si["NUMBER_OF_ASSET_OWNERS"]) if si["NUMBER_OF_ASSET_OWNERS"].isdigit() else None
    if owners is not None and si["IS_TPM"] in ("0", "1") and (owners > 0) != (si["IS_TPM"] == "1"):
        errors.append(("SI", "IS_TPM", "must be 1 exactly when NUMBER_OF_ASSET_OWNERS is above 0"))
    if si["IS_TPM"] == "1" or (owners or 0) > 0:
        result.left_out_reason = ("third-party mortgage: the owners who are not borrowers need asset-owner (ASO) "
                                  "rows, which this tool does not write; file this SI by hand")
    if owners == 0 and not any(b["IS_ASSET_OWNER"] == "1" for _, b in bors):
        errors.append(("SI", "NUMBER_OF_ASSET_OWNERS", "no owner: no borrower is marked IS_ASSET_OWNER = 1 and "
                                                       "NUMBER_OF_ASSET_OWNERS is 0"))
    if si["FINANCING_TYPE_ID"] == "1" and si["NUMBER_OF_CONSORTIUM_MEMBERS"] not in ("", "0"):
        errors.append(("SI", "NUMBER_OF_CONSORTIUM_MEMBERS", "must be 0 when FINANCING_TYPE_ID is 1 (Sole)"))
    if si["SI_TYPE_ID"] == "21" and not si["SI_TYPE_OTHERS"]:
        errors.append(("SI", "SI_TYPE_OTHERS", "is mandatory when SI_TYPE_ID is 21 (Other)"))
    if si["SI_TYPE_ID"] != "21" and si["SI_TYPE_OTHERS"]:
        errors.append(("SI", "SI_TYPE_OTHERS", "must be empty unless SI_TYPE_ID is 21 (Other)"))

    # IMM across fields
    if imm["ASSET_TYPE_ID"] == "3" and not imm["ASSET_TYPE_OTHERS"]:
        errors.append(("IMM", "ASSET_TYPE_OTHERS", "is mandatory when ASSET_TYPE_ID is 3 (Other)"))
    if imm["ASSET_TYPE_ID"] in ("1", "2") and imm["ASSET_TYPE_OTHERS"]:
        errors.append(("IMM", "ASSET_TYPE_OTHERS", "must be empty unless ASSET_TYPE_ID is 3 (Other)"))
    sub = imm["ASSET_SUB_TYPE_ID"]
    if sub in masters.asset_sub_types and masters.asset_sub_types[sub][0] != imm["ASSET_TYPE_ID"]:
        errors.append(("IMM", "ASSET_SUB_TYPE_ID", f"sub-type {sub} does not belong to asset type "
                                                   f"{imm['ASSET_TYPE_ID']}"))
    if not imm["SURVEY_NUMBER"] and not imm["PLOT_NUMBER"]:
        errors.append(("IMM", "SURVEY_NUMBER", "SURVEY_NUMBER or PLOT_NUMBER must be filled"))
    if imm["ASSET_AREA"] and not imm["ASSET_AREA_UNIT"]:
        errors.append(("IMM", "ASSET_AREA_UNIT", "is mandatory when ASSET_AREA is filled"))
    if not imm["ASSET_AREA"] and imm["ASSET_AREA_UNIT"]:
        warnings.append(("IMM", "ASSET_AREA_UNIT", "filled although ASSET_AREA is empty"))
    if sub in STRUCTURE_SUB_TYPES and not imm["HOUSE_NUMBER"]:
        errors.append(("IMM", "HOUSE_NUMBER", f"is mandatory for sub-type {sub} (a structure)"))
    if imm["ASSET_UNIQUE_ID"].upper() in ("NA", "N/A", "NIL"):
        warnings.append(("IMM", "ASSET_UNIQUE_ID", f"'{imm['ASSET_UNIQUE_ID']}' would be stored as the asset's ID"))
    check_geo(imm, masters, "IMM", errors)

    # BOR across fields
    for label, bor in bors:
        individual = bor["BORROWER_TYPE"] == "IND"
        unused = NON_INDIVIDUAL_FIELDS if individual else INDIVIDUAL_FIELDS
        for name in unused:
            if bor[name]:
                errors.append((label, name, f"must be empty for borrower type {bor['BORROWER_TYPE'] or '?'}"))
        if individual:
            if not bor["INDIVIDUAL_PAN"] and not bor["INDIVIDUAL_CKYC_NUMBER"]:
                errors.append((label, "INDIVIDUAL_PAN", "an individual needs a PAN or a CKYC number"))
            for name in ("GENDER", "INDIVIDUAL_NAME", "FATHER_MOTHER_NAME", "DATE_OF_BIRTH"):
                if not bor[name]:
                    errors.append((label, name, "is mandatory for an individual"))
        elif bor["BORROWER_TYPE"]:
            if not bor["BORROWER_PAN"] and not bor["BORROWER_CKYC_NUMBER"]:
                errors.append((label, "BORROWER_PAN", "a non-individual needs a PAN or a CKYC number"))
            for name in ("BORROWER_NAME", "BORROWER_REGISTRATION_DATE"):
                if not bor[name]:
                    errors.append((label, name, f"is mandatory for borrower type {bor['BORROWER_TYPE']}"))
            number = bor["BORROWER_REGISTRATION_NUMBER"]
            if bor["BORROWER_TYPE"] == "COM" and not CIN_RE.match(number):
                errors.append((label, "BORROWER_REGISTRATION_NUMBER", "a company needs its 21-character CIN"))
            elif bor["BORROWER_TYPE"] == "LLP" and not LLPIN_RE.match(number):
                errors.append((label, "BORROWER_REGISTRATION_NUMBER", "an LLP needs its 7-character LLPIN"))
            elif bor["BORROWER_TYPE"] not in ("COM", "LLP") and number:
                warnings.append((label, "BORROWER_REGISTRATION_NUMBER",
                                 f"only used for companies and LLPs, not {bor['BORROWER_TYPE']}"))
        check_geo(bor, masters, label, errors)
    check_geo(chg, masters, "CHG", errors)
    return result


# --------------------------------------------------------------------------
# The whole batch
# --------------------------------------------------------------------------
def check_batch(sheet, masters, held: dict, register: dict, filed_status: str) -> list:
    """Checks every block of the final Excel. Returns a BlockResult per block.
    held: {collateral ID: why it waits for an answer} as recorded by step 2."""
    results = []
    seen_collaterals, seen_tokens = set(), set()
    for block in sheet["blocks"]:
        result = check_block(block, masters)
        coll = result.collateral_id
        token = block["rows"][0]["values"][si_row.FIELDS.index("ENTITY_MIS_TOKEN")] if block["rows"] else ""
        entry = register.get(coll)
        if entry and entry.get("Status") == filed_status:
            result.left_out_reason = f"already filed in {entry.get('DAT file')} on {entry.get('Date')}"
        elif coll in held:
            result.left_out_reason = (f"held: {held[coll]}. Answer in the required-details file (Check and "
                                      "correct) and run step 2 again")
        if coll in seen_collaterals:
            result.errors.append(("SI", "", f"collateral {coll} appears twice in the file; only the first is used"))
        if token and token in seen_tokens:
            result.errors.append(("SI", "ENTITY_MIS_TOKEN", f"{token} is used by another SI in this file"))
        seen_collaterals.add(coll)
        seen_tokens.add(token)
        results.append(result)
    return results
