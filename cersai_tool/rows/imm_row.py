"""IMM row: the immovable property (one per SI). 27 columns (sheet
"immovable_asset" of the format description).

This file holds columns 1-17 (what the asset is). Columns 18-27 (where it is)
are in imm_location.py.

Property details come first from the LMS: COLLATERAL "Property Type",
"Carpert Area", the hand-filled columns "Structure Type", "Survey No." and
"Project Name" when the extract has them, and the pieces read from the
free-text "Full Address". The
title report (item 1 "Full Description of the property", item 2 "Nature of
immovable property") is read only for what the LMS does not give. When both
give a value and they disagree, the LMS value is used and the difference is
flagged side by side in the required-details file.
"""
import re

from .. import settings
from ..model import (SRC_LMS, SRC_PDF, blank, conflict, derived, fixed, invalid, manual, missing,
                     not_applicable, ok)
from ..rules.address import SURVEY_KEYS
from ..rules.text import DESCRIPTION, HOUSE, ORG_NAME, STREET, SURVEY, ADDRESS, as_text, clean, \
    normalise_for_compare

FIELDS = ["ROW_TYPE", "ASSET_CATEGORY_ID", "ASSET_TYPE_ID", "ASSET_TYPE_OTHERS", "ASSET_SUB_TYPE_ID",
          "ASSET_UNIQUE_ID", "ASSET_DESCRIPTION", "SURVEY_NUMBER", "PLOT_NUMBER", "ASSET_AREA",
          "ASSET_AREA_UNIT", "HOUSE_NUMBER", "FLOOR_NUMBER", "BUILDING_NAME", "PROJECT_NAME", "STREET_NAME",
          "POCKET", "LOCALITY", "CITY", "DISTRICT", "STATE", "PIN_CODE", "COUNTRY",
          "LATITUDE_LONGITUDE_1", "LATITUDE_LONGITUDE_2", "LATITUDE_LONGITUDE_3", "LATITUDE_LONGITUDE_4"]

# LMS "Property Type" -> CERSAI ASSET_TYPE_ID (1 Residential, 2 Commercial, 3 Other)
PROPERTY_TYPES = {"RESIDENTIAL": "1", "COMMERCIAL": "2", "INDUSTRIAL": "3", "MIXED USAGE": "3", "PG": "1"}
OTHER_TYPE_TEXT = {"INDUSTRIAL": "Industrial", "MIXED USAGE": "Mixed Usage", "PLOT": "Plot"}
# COLLATERAL "Structure Type" (letters only, lower case) -> the name in the ASSET_SUB_TYPE_ID master
STRUCTURE_TYPES = {"structurewithland": "Structure with Land", "structurewithoutland": "Structure without Land",
                   "landwithoutstructure": "Land without Structure", "unidentified": "Unidentified"}
# sub-types that need a house / flat / unit number
STRUCTURE_SUB_TYPES = {"1", "2", "4", "5", "7", "8"}
# area units -> square feet, only for comparing the LMS and title-report areas
TO_SQFT = {"6": 1.0, "7": 10.7639, "8": 9.0, "1": 43560.0, "5": 107639.0, "4": 1089.0, "2": 1076.39,
           "3": 435.6}


def _pdf_type(c) -> str:
    nature = (c.report or {}).get("nature", "").strip().lower()
    if nature.startswith("resid"):
        return "1"
    if nature.startswith("comm"):          # also the typo "Commerical"
        return "2"
    return ""


def _joined(numbers, sep):
    return sep.join(numbers)


def _same_numbers(a, b) -> bool:
    norm = lambda values: {re.sub(r"PT$", "", normalise_for_compare(v).replace(" ", "")) for v in values}
    return norm(a) == norm(b)


# Column 1 - ROW_TYPE
# Always the text "IMM".
def col_01_row_type(c, row):
    return fixed("ROW_TYPE", "IMM")


# Column 2 - ASSET_CATEGORY_ID
# Always 1 (Immovable).
def col_02_asset_category_id(c, row):
    return fixed("ASSET_CATEGORY_ID", "1")


# Column 3 - ASSET_TYPE_ID (1 Residential, 2 Commercial, 3 Other)
# Where it comes from: COLLATERAL "Property Type" (never blank):
#   RESIDENTIAL -> 1, COMMERCIAL -> 2, PG -> 1, INDUSTRIAL -> 3, MIXED USAGE -> 3.
#   PLOT: the plot's use decides, taken from title report item 2 ("Nature of
#   immovable property": Residential -> 1, Commercial -> 2).
# Rules: when the title report says Residential/Commercial and the LMS type
#   says the other, the LMS value is used and the difference is flagged.
# If it can't be found: a PLOT without a title report saying Residential or
#   Commercial, or an unknown property type -> left blank and asked for in the
#   required-details file (no type is assumed).
def col_03_asset_type_id(c, row):
    typed = c.ov("ASSET_TYPE_ID")
    if typed is not None:
        code = as_text(typed)
        if code not in ("1", "2", "3"):
            return invalid("ASSET_TYPE_ID", code, "must be 1, 2 or 3")
        return manual("ASSET_TYPE_ID", code)
    lms_type = c.bundle["property_type"]
    pdf_code = _pdf_type(c)
    if lms_type == "PLOT":
        if pdf_code:
            return ok("ASSET_TYPE_ID", pdf_code, SRC_PDF, f"plot; title report item 2 '{c.report['nature']}'")
        return missing("ASSET_TYPE_ID", "PLOT with no title report saying Residential or Commercial: type 1 "
                                        "(Residential), 2 (Commercial) or 3 (Other)")
    code = PROPERTY_TYPES.get(lms_type)
    if not code:
        return missing("ASSET_TYPE_ID", f"unknown COLLATERAL Property Type '{lms_type}': type 1 (Residential), "
                                        "2 (Commercial) or 3 (Other)")
    if pdf_code and code in ("1", "2") and pdf_code != code:
        return conflict("ASSET_TYPE_ID", code, SRC_LMS, f"{code} ({lms_type})",
                        f"{pdf_code} ({c.report['nature']})", "LMS property type and title report disagree")
    return ok("ASSET_TYPE_ID", code, SRC_LMS, f"COLLATERAL Property Type '{lms_type}'")


# Column 4 - ASSET_TYPE_OTHERS
# Only when column 3 is 3 (Other): a short description of the asset type,
# taken from COLLATERAL "Property Type" (e.g. "Industrial", "Mixed Usage",
# "Plot"). Blank otherwise. Max 50 characters.
def col_04_asset_type_others(c, row):
    if row["ASSET_TYPE_ID"].value != "3":
        return not_applicable("ASSET_TYPE_OTHERS", "only when the asset type is 3 (Other)")
    typed = c.ov("ASSET_TYPE_OTHERS")
    if typed is not None:
        return manual("ASSET_TYPE_OTHERS", clean(typed, DESCRIPTION, 50))
    lms_type = c.bundle["property_type"]
    text = OTHER_TYPE_TEXT.get(lms_type, lms_type.title())
    if text:
        return ok("ASSET_TYPE_OTHERS", clean(text, DESCRIPTION, 50), SRC_LMS, "from COLLATERAL Property Type")
    return missing("ASSET_TYPE_OTHERS", "describe the asset type (asset type is 3 Other)")


# Column 5 - ASSET_SUB_TYPE_ID
# Where it comes from: COLLATERAL "Structure Type", a column filled by hand,
#   when the extract has it and the cell is filled: "Structure with Land",
#   "Structure without Land", "Land without Structure" (or "Unidentified"), or
#   the code itself.
# Rules: the code depends on the asset type (column 3), from the
#   ASSET_SUB_TYPE_ID master: Residential 1/2/3, Commercial 4/5/6, Other 7/8/19.
#   It is never guessed from the address, the title report or the property
#   type: the unit words ("Flat", "Door No.", "Property bearing No.") proved
#   unreliable. Other wording in the column is reported, not guessed at.
# If it can't be found: left blank and asked for in the required-details file.
#   When the column names a structure but the asset type is not known yet, the
#   code is worked out once ASSET_TYPE_ID is typed.
def col_05_asset_sub_type_id(c, row):
    asset_type = row["ASSET_TYPE_ID"].value
    allowed = c.masters.sub_types_for(asset_type)
    typed = c.ov("ASSET_SUB_TYPE_ID")
    if typed is not None:
        code = as_text(typed)
        if code not in allowed:
            return invalid("ASSET_SUB_TYPE_ID", code, f"not a sub-type of asset type {asset_type}")
        return manual("ASSET_SUB_TYPE_ID", code)
    column, problem = c.column("Structure Type")
    if column:
        name = STRUCTURE_TYPES.get(re.sub(r"[^a-z]", "", column.lower()), "")
        understood = name or column in c.masters.asset_sub_types
        if understood and not asset_type:
            return blank("ASSET_SUB_TYPE_ID", f"worked out from COLLATERAL Structure Type '{column}' once the "
                                              "asset type (ASSET_TYPE_ID) is typed")
        code = c.masters.sub_type_id(asset_type, name) if name else (column if column in allowed else "")
        if code:
            return ok("ASSET_SUB_TYPE_ID", code, SRC_LMS, f"COLLATERAL Structure Type '{column}'")
        problem = (f"COLLATERAL Structure Type '{column}' is not understood for asset type {asset_type} (write "
                   "Structure with Land, Structure without Land or Land without Structure)" if understood else
                   f"COLLATERAL Structure Type '{column}' is not understood (write Structure with Land, "
                   "Structure without Land or Land without Structure)")
        c.issue("warning", "COLLATERAL columns", problem)
    why = f". {problem}" if problem else ""
    if not asset_type:
        return missing("ASSET_SUB_TYPE_ID", f"type the asset type (ASSET_TYPE_ID) and the sub-type that goes "
                                            f"with it{why}")
    return missing("ASSET_SUB_TYPE_ID", "not in COLLATERAL (Structure Type): pick Structure with Land, Structure "
                                        f"without Land or Land without Structure for asset type {asset_type}{why}")


# Column 6 - ASSET_UNIQUE_ID
# Optional. COLLATERAL "Global collateral No" (e.g. CMS00260877) when the
# setting USE_GLOBAL_COLLATERAL_NO_AS_ASSET_UNIQUE_ID is on; otherwise blank.
# Never "NA" (it would be stored as the asset's ID).
def col_06_asset_unique_id(c, row):
    typed = c.ov("ASSET_UNIQUE_ID")
    if typed is not None:
        return manual("ASSET_UNIQUE_ID", re.sub(r"[^A-Za-z0-9]", "", as_text(typed)))
    value = re.sub(r"[^A-Za-z0-9]", "", c.bundle["global_collateral_no"])
    if settings.USE_GLOBAL_COLLATERAL_NO_AS_ASSET_UNIQUE_ID and value:
        return ok("ASSET_UNIQUE_ID", value, SRC_LMS, "COLLATERAL Global collateral No")
    return blank("ASSET_UNIQUE_ID", "optional")


# Column 7 - ASSET_DESCRIPTION (mandatory, max 100 characters)
# Built from: property type + kind of unit + unit number + project name,
#   e.g. "COMMERCIAL OFFICE 21 YASHODHAM BUSINESS CSL". Each piece comes from the
#   LMS first; any piece the LMS lacks comes from title report item 1. With no
#   unit number, the plot number is used ("PLOT NO 45").
# Rules: cleaned to CERSAI's characters and cut at a word boundary.
# If it can't be found: at least "<property type> PROPERTY" is written, so it
#   is never empty.
def col_07_asset_description(c, row):
    typed = c.ov("ASSET_DESCRIPTION")
    if typed is not None:
        return manual("ASSET_DESCRIPTION", clean(typed, DESCRIPTION, 100))
    lms, pdf = c.lms_address, c.pdf
    used_pdf = False
    kind = lms.get("unit_kind") or ""
    numbers = lms.get("unit_numbers") or []
    if not numbers and pdf.get("unit_numbers"):
        kind, numbers, used_pdf = pdf.get("unit_kind") or kind, pdf["unit_numbers"], True
    project = lms.get("project") or ""
    if not project and pdf.get("project"):
        project, used_pdf = pdf["project"], True
    kind = {"property bearing": "property", "h": "house"}.get(kind, kind)
    number_text = "/".join(numbers)
    if not number_text:
        plots = lms.get("plot_numbers") or pdf.get("plot_numbers") or []
        if plots:
            kind, number_text = "", "PLOT NO " + "/".join(plots)
    property_type = c.bundle["property_type"] or ""
    pieces = [property_type if property_type != "PLOT" or not number_text else "", kind.upper(), number_text,
              project.upper()]
    text = " ".join(p for p in pieces if p)
    if not kind and not number_text and not project:
        text = f"{property_type} PROPERTY"
    source = SRC_PDF if used_pdf else SRC_LMS
    return ok("ASSET_DESCRIPTION", clean(text, DESCRIPTION, 100), source, "built from the property details")


# Column 8 - SURVEY_NUMBER
# Where it comes from: COLLATERAL "Survey No.", a column filled by hand, when
#   the extract has it and the cell is filled (several numbers separated by
#   commas, ";", "&" or "and"; a label such as "Survey No." in front is dropped
#   and "Hissa No." becomes "/"). Otherwise survey-type numbers in the LMS address (CTS No., Survey
#   No., S. No., Sy. No., C.S. No., Gat No., Khasra No., Khata/Khewat No.).
#   If the address has none, the same numbers from title report item 1.
# Rules: several numbers are joined with "^". When old and new survey numbers
#   are both given, only the new one is used, with its Hissa ("12/2").
#   When the LMS and the title report both give numbers and they differ, the
#   LMS numbers (column or address) are used and the difference is flagged.
# If it can't be found: blank. Either this column or PLOT_NUMBER (column 9)
#   must be filled; if both are empty, this column is reported as missing.
def col_08_survey_number(c, row):
    typed = c.ov("SURVEY_NUMBER")
    if typed is not None:
        return manual("SURVEY_NUMBER", clean(typed, SURVEY, 1000))
    pdf_numbers = c.pdf.get("survey_numbers") or []
    sep = settings.SURVEY_SEPARATOR
    column, _ = c.column("Survey No.")
    numbers = survey_list(column) if column else []
    if numbers:
        value = clean(_joined(numbers, sep), SURVEY, 1000)
        if pdf_numbers and not _same_numbers(numbers, pdf_numbers):
            return conflict("SURVEY_NUMBER", value, SRC_LMS, value, _joined(pdf_numbers, sep),
                            "COLLATERAL Survey No. and title report item 1 differ")
        return ok("SURVEY_NUMBER", value, SRC_LMS, "COLLATERAL Survey No.")
    if column:
        c.issue("warning", "COLLATERAL columns", f"COLLATERAL Survey No. '{column}' has no survey number in it")
    lms_numbers = c.lms_address.get("survey_numbers") or []
    if lms_numbers:
        value = clean(_joined(lms_numbers, sep), SURVEY, 1000)
        if pdf_numbers and not _same_numbers(lms_numbers, pdf_numbers):
            return conflict("SURVEY_NUMBER", value, SRC_LMS, value, _joined(pdf_numbers, sep),
                            "survey numbers in the LMS address and the title report differ")
        return ok("SURVEY_NUMBER", value, SRC_LMS, "from the LMS address")
    if pdf_numbers:
        return ok("SURVEY_NUMBER", clean(_joined(pdf_numbers, sep), SURVEY, 1000), SRC_PDF, "title report item 1")
    return blank("SURVEY_NUMBER", "not in the LMS address or the title report")


def survey_list(text):
    """COLLATERAL "Survey No." -> the survey numbers: "Survey No. 149/1/3, 150/6
    and 12 Hissa No. 2" -> ["149/1/3", "150/6", "12/2"]."""
    text = re.sub(r"\s*,?\s*\bhissa\s*no\.?\s*", "/", as_text(text), flags=re.I)
    out = []
    for part in re.split(r"\s*(?:[,;&^]|\band\b)\s*", text, flags=re.I):
        part = re.sub(rf"^(?:{SURVEY_KEYS})\s*[:.\-]?\s*", "", part.strip(), flags=re.I)
        part = clean(part, SURVEY, 1000).strip("^")
        if part and part not in out:
            out.append(part)
    return out


# Column 9 - PLOT_NUMBER (max 20 characters)
# Where it comes from: "Plot No." in the LMS address; if absent, "Plot No." or
#   the town-planning "FP/PF No." in title report item 1. Several plots are
#   joined with "/". A difference between the two sources is flagged.
# If it can't be found: blank - but if SURVEY_NUMBER is also empty, one of the
#   two is required, so SURVEY_NUMBER is reported as missing.
def col_09_plot_number(c, row):
    typed = c.ov("PLOT_NUMBER")
    if typed is not None:
        return manual("PLOT_NUMBER", clean(typed, ADDRESS, 20))
    lms_numbers = c.lms_address.get("plot_numbers") or []
    pdf_numbers = c.pdf.get("plot_numbers") or []
    if lms_numbers:
        value = clean("/".join(lms_numbers), ADDRESS, 20)
        if pdf_numbers and not _same_numbers(lms_numbers, pdf_numbers):
            return conflict("PLOT_NUMBER", value, SRC_LMS, value, "/".join(pdf_numbers),
                            "plot numbers in the LMS address and the title report differ")
        return ok("PLOT_NUMBER", value, SRC_LMS, "from the LMS address")
    if pdf_numbers:
        return ok("PLOT_NUMBER", clean("/".join(pdf_numbers), ADDRESS, 20), SRC_PDF, "title report item 1")
    if not row["SURVEY_NUMBER"].value:
        row["SURVEY_NUMBER"] = missing("SURVEY_NUMBER", "a survey number or a plot number is required")
    return blank("PLOT_NUMBER", "not in the LMS address or the title report")


# Column 10 - ASSET_AREA (optional; up to 99999.99)
# Where it comes from: COLLATERAL "Carpert Area" (the LMS does not state the
#   unit; it is taken to be square feet - to be confirmed with the LMS team).
#   If blank, the area in title report item 1 ("adm. 60.90 Sq. Mtrs.").
# Rules: when both exist, the title-report area is converted to square feet
#   and compared with the LMS carpet and built-up areas; if it is more than 10%
#   away from both, the LMS value is used and the difference is flagged.
# If it can't be found: blank (the area is optional).
def col_10_asset_area(c, row):
    typed = c.ov("ASSET_AREA")
    if typed is not None:
        value = _area(typed)
        return manual("ASSET_AREA", value) if value else invalid("ASSET_AREA", typed, "not an area up to 99999.99")
    lms_value = _area(c.bundle["carpet_area"])
    pdf_value, pdf_unit = _area(c.pdf.get("area", "")), c.pdf.get("area_unit", "")
    if lms_value:
        if pdf_value and pdf_unit in TO_SQFT and not _areas_close(c, float(pdf_value) * TO_SQFT[pdf_unit]):
            unit_name = c.masters.area_units.get(pdf_unit, pdf_unit)
            return conflict("ASSET_AREA", lms_value, SRC_LMS, f"{lms_value} sq ft (carpet)",
                            f"{pdf_value} {unit_name} ({c.pdf.get('area_type') or 'area'})",
                            "LMS and title-report areas differ by more than 10%")
        return ok("ASSET_AREA", lms_value, SRC_LMS, "COLLATERAL Carpert Area (sq ft)")
    if pdf_value and pdf_unit:
        return ok("ASSET_AREA", pdf_value, SRC_PDF, f"title report item 1 ({c.pdf.get('area_type') or 'area'})")
    if as_text(c.bundle["carpet_area"]):
        c.issue("warning", "Area", f"COLLATERAL Carpert Area '{c.bundle['carpet_area']}' is not a valid area")
    return blank("ASSET_AREA", "optional")


def _area(value) -> str:
    text = as_text(value).replace(",", "")
    if not re.fullmatch(r"\d+(\.\d+)?", text):
        return ""
    number = round(float(text), 2)
    if number <= 0 or number > 99999.99:
        return ""
    return f"{number:.2f}".rstrip("0").rstrip(".")


def _areas_close(c, pdf_sqft: float) -> bool:
    for raw in (c.bundle["carpet_area"], c.bundle["built_up_area"]):
        value = _area(raw)
        if value and abs(float(value) - pdf_sqft) <= settings.AREA_TOLERANCE * max(float(value), pdf_sqft):
            return True
    return False


# Column 11 - ASSET_AREA_UNIT
# Follows the area (column 10): an LMS area is in square feet -> 6; a
# title-report area uses the unit written next to it (Sq. Mtrs. -> 7,
# Sq. Ft. -> 6, Sq. Yds. -> 8, Acre -> 1 ...). Blank when there is no area.
# If an area was typed by hand without a unit, the unit is reported as missing.
def col_11_asset_area_unit(c, row):
    area = row["ASSET_AREA"]
    typed = c.ov("ASSET_AREA_UNIT")
    if not area.value:
        return not_applicable("ASSET_AREA_UNIT", "no area")
    if typed is not None:
        code = as_text(typed)
        if code not in c.masters.area_units:
            return invalid("ASSET_AREA_UNIT", code, "not an AREA_UNIT code")
        return manual("ASSET_AREA_UNIT", code)
    if area.source == SRC_LMS:
        return ok("ASSET_AREA_UNIT", settings.LMS_AREA_UNIT_ID, SRC_LMS, "LMS areas are in square feet")
    if area.source == SRC_PDF and c.pdf.get("area_unit"):
        return ok("ASSET_AREA_UNIT", c.pdf["area_unit"], SRC_PDF, "unit written in title report item 1")
    return missing("ASSET_AREA_UNIT", "give the unit of the area typed by hand")


# Column 12 - HOUSE_NUMBER (max 20 characters: letters, digits, "-", "/", space)
# Where it comes from: the unit number in the LMS address ("Flat No. 1102" ->
#   1102); if absent, the unit number in title report item 1. Several units
#   ("Flat No. 1301 & 1302") are joined with "/" -> 1301/1302; if that is longer
#   than 20 characters, the first unit is used. A difference between the two
#   sources is flagged.
# If it can't be found: required when the sub-type is a structure (with or
#   without land) -> reported as missing. Also reported as missing when the
#   sub-type itself is not known yet (unless the property is a PLOT), so both
#   can be filled in one go; otherwise blank.
def col_12_house_number(c, row):
    typed = c.ov("HOUSE_NUMBER")
    if typed is not None:
        return manual("HOUSE_NUMBER", clean(typed, HOUSE, 20))
    lms_numbers = c.lms_address.get("unit_numbers") or []
    pdf_numbers = c.pdf.get("unit_numbers") or []
    if lms_numbers:
        value = _house(lms_numbers)
        if pdf_numbers and not _same_numbers(lms_numbers, pdf_numbers):
            return conflict("HOUSE_NUMBER", value, SRC_LMS, value, _house(pdf_numbers),
                            "unit numbers in the LMS address and the title report differ")
        return ok("HOUSE_NUMBER", value, SRC_LMS, "from the LMS address")
    if pdf_numbers:
        return ok("HOUSE_NUMBER", _house(pdf_numbers), SRC_PDF, "title report item 1")
    sub_type = row["ASSET_SUB_TYPE_ID"].value
    if sub_type in STRUCTURE_SUB_TYPES:
        return missing("HOUSE_NUMBER", "required for a structure; not in the LMS address or title report")
    if not sub_type and c.bundle["property_type"] != "PLOT":
        return missing("HOUSE_NUMBER", "required unless the property is land without a structure; "
                                       "not in the LMS address or title report")
    return blank("HOUSE_NUMBER", "not needed for land")


def _house(numbers) -> str:
    joined = clean("/".join(numbers), HOUSE, 40)
    return joined if len(joined) <= 20 else clean(numbers[0], HOUSE, 20)


# Column 13 - FLOOR_NUMBER (optional; a whole number from -3 to 100)
# Where it comes from: the LMS address ("11th Floor" -> 11, "Ground Floor" ->
#   0, "Basement" -> -1); if absent, title report item 1. Addresses naming
#   several floors give no value. A difference between the sources is flagged.
# If it can't be found: blank.
def col_13_floor_number(c, row):
    typed = c.ov("FLOOR_NUMBER")
    if typed is not None:
        text = as_text(typed)
        if not re.fullmatch(r"-?\d{1,3}", text) or not -3 <= int(text) <= 100:
            return invalid("FLOOR_NUMBER", text, "must be a whole number from -3 to 100")
        return manual("FLOOR_NUMBER", str(int(text)))
    lms_floor, pdf_floor = c.lms_address.get("floor"), c.pdf.get("floor")
    if lms_floor is not None:
        if pdf_floor is not None and pdf_floor != lms_floor:
            return conflict("FLOOR_NUMBER", lms_floor, SRC_LMS, lms_floor, pdf_floor,
                            "floor in the LMS address and the title report differ")
        return ok("FLOOR_NUMBER", lms_floor, SRC_LMS, "from the LMS address")
    if pdf_floor is not None:
        return ok("FLOOR_NUMBER", pdf_floor, SRC_PDF, "title report item 1")
    return blank("FLOOR_NUMBER", "optional")


# Column 14 - BUILDING_NAME (optional; max 50 characters)
# Where it comes from: the part of the LMS address naming a wing, tower or
#   building; if absent, the building named in title report item 1
#   ("Building No. 3 known as 'Dev Daya Park'").
# If it can't be found: blank (never "NA").
def col_14_building_name(c, row):
    typed = c.ov("BUILDING_NAME")
    if typed is not None:
        return manual("BUILDING_NAME", clean(typed, ORG_NAME, 50))
    if c.lms_address.get("building"):
        return ok("BUILDING_NAME", clean(c.lms_address["building"], ORG_NAME, 50), SRC_LMS, "from the LMS address")
    if c.pdf.get("building"):
        return ok("BUILDING_NAME", clean(c.pdf["building"], ORG_NAME, 50), SRC_PDF, "title report item 1")
    return blank("BUILDING_NAME", "optional")


# Column 15 - PROJECT_NAME (mandatory; max 50 characters)
# Where it comes from: COLLATERAL "Project Name", a column filled by hand, when
#   the extract has it and the cell is filled. Otherwise the part of the LMS
#   address naming the society,
#   project or complex (CHS, CHSL, Society, Apartments, Residency, Complex...).
#   If absent, title report item 1: "project known as ..." first, otherwise
#   "Society known as ...".
# If it can't be found: reported as missing. (For a village plot with no
#   project, the colony or village name can be typed.)
def col_15_project_name(c, row):
    typed = c.ov("PROJECT_NAME")
    if typed is not None:
        return manual("PROJECT_NAME", clean(typed, ORG_NAME, 50))
    column, _ = c.column("Project Name")
    if column:
        name = clean(column, ORG_NAME, 50)
        if name:
            return ok("PROJECT_NAME", name, SRC_LMS, "COLLATERAL Project Name")
        c.issue("warning", "COLLATERAL columns", f"COLLATERAL Project Name '{column}' has no letters or digits")
    if c.lms_address.get("project"):
        return ok("PROJECT_NAME", clean(c.lms_address["project"], ORG_NAME, 50), SRC_LMS, "from the LMS address")
    if c.pdf.get("project"):
        return ok("PROJECT_NAME", clean(c.pdf["project"], ORG_NAME, 50), SRC_PDF, "title report item 1")
    return missing("PROJECT_NAME", "no society/project name in the LMS address or the title report")


# Column 16 - STREET_NAME (optional; max 50 characters: letters, digits, "-", "/")
# Where it comes from: the part of the LMS address naming a road, marg, street
#   or lane (landmarks such as "Near ..." are skipped); if absent, a road
#   named in title report item 1. Cut at a word boundary, never by removing
#   spaces.
# If it can't be found: blank.
def col_16_street_name(c, row):
    typed = c.ov("STREET_NAME")
    if typed is not None:
        return manual("STREET_NAME", clean(typed, STREET, 50))
    if c.lms_address.get("street"):
        return ok("STREET_NAME", clean(c.lms_address["street"], STREET, 50), SRC_LMS, "from the LMS address")
    if c.pdf.get("street"):
        return ok("STREET_NAME", clean(c.pdf["street"], STREET, 50), SRC_PDF, "title report item 1")
    return blank("STREET_NAME", "optional")


# Column 17 - POCKET (optional)
# Not in any source; blank unless typed in the required-details file.
def col_17_pocket(c, row):
    typed = c.ov("POCKET")
    if typed is not None:
        return manual("POCKET", clean(typed, ADDRESS, 20))
    return blank("POCKET", "optional")


def build(c):
    from . import build_row
    from .imm_location import LOCATION_COLUMNS
    columns = [col_01_row_type, col_02_asset_category_id, col_03_asset_type_id, col_04_asset_type_others,
               col_05_asset_sub_type_id, col_06_asset_unique_id, col_07_asset_description,
               col_08_survey_number, col_09_plot_number, col_10_asset_area, col_11_asset_area_unit,
               col_12_house_number, col_13_floor_number, col_14_building_name, col_15_project_name,
               col_16_street_name, col_17_pocket] + LOCATION_COLUMNS
    return build_row(FIELDS, columns, c, cleared=c.cleared())
