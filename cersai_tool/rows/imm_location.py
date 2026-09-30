"""IMM row, columns 18-27: where the property is (locality, city, district,
state, PIN, country, co-ordinates). Columns 1-17 are in imm_row.py.

City, district and state are CERSAI codes from the masters, worked out from
the PIN (see rules/geo.py): the PIN gives the district, the district gives
the state, and the city must be one of that district's cities.
"""
import re

from .. import settings
from ..model import SRC_LMS, SRC_MANUAL, SRC_PDF, blank, conflict, fixed, invalid, manual, missing, ok
from ..rules.geo import place_words, typed_city, typed_district
from ..rules.identifiers import clean_pin
from ..rules.text import ADDRESS, as_text, clean


def _written_out(text: str) -> str:
    return re.sub(r"\(\s*([EWNS])\s*\)",
                  lambda m: " " + {"E": "East", "W": "West", "N": "North", "S": "South"}[m.group(1).upper()],
                  text or "", flags=re.I)


# Column 18 - LOCALITY (mandatory; max 50 characters)
# Where it comes from: the area name just before the city at the end of the
#   LMS address (e.g. "Goregaon (E)" -> "Goregaon East"). Pieces that only
#   repeat the city or state, landmarks, and survey/area descriptions are
#   skipped. If the address gives nothing usable, the named area or village in
#   title report item 1 ("Vartak Nagar", "Village Mahajanwadi").
# If it can't be found: reported as missing.
def col_18_locality(c, row):
    typed = c.ov("LOCALITY")
    if typed is not None:
        return manual("LOCALITY", clean(typed, ADDRESS, 50))
    if c.lms_address.get("locality"):
        return ok("LOCALITY", clean(_written_out(c.lms_address["locality"]), ADDRESS, 50), SRC_LMS,
                  "from the LMS address")
    if c.pdf.get("locality"):
        return ok("LOCALITY", clean(c.pdf["locality"], ADDRESS, 50), SRC_PDF, "title report item 1")
    return missing("LOCALITY", "no locality in the LMS address or the title report")


# Column 19 - CITY (a code from the CERSAI CITY master)
# Where it comes from: of the cities in the property's district (column 20),
#   the longest name found in the LMS address wins ("GOREGAON EAST" beats
#   "MUMBAI"). Fallbacks: the city named at the end of the LMS address, if the
#   district has a city of that name; then a city named like the district.
# If it can't be found: reported as missing. A city can be typed by hand as a
#   CITY code or as a city name (looked up among the district's cities).
def col_19_city(c, row):
    typed = c.ov("CITY")
    if typed is not None:
        district = c.geo.district
        if c.ov("DISTRICT") is not None:
            district = typed_district(c.masters, as_text(c.ov("DISTRICT")), row["PIN_CODE"].value)[0] or district
        code, problem = typed_city(c.masters, as_text(typed), district)
        return manual("CITY", code) if code else invalid("CITY", typed, problem)
    geo = c.geo
    if geo.city:
        return ok("CITY", geo.city, SRC_LMS, f"{c.masters.cities[geo.city][0]}: {geo.city_note}")
    return missing("CITY", geo.problem or "city not found")


# Column 20 - DISTRICT (a code from the CERSAI DISTRICT master)
# Where it comes from: the property PIN (column 22) looked up in the PIN_CODE
#   master. When a PIN belongs to several districts, the district named in the
#   LMS address or in title report item 1 ("Taluka and District Thane") is used.
#   A city typed by hand also fixes the district.
# Rules: if the title report names a different district, it is flagged.
# If it can't be found: reported as missing. A district can be typed by hand
#   as a DISTRICT code or a district name.
def col_20_district(c, row):
    typed = c.ov("DISTRICT")
    if typed is not None:
        code, problem = typed_district(c.masters, as_text(typed), row["PIN_CODE"].value)
        return manual("DISTRICT", code) if code else invalid("DISTRICT", typed, problem)
    if row["CITY"].source == SRC_MANUAL and row["CITY"].value:
        return ok("DISTRICT", c.masters.district_of_city(row["CITY"].value), SRC_MANUAL,
                  "district of the city typed by hand")
    geo = c.geo
    if not geo.district:
        return missing("DISTRICT", geo.problem or "district not found")
    name = c.masters.districts[geo.district][0]
    pdf_district = c.pdf.get("district_name", "")
    if pdf_district and place_words(pdf_district).strip() not in place_words(name) \
            and place_words(name).strip() not in place_words(pdf_district):
        return conflict("DISTRICT", geo.district, SRC_LMS, f"{geo.district} ({name})", pdf_district,
                        "the title report names a different district")
    return ok("DISTRICT", geo.district, SRC_LMS, f"{name}, from PIN {geo.pin}")


# Column 21 - STATE (a code from the CERSAI STATE master)
# Where it comes from: the DISTRICT master gives each district's state.
# If it can't be found: missing when the district is missing.
def col_21_state(c, row):
    typed = c.ov("STATE")
    if typed is not None:
        code = as_text(typed)
        return manual("STATE", code) if code in c.masters.states else invalid("STATE", code, "not a STATE code")
    district = row["DISTRICT"].value
    if not district:
        return missing("STATE", "the district is not known")
    state = c.masters.state_of_district(district)
    return ok("STATE", state, row["DISTRICT"].source or SRC_LMS, c.masters.states.get(state, ""))


# Column 22 - PIN_CODE (6 digits, must be in the CERSAI PIN_CODE master)
# Where it comes from: COLLATERAL "Property Pincode". Only if that is blank or
#   not in the PIN master: a PIN written in title report item 1 (searched only
#   inside item 1 - the report also contains ARGFL's and the law firm's PINs).
# Rules: if the title report states a different PIN, it is flagged.
# If it can't be found: reported as missing.
def col_22_pin_code(c, row):
    typed = c.ov("PIN_CODE")
    if typed is not None:
        pin = clean_pin(typed)
        if not pin or pin not in c.masters.pin_districts:
            return invalid("PIN_CODE", typed, "not in the CERSAI PIN master")
        return manual("PIN_CODE", pin)
    pin, source = c.property_pin()
    if not pin or pin not in c.masters.pin_districts:
        found = c.bundle["property_pincode"] or "blank"
        return missing("PIN_CODE", f"COLLATERAL Property Pincode ({found}) is not in the CERSAI PIN master")
    pdf_pin = clean_pin(c.pdf.get("pin", ""))
    if source == SRC_LMS and pdf_pin and pdf_pin != pin:
        return conflict("PIN_CODE", pin, SRC_LMS, pin, pdf_pin, "the title report states a different PIN")
    note = "COLLATERAL Property Pincode" if source == SRC_LMS else "PIN stated in title report item 1"
    return ok("PIN_CODE", pin, source, note)


# Column 23 - COUNTRY
# Always "IND" (the only India code in the CERSAI masters).
def col_23_country(c, row):
    return fixed("COUNTRY", settings.COUNTRY)


# Columns 24-27 - LATITUDE_LONGITUDE_1 to 4
# Not in any source: always blank. The four empty positions are still written
# in the .DAT file, so the IMM line ends with "IND||||".
def col_24_latitude_longitude_1(c, row):
    return blank("LATITUDE_LONGITUDE_1", "not collected")


def col_25_latitude_longitude_2(c, row):
    return blank("LATITUDE_LONGITUDE_2", "not collected")


def col_26_latitude_longitude_3(c, row):
    return blank("LATITUDE_LONGITUDE_3", "not collected")


def col_27_latitude_longitude_4(c, row):
    return blank("LATITUDE_LONGITUDE_4", "not collected")


# PIN is worked out before city and district, which depend on it, then the
# columns are put back in CERSAI's order by the row builder.
LOCATION_COLUMNS = [col_18_locality, col_22_pin_code, col_19_city, col_20_district, col_21_state,
                    col_23_country, col_24_latitude_longitude_1, col_25_latitude_longitude_2,
                    col_26_latitude_longitude_3, col_27_latitude_longitude_4]
