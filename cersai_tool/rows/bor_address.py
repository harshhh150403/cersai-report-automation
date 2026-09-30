"""BOR row, columns 20-27: the borrower's address and its location codes.
Columns 1-19 are in bor_row.py.

Which address: COLLATERAL "Customer Address" together with its "ADDRESS_TYPE"
(CKYC's address has no type, and some customers have more than one):
  - individuals:     ResidentialAddress, then PermanentAddress, then OfficeAddress;
  - non-individuals: OfficeAddress, then AlternateBusinessAddress, then ResidentialAddress.
If none is found there, the first CKYC "Full Address".
The ending ", <City>, <State>- <PIN>, India" is removed (city, state and PIN
have their own columns) and the rest is cleaned and split into up to three
lines of at most 100 characters, at comma or word boundaries.
"""
from .. import settings
from ..model import SRC_LMS, SRC_MANUAL, blank, fixed, invalid, manual, missing, ok
from ..rules import geo as geo_rules
from ..rules.geo import typed_city, typed_district
from ..rules.address import split_lines, split_tail
from ..rules.identifiers import clean_pin
from ..rules.text import ADDRESS, as_text, clean
from .bor_row import is_individual

PREFERENCE_INDIVIDUAL = ["ResidentialAddress", "PermanentAddress", "OfficeAddress", "AlternateBusinessAddress"]
PREFERENCE_OTHERS = ["OfficeAddress", "AlternateBusinessAddress", "ResidentialAddress", "PermanentAddress"]


def borrower_address(c, b, row) -> dict:
    """The chosen address, split into its parts (cached per borrower).
    An address typed in the required-details file replaces the address lines;
    if it does not end with its own PIN, the city, state and PIN of the found
    address are kept."""
    def build():
        order = PREFERENCE_INDIVIDUAL if is_individual(row, b) else PREFERENCE_OTHERS
        raw, where = "", ""
        for address_type in order:
            if b["addresses_by_type"].get(address_type):
                raw, where = b["addresses_by_type"][address_type], f"COLLATERAL {address_type}"
                break
        if not raw and b["ckyc_addresses"]:
            raw, where = b["ckyc_addresses"][0], "CKYC Full Address"
        body, city, state, pin = split_tail(raw)
        typed = c.bov(b, "ADDRESS")
        if typed is not None:
            typed_body, typed_city_name, typed_state, typed_pin = split_tail(as_text(typed))
            body, where = typed_body, "typed in the required-details file"
            if typed_pin:
                city, state, pin = typed_city_name, typed_state, typed_pin
            raw = " ".join(p for p in (typed_body, city, state, pin) if p)
        lines = [clean(line, ADDRESS, 100) for line in split_lines(body, 100, 99)]
        lines = [line for line in lines if line]
        if len(lines) > 3:
            c.issue("warning", "Borrower address",
                    f"{b['name']}: the address needs more than 3 lines of 100 characters; the rest was cut")
            lines = lines[:3]
        return {"raw": raw, "where": where, "body": body, "lines": lines, "city": city, "state": state,
                "pin": pin}
    return c._cached(("bor_address", b["neo_cif"]), build)


def borrower_geo(c, b, row):
    def build():
        address = borrower_address(c, b, row)
        typed_pin = c.bov(b, "PIN_CODE")
        pin = clean_pin(typed_pin) if typed_pin is not None else clean_pin(address["pin"])
        return geo_rules.resolve(c.masters, pin, address["raw"], address["city"])
    return c._cached(("bor_geo", b["neo_cif"]), build)


def _line(c, b, row, number):
    field = f"ADDRESS_LINE_{number}"
    typed = c.bov(b, field)
    if typed is not None:
        return manual(field, clean(typed, ADDRESS, 100))
    address = borrower_address(c, b, row)
    lines = address["lines"]
    if len(lines) >= number:
        source = SRC_MANUAL if address["where"].startswith("typed") else SRC_LMS
        return ok(field, lines[number - 1], source, address["where"])
    if number == 1:
        return missing(field, "no address in COLLATERAL or CKYC")
    return blank(field, "address fits on fewer lines")


# Column 20 - ADDRESS_LINE_1 (mandatory; max 100 characters)
# Where it comes from: the chosen borrower address (see the top of this file),
#   first line. Cut at comma or word boundaries, never mid-word.
#   An address typed in the required-details file replaces the found one; its
#   city, state and PIN are kept unless the typed address ends with its own PIN.
# If it can't be found: reported as missing.
def col_20_address_line_1(c, b, serial, row):
    return _line(c, b, row, 1)


# Column 21 - ADDRESS_LINE_2 (optional)
# The rest of the address when it does not fit on line 1; otherwise blank.
def col_21_address_line_2(c, b, serial, row):
    return _line(c, b, row, 2)


# Column 22 - ADDRESS_LINE_3 (optional)
# The rest of the address when it does not fit on lines 1-2; otherwise blank.
def col_22_address_line_3(c, b, serial, row):
    return _line(c, b, row, 3)


# Column 23 - CITY (a code from the CERSAI CITY master)
# Where it comes from: of the cities in the borrower's district (column 24),
#   the longest name found in the address; fallbacks as for the property
#   (the LMS city name, then a city named like the district).
# If it can't be found: reported as missing. A city can be typed by hand as a
#   CITY code or as a city name (looked up among the district's cities).
def col_23_city(c, b, serial, row):
    typed = c.bov(b, "CITY")
    if typed is not None:
        district = borrower_geo(c, b, row).district
        if c.bov(b, "DISTRICT") is not None:
            district = typed_district(c.masters, as_text(c.bov(b, "DISTRICT")), row["PIN_CODE"].value)[0] or district
        code, problem = typed_city(c.masters, as_text(typed), district)
        return manual("CITY", code) if code else invalid("CITY", typed, problem)
    geo = borrower_geo(c, b, row)
    if geo.city:
        return ok("CITY", geo.city, SRC_LMS, f"{c.masters.cities[geo.city][0]}: {geo.city_note}")
    return missing("CITY", geo.problem or "city not found")


# Column 24 - DISTRICT (a code from the CERSAI DISTRICT master)
# Where it comes from: the borrower's PIN (column 26) in the PIN_CODE master;
#   when a PIN belongs to several districts, the one named in the address.
#   A city typed by hand also fixes the district.
# If it can't be found: reported as missing. A district can be typed by hand
#   as a DISTRICT code or a district name.
def col_24_district(c, b, serial, row):
    typed = c.bov(b, "DISTRICT")
    if typed is not None:
        code, problem = typed_district(c.masters, as_text(typed), row["PIN_CODE"].value)
        return manual("DISTRICT", code) if code else invalid("DISTRICT", typed, problem)
    if row["CITY"].source == SRC_MANUAL and row["CITY"].value:
        return ok("DISTRICT", c.masters.district_of_city(row["CITY"].value), SRC_MANUAL,
                  "district of the city typed by hand")
    geo = borrower_geo(c, b, row)
    if geo.district:
        return ok("DISTRICT", geo.district, SRC_LMS, f"{c.masters.districts[geo.district][0]}, from PIN {geo.pin}")
    return missing("DISTRICT", geo.problem or "district not found")


# Column 25 - STATE (a code from the CERSAI STATE master)
# Where it comes from: the state of the district in column 24 (DISTRICT master).
def col_25_state(c, b, serial, row):
    district = row["DISTRICT"].value
    if not district:
        return missing("STATE", "the district is not known")
    state = c.masters.state_of_district(district)
    return ok("STATE", state, row["DISTRICT"].source or SRC_LMS, c.masters.states.get(state, ""))


# Column 26 - PIN_CODE (6 digits, must be in the CERSAI PIN_CODE master)
# Where it comes from: the 6 digits before ", India" at the end of the chosen
#   address.
# If it can't be found: reported as missing (also when the PIN is not in the
#   CERSAI master).
def col_26_pin_code(c, b, serial, row):
    typed = c.bov(b, "PIN_CODE")
    if typed is not None:
        pin = clean_pin(typed)
        if not pin or pin not in c.masters.pin_districts:
            return invalid("PIN_CODE", typed, "not in the CERSAI PIN master")
        return manual("PIN_CODE", pin)
    pin = clean_pin(borrower_address(c, b, row)["pin"])
    if not pin:
        return missing("PIN_CODE", "no PIN at the end of the address")
    if pin not in c.masters.pin_districts:
        return invalid("PIN_CODE", pin, "not in the CERSAI PIN master")
    return ok("PIN_CODE", pin, SRC_LMS, "end of the address")


# Column 27 - COUNTRY
# Always "IND" (the only India code in the CERSAI masters; the spec text says
# "IN" here - confirm with the first upload).
def col_27_country(c, b, serial, row):
    return fixed("COUNTRY", settings.COUNTRY)


# PIN is worked out before city and district, which depend on it; the row
# builder puts the columns back in CERSAI's order.
ADDRESS_COLUMNS = [col_20_address_line_1, col_21_address_line_2, col_22_address_line_3, col_26_pin_code,
                   col_23_city, col_24_district, col_25_state, col_27_country]
