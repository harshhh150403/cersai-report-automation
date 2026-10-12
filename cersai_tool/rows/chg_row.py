"""CHG row: the charge holder (ARGFL). 11 columns (sheet "charge_holder").

This row is fixed: it is the same for every security interest and never comes
from any input file. To change it, change the values below.
"""
from .. import settings
from ..model import blank, fixed

FIELDS = ["ROW_TYPE", "OFFICE_CODE", "OFFICE_NAME", "ADDRESS_LINE_1", "ADDRESS_LINE_2",
          "ADDRESS_LINE_3", "CITY", "DISTRICT", "STATE", "PIN_CODE", "COUNTRY"]

FIXED_VALUES = {
    "ROW_TYPE": "CHG",
    "OFFICE_CODE": "",                                   # must be empty
    "OFFICE_NAME": "ANAND RATHI GLOBAL FINANCE LIMITED",
    "ADDRESS_LINE_1": "Express Zone B Wing 5th Floor Western Express Highway Opposite Oberoi Mall "
                      "Malad East Mumbai 400063",
    "ADDRESS_LINE_2": "",
    "ADDRESS_LINE_3": "",
    "CITY": "81133",            # MALAD, in district 10680 like PIN 400063
    "DISTRICT": "10680",        # MUMBAI SUBURBAN
    "STATE": "21",              # Maharashtra
    "PIN_CODE": "400063",
    "COUNTRY": settings.BOR_CHG_COUNTRY,   # "IN": a constant in the batch format description
}


def build():
    return [fixed(name, FIXED_VALUES[name]) if FIXED_VALUES[name] else blank(name) for name in FIELDS]
