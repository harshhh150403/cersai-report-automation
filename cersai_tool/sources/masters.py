"""CERSAI master tables (Master data/CERSAI2.0_Masters.xls).

Only the sheets the SI registration batch needs are read. Every code is kept
as text, exactly as in the master.
"""
from collections import defaultdict
from pathlib import Path

import xlrd

from ..rules.text import as_text


class Masters:
    def __init__(self, path: Path):
        if not path.exists():
            raise FileNotFoundError(f"CERSAI masters not found: {path}")
        book = xlrd.open_workbook(str(path))
        self.version = as_text(book.sheet_by_index(0).cell_value(0, 0))

        def rows(sheet_name):
            sheet = book.sheet_by_name(sheet_name)
            for r in range(3, sheet.nrows):          # headers are on row 3
                values = [as_text(v) for v in sheet.row_values(r)]
                if any(values):
                    yield values

        # SI type (immovable rows only: IDs 17-21 plus attachment order 6)
        self.si_types = {}
        for code, description, category in (r[:3] for r in rows("SI_SUB_TYPE_ID")):
            self.si_types[code] = (description, category)
        self.immovable_si_types = {c for c, (_, cat) in self.si_types.items() if cat == "Immovable"}

        self.financing_types = {r[0]: r[1] for r in rows("FINANCING_TYPE_ID")}
        self.asset_types = {r[0]: (r[1], r[2]) for r in rows("ASSET_TYPE_ID")}          # id -> (category, name)
        self.asset_sub_types = {r[0]: (r[1], r[2]) for r in rows("ASSET_SUB_TYPE_ID")}  # id -> (type, name)
        self.area_units = {r[0]: r[1] for r in rows("AREA_UNIT")}
        self.genders = {r[0]: r[1] for r in rows("GENDER")}
        self.borrower_types = {r[0]: r[1] for r in rows("BORROWER_ASSET_OWNER_TYPE")}
        self.countries = {r[0]: r[1] for r in rows("COUNTRY")}
        self.states = {r[0]: r[1] for r in rows("STATE")}

        self.districts = {}                        # code -> (name, state code)
        for code, name, state_code, _state in (r[:4] for r in rows("DISTRICT")):
            self.districts[code] = (name, state_code)

        self.cities = {}                           # code -> (name, district code)
        self.cities_by_district = defaultdict(list)
        for code, name, district_code, _district in (r[:4] for r in rows("CITY")):
            self.cities[code] = (name, district_code)
            self.cities_by_district[district_code].append((code, name))

        self.pin_districts = defaultdict(list)     # pin -> [district codes]
        for values in rows("PIN_CODE"):
            pin, district_code = values[0], values[1]
            if district_code not in self.pin_districts[pin]:
                self.pin_districts[pin].append(district_code)

    # -- small lookups used by the rows and the checks --------------------
    def sub_types_for(self, asset_type_id: str) -> dict:
        return {k: v[1] for k, v in self.asset_sub_types.items() if v[0] == asset_type_id}

    def sub_type_id(self, asset_type_id: str, description: str) -> str:
        """'Structure without Land' under asset type '2' -> '5'."""
        wanted = description.strip().lower()
        for code, name in self.sub_types_for(asset_type_id).items():
            if name.strip().lower() == wanted:
                return code
        return ""

    def state_of_district(self, district_code: str) -> str:
        return self.districts.get(district_code, ("", ""))[1]

    def district_of_city(self, city_code: str) -> str:
        return self.cities.get(city_code, ("", ""))[1]
