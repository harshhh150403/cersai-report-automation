"""PIN -> district -> state -> city, using the CERSAI masters.

1. The PIN must be in the PIN_CODE master. If it maps to one district, that is
   the district. If it maps to several, the district whose name appears in the
   address (or in the title report, e.g. "Taluka and District Thane") wins.
2. The state comes straight from the DISTRICT master.
3. The city must belong to that district: of the district's cities, the
   longest name found in the address wins ("GOREGAON EAST" beats "MALAD").
   The general city at the end of the address ("Mumbai") is used only when no
   more specific city is named, and a "city" named like the state (the master
   has one called "UTTAR PRADESH") never counts. Fallbacks: the city word at
   the end of the LMS address, if the district has a city of that name; then
   a city named like the district.
City names repeat across districts, so a city is never picked without the
district filter.
"""
import re
from dataclasses import dataclass

from .text import normalise_for_compare

_SHORT = {"EAST": "E", "WEST": "W", "NORTH": "N", "SOUTH": "S"}


def place_words(text: str) -> str:
    """Upper case, punctuation removed, padded with spaces so whole-phrase
    matching is easy. Directions are written short on both sides, so
    "Goregaon (E)", "Goregaon East" and "GOREGAON E" all read "GOREGAON E"."""
    text = re.sub(r"\(\s*([EWNS])\s*\)", r" \1 ", text or "", flags=re.I)
    words = [_SHORT.get(w, w) for w in normalise_for_compare(text).split()]
    return " " + " ".join(words) + " "


@dataclass
class GeoResult:
    pin: str = ""
    district: str = ""
    state: str = ""
    city: str = ""
    problem: str = ""          # why something could not be resolved
    city_note: str = ""        # how the city was chosen


def resolve(masters, pin: str, address_text: str, lms_city: str = "", hints=()) -> GeoResult:
    result = GeoResult(pin=pin)
    if not pin:
        result.problem = "no PIN code"
        return result
    districts = masters.pin_districts.get(pin, [])
    if not districts:
        result.problem = f"PIN {pin} is not in the CERSAI PIN master"
        return result
    words = place_words(" ".join([address_text or "", lms_city or ""] + [h for h in hints if h]))
    if len(districts) == 1:
        result.district = districts[0]
    else:
        named = [d for d in districts
                 if place_words(masters.districts.get(d, ("", ""))[0]).strip()
                 and place_words(masters.districts[d][0]) in words]
        if len(named) == 1:
            result.district = named[0]
        else:
            names = ", ".join(f"{d} {masters.districts.get(d, ('?', ''))[0]}" for d in districts)
            result.problem = f"PIN {pin} belongs to several districts ({names}); none is named in the address"
            return result
    result.state = masters.state_of_district(result.district)

    candidates = masters.cities_by_district.get(result.district, [])
    address_words = place_words(address_text or "")
    state_words = place_words(masters.states.get(result.state, "")).strip()
    general = place_words(lms_city).strip()
    best, general_match = None, None
    for code, name in candidates:
        name_words = place_words(name)
        key = name_words.strip()
        if not key or key == state_words or name_words not in address_words:
            continue
        if key == general:
            general_match = general_match or (code, name_words)
        elif best is None or len(name_words) > len(best[1]):
            best = (code, name_words)
    best = best or general_match
    if best:
        result.city = best[0]
        result.city_note = f"city '{best[1].strip()}' found in the address"
        return result
    if lms_city:
        lms_words = place_words(lms_city)
        for code, name in candidates:
            if place_words(name) == lms_words:
                result.city = code
                result.city_note = f"city '{name}' from the LMS city name"
                return result
    district_words = place_words(masters.districts.get(result.district, ("", ""))[0])
    for code, name in candidates:
        if place_words(name) == district_words:
            result.city = code
            result.city_note = f"city '{name}' named like the district"
            return result
    result.problem = "no city of this district is named in the address"
    return result


def typed_city(masters, typed: str, district: str = ""):
    """A city typed in the required-details file: either a CITY code, or a city
    name, looked up among the cities of the district when the district is
    known. Returns (code, problem)."""
    text = (typed or "").strip()
    if text in masters.cities:
        return text, ""
    words = place_words(text)
    pool = masters.cities_by_district.get(district, []) if district else \
        [(code, name) for code, (name, _) in masters.cities.items()]
    matches = [code for code, name in pool if place_words(name) == words]
    if len(matches) == 1:
        return matches[0], ""
    where = f" in district {district} ({masters.districts.get(district, ('?', ''))[0]})" if district else ""
    if not matches:
        return "", f"no city named '{text}'{where}"
    return "", f"several cities are named '{text}'{where}; type the city code"


def typed_district(masters, typed: str, pin: str = ""):
    """A district typed in the required-details file: a DISTRICT code, or a
    district name (looked up among the PIN's districts first). Returns
    (code, problem)."""
    text = (typed or "").strip()
    if text in masters.districts:
        return text, ""
    words = place_words(text)
    for pool in (masters.pin_districts.get(pin, []), list(masters.districts)):
        matches = [code for code in pool if place_words(masters.districts.get(code, ("", ""))[0]) == words]
        if len(matches) == 1:
            return matches[0], ""
        if len(matches) > 1:
            return "", f"several districts are named '{text}'; type the district code"
    return "", f"no district named '{text}'"
