"""Everything the row files need about one collateral, in one place.

A CollateralContext wraps the collateral's LMS bundle, its title report (if
any) and the values typed in the required-details file. It also works out, once,
the decisions several columns share:
  - is the title report usable, and is it the right PDF for this collateral
    (and not the report of another collateral of the same application)?
  - who owns the property (COLLATERAL owners, checked against their shares
    and the title report; the title report's owners when COLLATERAL has none)?
  - which borrowers are owners, and how many owners are not borrowers?
  - where is the property (PIN -> district -> state -> city)?
"""
import re
from dataclasses import dataclass, field

from . import settings
from .model import BORROWER_KEY_FIELDS, Issue, SRC_LMS, SRC_MANUAL, SRC_PDF, is_clear
from .rules import geo as geo_rules
from .rules.address import parse_property_address
from .rules.identifiers import clean_pin
from .rules.owners import (BUSINESS_WORDS, assign_owners, best_match, lists_agree, name_score, normalise_name,
                           split_owner_list, unique_names)
from .rules.text import normalise_for_compare

PDF_OK = "PDF OK"
WRONG_PDF = "WRONG PDF"


@dataclass
class OwnersDecision:
    owners: list
    source: str
    status: str            # manual / verified / differs / completed / not_verified / from_title_report / missing
    note: str = ""
    lms_owners: list = field(default_factory=list)
    pdf_owners: list = field(default_factory=list)
    count_problem: str = ""   # the COLLATERAL owners' shares or Ownership Status do not add up
    split_note: str = ""      # an Owner Name cell taken as several people, or that may hold several


@dataclass
class OwnerLink:
    owner: str
    neo_cif: str           # "" when the owner is not taken to be a borrower
    borrower_name: str
    score: int
    kind: str = "same"     # same / spelling / unclear / relative / different (rules/owners.compare_names)
    near_cif: str = ""     # the borrower the owner was compared with
    near_name: str = ""
    answer: str = ""       # Y or N typed in the required-details file

    @property
    def uncertain(self) -> bool:
        """The tool decided on its own between similar names: listed for review."""
        return self.kind in ("spelling", "unclear", "relative") and not self.answer


class CollateralContext:
    def __init__(self, bundle, title_report, masters, overrides, serial, run_date, siblings=()):
        self.bundle = bundle
        # the other collaterals of the same application: saved with the bundle
        # (also those filed earlier); older batches pass the batch's own ones
        self.siblings = list(bundle.get("siblings", siblings))
        self.collateral_id = bundle["collateral_id"]
        self.application_number = bundle["application_number"]
        self.serial = serial
        self.masters = masters
        self.overrides = overrides
        self.run_date = run_date
        self.raw_report = title_report
        self.issues = []
        self._cache = {}
        for problem in bundle.get("problems", []):
            self.issue("warning", "LMS data", problem)

    # -- values typed in the required-details file -------------------------
    # (BLANK is not a value: it empties the field, see cleared() below)
    def ov(self, field_name):
        value = self.overrides.collateral.get((self.collateral_id, field_name))
        return None if value in (None, "") or is_clear(value) else value

    def bov(self, borrower, *field_names):
        for name in field_names:
            value = self.overrides.borrower.get((self.application_number, borrower["neo_cif"], name))
            if value not in (None, "") and not is_clear(value):
                return value
        return None

    def pov(self, borrower, field_name):
        value = self.overrides.per_owner.get((self.collateral_id, borrower["neo_cif"], field_name))
        return None if value in (None, "") or is_clear(value) else value

    def cleared(self) -> set:
        """SI/IMM fields where BLANK was typed."""
        return {name for (coll, name), value in self.overrides.collateral.items()
                if coll == self.collateral_id and is_clear(value)}

    def cleared_for(self, borrower) -> set:
        """BOR fields of this borrower where BLANK was typed."""
        out = set()
        for (app, cif, key), value in self.overrides.borrower.items():
            if app == self.application_number and cif == borrower["neo_cif"] and is_clear(value):
                out.update(BORROWER_KEY_FIELDS.get(key, (key,)))
        return out

    def issue(self, level, topic, message):
        self.issues.append(Issue(self.collateral_id, level, topic, message))

    def _cached(self, key, build):
        if key not in self._cache:
            self._cache[key] = build()
        return self._cache[key]

    # -- basic data --------------------------------------------------------
    @property
    def borrowers(self) -> list:
        return self.bundle["borrowers"]

    @property
    def lms_address(self) -> dict:
        return self._cached("address", lambda: parse_property_address(self.bundle["property_address"]))

    # -- the hand-filled COLLATERAL columns (sources/lms.py) -----------------
    def column(self, name):
        """(value, problem) of a hand-filled COLLATERAL column for this
        collateral, e.g. column("Survey No."). ("", "") when the extract has no
        such column or its cells are blank; when the collateral's rows give
        different values, none is used and the problem is reported."""
        return self._cached(f"column:{name}", lambda: self._single(
            self.bundle.get("lms_columns", {}).get(name, []), f"COLLATERAL {name}", "rows of this collateral"))

    def father_name(self, borrower):
        """(value, problem) of COLLATERAL "Father Name" for this borrower, from
        any row of the application."""
        return self._cached(f"father:{borrower['neo_cif']}", lambda: self._single(
            borrower.get("father_names", []), "COLLATERAL Father Name", f"rows of {borrower['name']}"))

    def _single(self, values, label, where):
        if len(values) <= 1:
            return (values[0] if values else ""), ""
        problem = f"{label} differs between the {where} ({'; '.join(values)}), so it is not used"
        self.issue("warning", "COLLATERAL columns", problem)
        return "", problem

    @property
    def candidate_report(self):
        """The title report to check: None when there is none, it has no
        readable text, or someone marked it as the wrong PDF."""
        def build():
            rep = self.raw_report
            if not rep:
                return None
            if rep.get("status") != "ok":
                self.issue("warning", "Title report", f"{rep.get('file')} has no readable text (scanned?); not used")
                return None
            if (self.ov("PDF_CHECK") or "").upper() == WRONG_PDF:
                self.issue("info", "Title report", f"{rep.get('file')} marked as the wrong PDF; not used")
                return None
            return rep
        return self._cached("candidate_report", build)

    @property
    def report(self):
        """The title report whose values may be used: the candidate, unless the
        right-PDF check holds it (it may describe another property)."""
        rep = self.candidate_report
        return None if rep is None or self.right_pdf["held"] else rep

    @property
    def pdf(self) -> dict:
        """The property details read from item 1 of the usable title report."""
        return self.report["details"] if self.report else {}

    # -- is this the right PDF? -------------------------------------------
    @property
    def right_pdf(self) -> dict:
        return self._cached("right_pdf", self._check_right_pdf)

    def _check_right_pdf(self):
        rep = self.candidate_report
        result = {"held": False, "failures": 0, "checks": []}
        if not rep:
            return result
        names = [b["name"] for b in self.borrowers]
        applicant = re.sub(r"\(.*?\)", "", rep.get("loan_applicant", "")).strip()
        if applicant and names:
            _, score = best_match(applicant, names)
            result["checks"].append(("Loan applicant", score >= settings.PROBABLE_PERSON_SCORE,
                                     f"title report '{applicant}' vs borrowers {', '.join(names)}"))
        pdf_owners = rep.get("owners_item4") or rep.get("owners_header") or []
        lms_owners = self.collateral_owners[0]
        if pdf_owners and lms_owners:
            shared = any(name_score(a, b) >= settings.SAME_PERSON_SCORE for a in pdf_owners for b in lms_owners)
            result["checks"].append(("Owners", shared, f"title report {pdf_owners} vs LMS {lms_owners}"))
        # Flat / unit number: when the report's unit is not in this collateral's
        # LMS address, it fails only if the LMS address names a unit of its own
        # (an address without one cannot be compared). A unit that belongs to
        # another collateral of the same application means the two reports
        # were probably swapped: that alone holds the SI, because the loan
        # applicant and often the owners are the same for both.
        pdf_units = rep["details"].get("unit_numbers") or []
        lms_text = f" {normalise_for_compare(self.bundle['property_address'])} "
        sibling = ""
        if pdf_units and lms_text.strip():
            found = any(f" {normalise_for_compare(u)} " in lms_text for u in pdf_units)
            own_units = self.lms_address.get("unit_numbers") or []
            if not found:
                sibling = self._sibling_with_units(pdf_units)
            if found:
                result["checks"].append(("Flat/unit number", True, ""))
            elif own_units or sibling:
                where = f"; it is the unit of {sibling} of the same application" if sibling else ""
                result["checks"].append(("Flat/unit number", False, f"title report {', '.join(pdf_units)} vs LMS "
                                         f"{', '.join(own_units) or 'address'}{where}"))
        result["sibling"] = sibling
        result["failures"] = sum(1 for _, passed, _ in result["checks"] if not passed)
        released = (self.ov("PDF_CHECK") or "").upper() == PDF_OK
        if (result["failures"] >= settings.RIGHT_PDF_HOLD_FAILURES or sibling) and not released:
            result["held"] = True
            failed = "; ".join(f"{name}: {detail}" for name, passed, detail in result["checks"] if not passed)
            what = (f"seems to be the report of {sibling}, another collateral of the same application"
                    if sibling else "may belong to another property")
            self.issue("hold", "Right PDF check",
                       f"{rep['file']} {what} ({failed}). The SI is held until someone types '{PDF_OK}' or "
                       f"'{WRONG_PDF}' in the Check and correct sheet (or the files are renamed and step 1 is "
                       "run again).")
        elif result["failures"]:
            failed = "; ".join(f"{name}: {detail}" for name, passed, detail in result["checks"] if not passed)
            self.issue("warning", "Right PDF check", f"one check failed for {rep['file']}: {failed}")
        return result

    def _sibling_with_units(self, units) -> str:
        """The other collateral of this application whose LMS address names
        one of these unit numbers, or ''."""
        keys = {normalise_for_compare(u).replace(" ", "") for u in units}
        for other in self.siblings:
            other_units = parse_property_address(other["property_address"]).get("unit_numbers") or []
            if keys & {normalise_for_compare(u).replace(" ", "") for u in other_units}:
                return other["collateral_id"]
        return ""

    # -- owners ------------------------------------------------------------
    @property
    def collateral_owners(self):
        """(owners, notes): the COLLATERAL owners, with an Owner Name cell that
        holds several people split into them."""
        return self._cached("collateral_owners", lambda: self._split_joined_owners(self.bundle["lms_owners"]))

    def _split_joined_owners(self, cells):
        """An Owner Name cell sometimes holds two people ("KALAISELVI AND
        RADHAKRISHNAN SANKARANARAYANAN"). It is taken as several owners when each
        part is a different borrower and the whole name is not itself a
        borrower's name (company names contain "and" too: "JAMSON MACHINE AND
        TOOLS"). Other cells that may hold several people are pointed out."""
        def without_and(name):
            return " ".join(word for word in normalise_name(name).split() if word != "AND")

        names = [b["name"] for b in self.borrowers]
        borrower_names = {without_and(name) for name in names}
        owners, notes = [], []
        for cell in cells:
            parts = split_owner_list(cell)
            whole_is_borrower = without_and(cell) in borrower_names
            if len(parts) > 1 and not whole_is_borrower:
                if all(index is not None for index, *_ in assign_owners(parts, names)):
                    owners += parts
                    notes.append(f"the Owner Name '{cell}' is taken as {len(parts)} owners ({'; '.join(parts)}): "
                                 "each is a borrower")
                    continue
                if not BUSINESS_WORDS.search(cell):
                    notes.append(f"the Owner Name '{cell}' may be {len(parts)} people written in one cell")
            owners.append(cell)
        return unique_names(owners), notes

    @property
    def owners(self) -> OwnersDecision:
        return self._cached("owners", self._decide_owners)

    def _decide_owners(self):
        """The owners of this collateral. COLLATERAL "Owner Name" comes first (its
        "Percentage Share" and "Ownership Status" double-check the number of
        owners). The title report is the main source only when COLLATERAL has
        no owners; otherwise it verifies them, and owners it names differently
        are shown in the Check and correct sheet. It fills a gap only when the
        COLLATERAL shares show an owner is missing and it names every COLLATERAL
        owner plus others: those others are added."""
        typed = self.ov("OWNERS")
        cells = list(self.bundle["lms_owners"])
        lms, split_notes = self.collateral_owners
        rep = self.report             # not a title report held as possibly another property's
        pdf = (rep.get("owners_item4") or rep.get("owners_header") or []) if rep else []
        if rep:
            places = [("item 4", rep.get("owners_item4")), ("header box", rep.get("owners_header")),
                      ("item 12", rep.get("owners_item12"))]
            lists = [(where, names) for where, names in places if names]
            if len(lists) > 1 and not all(lists_agree(lists[0][1], names) for _, names in lists[1:]):
                self.issue("warning", "Owners", "the title report names different owners in "
                           + ", ".join(f"{w} ({'; '.join(n)})" for w, n in lists))
        count_problem, owner_missing = self._check_owner_count(cells) if cells else ("", False)
        if count_problem:
            self.issue("warning", "Owners in COLLATERAL", count_problem)
        split_note = "; ".join(split_notes)
        if split_note:
            self.issue("warning", "Owners in COLLATERAL", split_note)
        extra = (cells, pdf, count_problem, split_note)
        if typed:
            owners = [o.strip() for o in re.split(r";", typed) if o.strip()]
            return OwnersDecision(owners, SRC_MANUAL, "manual", "typed in the required-details file", *extra)
        if lms:
            if not rep:
                why = ("the title report is held as possibly another property's" if self.candidate_report
                       else "no usable title report to check them against")
                return OwnersDecision(lms, SRC_LMS, "not_verified", f"COLLATERAL owners; {why}", *extra)
            if pdf and lists_agree(lms, pdf):
                return OwnersDecision(lms, SRC_LMS, "verified", "COLLATERAL owners, the same in the title report",
                                      *extra)
            added = self._owners_to_add(lms, pdf) if owner_missing else []
            if added:
                return OwnersDecision(lms + added, SRC_PDF, "completed", "COLLATERAL owners, plus "
                                      f"{'; '.join(added)} from the title report (the COLLATERAL shares show an "
                                      "owner is missing)", *extra)
            if pdf:
                return OwnersDecision(lms, SRC_LMS, "differs", "COLLATERAL owners are used; the title report "
                                      f"names: {'; '.join(pdf)}", *extra)
            return OwnersDecision(lms, SRC_LMS, "not_verified", "COLLATERAL owners; the title report names no "
                                  "owners", *extra)
        if pdf:
            return OwnersDecision(pdf, SRC_PDF, "from_title_report", "COLLATERAL has no owners: the title "
                                  "report's owners are used", *extra)
        return OwnersDecision([], "", "missing", "no owners in COLLATERAL or a usable title report", *extra)

    def _check_owner_count(self, owners: list):
        """Do the COLLATERAL owners add up? Their Percentage Shares should total
        100%, and Ownership Status should be Single for one owner and Joint for
        two or more. Returns (the problems, whether an owner seems missing)."""
        problems, missing = [], False
        shares = self.bundle.get("lms_owner_shares") or []
        numbers = [float(x) for x in shares if re.fullmatch(r"\d+(\.\d+)?", x or "")]
        if numbers and len(numbers) == len(owners):
            total = sum(numbers)
            if abs(total - 100) > 1:
                missing = total < 100
                why = "an owner may be missing" if missing else "an owner may be listed twice"
                problems.append(f"the Percentage Shares add up to {total:g}%, not 100%, so {why}")
        status = (self.bundle.get("ownership_status") or "").strip().lower()
        if status == "single" and len(owners) != 1:
            problems.append(f"Ownership Status is Single but {len(owners)} owners are listed")
        if status == "joint" and len(owners) < 2:
            missing = True
            problems.append("Ownership Status is Joint but only one owner is listed")
        return "; ".join(problems), missing

    @staticmethod
    def _owners_to_add(lms: list, pdf: list) -> list:
        """When COLLATERAL seems to miss an owner: the title report's owners that
        are not in COLLATERAL, provided the title report names every COLLATERAL
        owner too (otherwise the two lists disagree and nothing is added)."""
        links = assign_owners(lms, pdf)
        if any(index is None for index, *_ in links):
            return []
        taken = {index for index, *_ in links}
        return [name for i, name in enumerate(pdf) if i not in taken]

    @property
    def owner_links(self) -> list:
        return self._cached("owner_links", self._link_owners)

    def _link_owners(self):
        """Each owner of this collateral -> the borrower it is (or none), decided
        automatically: each borrower is linked to one owner at most, similar
        names are taken to be the same person, and names whose given names
        differ (usually relatives) are not (rules/owners.assign_owners). A Y or
        N typed in the required-details file overrides the decision. Worked out
        separately for every collateral."""
        links = []
        names = [b["name"] for b in self.borrowers]
        owners = self.owners.owners
        for owner, (index, score, kind, near) in zip(owners, assign_owners(owners, names)):
            borrower = self.borrowers[index] if index is not None else None
            near_borrower = self.borrowers[near] if near is not None else None
            answer = ""
            if near_borrower is not None and kind in ("spelling", "unclear", "relative"):
                answer = (self.pov(near_borrower, "OWNER_MATCH") or "").strip().upper()[:1]
                if answer == "Y":
                    borrower = near_borrower
                elif answer == "N":
                    borrower = None
                else:
                    answer = ""
            links.append(OwnerLink(owner, borrower["neo_cif"] if borrower else "",
                                   borrower["name"] if borrower else "", score, kind,
                                   near_borrower["neo_cif"] if near_borrower else "",
                                   near_borrower["name"] if near_borrower else "", answer))
        notes = []
        for link in links:
            if link.uncertain and link.neo_cif:
                notes.append(f"owner '{link.owner}' is taken to be borrower '{link.borrower_name}' (similar names)")
            elif link.uncertain:
                notes.append(f"owner '{link.owner}' is taken to be a different person from borrower "
                             f"'{link.near_name}' (the given names differ: probably a relative)")
        if notes:
            self.issue("warning", "Owner names", "; ".join(notes) + ". If this is wrong, type Y or N (OWNER_MATCH) "
                                                 "in the Check and correct sheet.")
        return links

    @property
    def hold_reasons(self) -> list:
        """Why this SI must stay out of the .DAT file until someone answers a
        question in the required-details file (empty when it may go)."""
        if self.right_pdf["held"]:
            return ["the title report may belong to another property (answer PDF_CHECK)"]
        return []

    def non_borrower_owners(self) -> list:
        return [link.owner for link in self.owner_links if not link.neo_cif]

    # -- where is the property? -------------------------------------------
    @property
    def geo(self):
        return self._cached("geo", self._resolve_geo)

    def _resolve_geo(self):
        pin, source = self.property_pin()
        hints = [self.pdf.get("district_name", "")]
        result = geo_rules.resolve(self.masters, pin, self.bundle["property_address"],
                                   self.lms_address.get("lms_city", ""), hints)
        result.pin_source = source
        return result

    def property_pin(self):
        """(PIN, where it came from): typed value, else the LMS 'Property
        Pincode' if it is in the PIN master, else a PIN stated in item 1 of the
        title report if it is in the master."""
        typed = self.ov("PIN_CODE")
        if typed:
            return clean_pin(typed), SRC_MANUAL
        lms_pin = clean_pin(self.bundle["property_pincode"])
        if lms_pin and lms_pin in self.masters.pin_districts:
            return lms_pin, SRC_LMS
        pdf_pin = clean_pin(self.pdf.get("pin", ""))
        if pdf_pin and pdf_pin in self.masters.pin_districts:
            return pdf_pin, SRC_PDF
        return lms_pin, SRC_LMS
