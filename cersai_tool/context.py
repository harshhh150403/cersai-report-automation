"""Everything the row files need about one collateral, in one place.

A CollateralContext wraps the collateral's LMS bundle, its title report (if
any) and the values typed in the required-details file. It also works out, once,
the decisions several columns share:
  - is the title report usable, and is it the right PDF for this collateral?
  - who owns the property (LMS owners verified against the title report)?
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
from .rules.owners import best_match, lists_agree, name_score, split_owner_list
from .rules.text import normalise_for_compare

PDF_OK = "PDF OK"
WRONG_PDF = "WRONG PDF"


@dataclass
class OwnersDecision:
    owners: list
    source: str
    status: str            # verified / conflict / from_title_report / not_verified / manual / missing / held
    note: str = ""
    lms_owners: list = field(default_factory=list)
    pdf_owners: list = field(default_factory=list)


@dataclass
class OwnerLink:
    owner: str
    neo_cif: str           # "" when the owner is not a borrower
    borrower_name: str
    score: int
    probable: bool         # a close but not certain match: someone should confirm it


class CollateralContext:
    def __init__(self, bundle, title_report, masters, overrides, serial, run_date):
        self.bundle = bundle
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
        lms_owners = self.bundle["lms_owners"]
        if pdf_owners and lms_owners:
            shared = any(name_score(a, b) >= settings.SAME_PERSON_SCORE for a in pdf_owners for b in lms_owners)
            result["checks"].append(("Owners", shared, f"title report {pdf_owners} vs LMS {lms_owners}"))
        pdf_units = rep["details"].get("unit_numbers") or []
        lms_text = f" {normalise_for_compare(self.bundle['property_address'])} "
        if pdf_units and lms_text.strip():
            found = any(f" {normalise_for_compare(u)} " in lms_text for u in pdf_units)
            result["checks"].append(("Flat/unit number", found,
                                     f"title report {pdf_units} not in the LMS address" if not found else ""))
        result["failures"] = sum(1 for _, passed, _ in result["checks"] if not passed)
        released = (self.ov("PDF_CHECK") or "").upper() == PDF_OK
        if result["failures"] >= settings.RIGHT_PDF_HOLD_FAILURES and not released:
            result["held"] = True
            failed = "; ".join(f"{name}: {detail}" for name, passed, detail in result["checks"] if not passed)
            self.issue("hold", "Right PDF check",
                       f"{rep['file']} may belong to another property ({failed}). The SI is held until "
                       f"someone types '{PDF_OK}' or '{WRONG_PDF}' in the Check and correct sheet.")
        elif result["failures"]:
            failed = "; ".join(f"{name}: {detail}" for name, passed, detail in result["checks"] if not passed)
            self.issue("warning", "Right PDF check", f"one check failed for {rep['file']}: {failed}")
        return result

    # -- owners ------------------------------------------------------------
    @property
    def owners(self) -> OwnersDecision:
        return self._cached("owners", self._decide_owners)

    def _decide_owners(self):
        typed = self.ov("OWNERS")
        lms = self.bundle["lms_owners"]
        rep = self.candidate_report
        pdf = (rep.get("owners_item4") or rep.get("owners_header") or []) if rep else []
        if rep:
            places = [("item 4", rep.get("owners_item4")), ("header box", rep.get("owners_header")),
                      ("item 12", rep.get("owners_item12"))]
            lists = [(where, names) for where, names in places if names]
            if len(lists) > 1 and not all(lists_agree(lists[0][1], names) for _, names in lists[1:]):
                self.issue("warning", "Owners", "the title report names different owners in "
                           + ", ".join(f"{w} ({'; '.join(n)})" for w, n in lists))
        if typed:
            owners = [o.strip() for o in re.split(r";", typed) if o.strip()]
            return OwnersDecision(owners, SRC_MANUAL, "manual", "typed in the required-details file", lms, pdf)
        if lms and pdf:
            if lists_agree(lms, pdf):
                return OwnersDecision(lms, SRC_LMS, "verified", "LMS owners match the title report", lms, pdf)
            if self.right_pdf["held"]:
                return OwnersDecision(lms, SRC_LMS, "held",
                                      "LMS and title report owners differ, and the PDF may be the wrong one", lms, pdf)
            return OwnersDecision(pdf, SRC_PDF, "conflict",
                                  "LMS and title report owners differ; the title report's owners are used", lms, pdf)
        if pdf:
            return OwnersDecision(pdf, SRC_PDF, "from_title_report", "no owners in the LMS", lms, pdf)
        if lms:
            return OwnersDecision(lms, SRC_LMS, "not_verified", "no usable title report to verify the owners",
                                  lms, pdf)
        return OwnersDecision([], "", "missing", "no owners in the LMS or the title report", lms, pdf)

    @property
    def owner_links(self) -> list:
        return self._cached("owner_links", self._link_owners)

    def _link_owners(self):
        links = []
        names = [b["name"] for b in self.borrowers]
        for owner in self.owners.owners:
            index, score = best_match(owner, names)
            borrower = self.borrowers[index] if index is not None else None
            probable = False
            if borrower is None or score < settings.PROBABLE_PERSON_SCORE:
                borrower = None
            elif score < settings.SAME_PERSON_SCORE:
                answer = (self.pov(borrower, "OWNER_MATCH") or "").upper()
                if answer.startswith("N"):
                    borrower = None
                elif not answer.startswith("Y"):
                    probable = True
            links.append(OwnerLink(owner, borrower["neo_cif"] if borrower else "",
                                   borrower["name"] if borrower else "", score, probable))
        return links

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
