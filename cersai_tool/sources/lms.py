"""Reading the LMS extracts (CKYC.xlsx and COLLATERAL.xlsx).

Each input folder must hold exactly one .xlsx file (its name does not matter).
The data sheet is "Export Worksheet". Only active loans (COLLATERAL
"Loan Status" = A) are used, and each active collateral becomes one "bundle":
a plain dictionary with everything the rows need about that collateral, its
borrowers and its owners. Bundles are saved with the batch, so later steps
never have to read the input files again.
"""
from collections import OrderedDict
from pathlib import Path

import pandas as pd

from .. import settings
from ..rules.owners import unique_names
from ..rules.text import as_text

CKYC_COLUMNS = ["First Disbursal Date", "Application Number", "Loan Account No", "Customer Type",
                "Applicant Type", "Customer Name", "CONSTITUTION", "Neo CIF ID", "Date Of Birth",
                "Incorporation Date", "First Name", "Middle Name", "Last Name", "Identification Type",
                "Identification Number", "Mobile No", "Email ID", "Full Address"]
COLLATERAL_COLUMNS = ["Application number", "Loan Account #", "Loan Status", "Disbursal Date",
                      "Customer Name", "Customer Type", "Neo CIF ID", "Collateral ID",
                      "Global collateral No", "Owner Name", "Property Type", "Full Address",
                      "Property Pincode", "Build up Area", "Carpert Area", "Sanctioned Loan Amount",
                      "PHONE_NUMBER", "EMAIL_ADDRESS", "Identification Number", "ADDRESS_TYPE",
                      "Customer Address"]


class InputError(Exception):
    """A problem with the input files that stops the run."""


def find_single_xlsx(folder: Path) -> Path:
    if not folder.exists():
        raise InputError(f"Folder not found: {folder}")
    files = sorted(p for p in folder.glob("*.xlsx") if not p.name.startswith("~$"))
    if len(files) != 1:
        names = ", ".join(p.name for p in files) or "none"
        raise InputError(f"{folder} must contain exactly one .xlsx file (found: {names})")
    return files[0]


def read_export(path: Path, required: list) -> pd.DataFrame:
    book = pd.ExcelFile(path)
    sheet = "Export Worksheet" if "Export Worksheet" in book.sheet_names else book.sheet_names[0]
    df = pd.read_excel(book, sheet_name=sheet, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise InputError(f"{path.name} is missing columns: {', '.join(missing)}")
    return df.map(as_text)


def _first(values):
    for v in values:
        if v:
            return v
    return ""


def _distinct(values):
    return list(OrderedDict.fromkeys(v for v in values if v))


def build_bundles(ckyc: pd.DataFrame, coll: pd.DataFrame):
    """Returns (bundles, notes). notes lists problems found while reading."""
    notes = []
    active = coll[coll["Loan Status"].str.upper() == "A"]
    dropped = coll[coll["Loan Status"].str.upper() != "A"]["Application number"].nunique()
    if dropped:
        notes.append(f"{dropped} application(s) have cancelled or closed loan accounts (status X/C); "
                     "those rows are ignored")

    ckyc_by_app = {app: rows for app, rows in ckyc.groupby("Application Number")}
    # which active applications share a property (same Global collateral No)
    apps_by_property = active.groupby("Global collateral No")["Application number"].apply(
        lambda s: sorted(set(s))).to_dict()

    bundles = []
    for coll_id, rows in active.groupby("Collateral ID", sort=False):
        if not coll_id:
            continue
        app = _first(rows["Application number"])
        first = rows.iloc[0]
        problems = []
        for column, label in (("Loan Account #", "loan account"), ("Sanctioned Loan Amount", "sanctioned amount"),
                              ("Disbursal Date", "disbursal date")):
            if len(_distinct(rows[column])) > 1:
                problems.append(f"active rows show more than one {label}: {', '.join(_distinct(rows[column]))}")

        app_ckyc = ckyc_by_app.get(app)
        borrowers = _borrowers(app_ckyc, rows) if app_ckyc is not None else []
        if app_ckyc is None:
            problems.append("no CKYC rows for this application")

        global_no = as_text(first["Global collateral No"])
        shared = [a for a in apps_by_property.get(global_no, []) if a != app] if global_no else []

        bundles.append({
            "collateral_id": coll_id,
            "application_number": app,
            "loan_account": _first(rows["Loan Account #"]),
            "global_collateral_no": global_no,
            "property_type": as_text(first["Property Type"]).upper(),
            "property_address": as_text(first["Full Address"]),
            "property_pincode": as_text(first["Property Pincode"]),
            "carpet_area": as_text(first["Carpert Area"]),
            "built_up_area": as_text(first["Build up Area"]),
            "sanctioned_amounts": _distinct(rows["Sanctioned Loan Amount"]),
            "collateral_disbursal_date": _first(rows["Disbursal Date"]),
            "first_disbursal_date": _first(app_ckyc["First Disbursal Date"]) if app_ckyc is not None else "",
            "lms_owners": unique_names(rows["Owner Name"]),
            "borrowers": borrowers,
            "shared_with_applications": shared,
            "problems": problems,
        })
    return bundles, notes


def _borrowers(app_ckyc: pd.DataFrame, coll_rows: pd.DataFrame) -> list:
    """One entry per customer (Neo CIF ID) of the application: primary first,
    then co-applicants. Guarantors are included only if the setting says so."""
    out = []
    for cif, rows in app_ckyc.groupby("Neo CIF ID", sort=False):
        customer_type = _first(rows["Customer Type"])
        if customer_type.lower() == "guarantor" and not settings.INCLUDE_GUARANTORS:
            continue
        ids = {}
        for id_type, id_number in zip(rows["Identification Type"], rows["Identification Number"]):
            key = id_type.strip().upper()
            if id_number and key not in ids:
                ids[key] = id_number
        mine = coll_rows[coll_rows["Neo CIF ID"] == cif]
        addresses_by_type = {}
        for address_type, address in zip(mine["ADDRESS_TYPE"], mine["Customer Address"]):
            if address and address_type and address_type not in addresses_by_type:
                addresses_by_type[address_type] = address
        out.append({
            "neo_cif": cif,
            "customer_type": customer_type,
            "applicant_type": _first(rows["Applicant Type"]),
            "name": _first(rows["Customer Name"]),
            "constitution": _first(rows["CONSTITUTION"]),
            "first_name": _first(rows["First Name"]),
            "middle_name": _first(rows["Middle Name"]),
            "last_name": _first(rows["Last Name"]),
            "date_of_birth": _first(rows["Date Of Birth"]),
            "incorporation_date": _first(rows["Incorporation Date"]),
            "pan": ids.get("PAN", ""),
            "ckyc_number": ids.get("KYC IDENTIFICATION NUMBER", ""),
            "cin": ids.get("CIN NO", ""),
            "llpin": _first(v for k, v in ids.items() if "LLP" in k),
            "mobile": _first(rows["Mobile No"]),
            "email": _first(rows["Email ID"]),
            "ckyc_addresses": _distinct(rows["Full Address"]),
            "addresses_by_type": addresses_by_type,
            "collateral_phone": _first(mine["PHONE_NUMBER"]),
            "collateral_email": _first(mine["EMAIL_ADDRESS"]),
            "collateral_pan": _first(mine["Identification Number"]),
        })
    order = {"primary": 0, "co-applicant": 1, "guarantor": 2}
    out.sort(key=lambda b: (order.get(b["customer_type"].lower(), 3), b["neo_cif"]))
    return out


def load_bundles(paths):
    """Reads both LMS files and returns (bundles, notes, input_info)."""
    ckyc_file = find_single_xlsx(paths.ckyc_folder)
    coll_file = find_single_xlsx(paths.collateral_folder)
    ckyc = read_export(ckyc_file, CKYC_COLUMNS)
    coll = read_export(coll_file, COLLATERAL_COLUMNS)
    bundles, notes = build_bundles(ckyc, coll)
    info = {"ckyc_file": str(ckyc_file), "ckyc_rows": len(ckyc),
            "collateral_file": str(coll_file), "collateral_rows": len(coll)}
    return bundles, notes, info
