"""Creates the dummy input data used while developing and testing the tool.

    python dev_tools/make_dummy_data.py

Writes (replacing earlier dummy files):
    Dummy input data/Dummy CKYC/CKYC.xlsx
    Dummy input data/Dummy COLLATERAL/COLLATERAL.xlsx
    Dummy input data/Dummy Legal/*.pdf
    Dummy input data/Dummy data cases.md    (what each application tests)
    Dummy input data/Dummy required details answers.xlsx
                                            (what a person would type in the
                                            required-details file; typed in by
                                            dev_tools/fill_dummy_required_details.py)

The hand-made applications each test a kind of data found in the real
extracts. Regular applications are then generated (with a fixed seed, so every
run gives the same files) until COLLATERAL has exactly COLLATERALS distinct
Collateral IDs; each one also gets its CKYC rows and a title report for every
collateral. Every loan is active (Loan Status A): the tool files only active
loans, so the dummy data has no cancelled (X) or closed (C) ones.

Every person, company, PAN, phone number, e-mail and law firm is made up.
PIN codes and place names are real, because the tool looks them up in the
CERSAI masters. The files copy the layout of the real LMS extracts: the same
sheets ("Export Worksheet" and "SQL"), columns, and value formats (DD-MM-YYYY
and DD-MON-YY dates, "+91--" mobiles, numbers stored as numbers). CKYC also
has the Gender column (Male / Female) that the CKYC table provides. Rows are
built the way the LMS queries build them: CKYC has one row per customer x ID
document x address; COLLATERAL has one row per loan account x customer x
owner x customer address.
"""
import difflib
import random
import re
from datetime import date, timedelta
from pathlib import Path

import openpyxl
import pymupdf
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Dummy input data"
CKYC_DIR, COLL_DIR, LEGAL_DIR = OUT / "Dummy CKYC", OUT / "Dummy COLLATERAL", OUT / "Dummy Legal"
# the values a person would type in the required-details file (see write_answers)
ANSWERS_FILE = OUT / "Dummy required details answers.xlsx"

CKYC_HEADERS = ["First Disbursal Date", "Branch Name", "DOD", "Application Number", "Loan Account No",
                "Customer Type", "Applicant Type", "Customer Name", "CONSTITUTION", "Neo CIF ID", "Date Of Birth",
                "Incorporation Date", "First Name", "Middle Name", "Last Name", "Identification Type",
                "Identification Number", "Mobile No", "Email ID", "Full Address", "Customer Risk Category",
                "Primary Mobile", "Product ID", "Application Received Date", "Gender"]
COLL_HEADERS = ["Application number", "Loan Account #", "Loan Status", "Disbursal Date", "Branch Name",
                "Customer Name", "Customer Type", "Neo CIF ID", "Property Purpose", "Collateral ID",
                "Global collateral No", "Property Ownership", "Ownership Status", "Owner Name", "Percentage Share",
                "Property Type", "Property Classification", "Nature of Property", "Full Address",
                "Property Pincode", "FOIR(%)", "Type of loan eligibility program", "Build up Area", "Carpert Area",
                "Age of Property (In Years)", "Market Value", "Sanctioned Loan Amount", "Disbursed Loan Amount",
                "PHONE_NUMBER", "EMAIL_ADDRESS", "Identification Number", "ADDRESS_TYPE", "Customer Address"]
MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]


# --------------------------------------------------------------------------
# Building blocks
# --------------------------------------------------------------------------
# Neo CIF IDs and Global collateral Nos are made 9 million higher than the
# numbers used below, so that they cannot be the IDs of real customers or
# properties (the real ones are below 300,000).
ID_OFFSET = 9_000_000


def cif(n):
    return f"GLBCUST{n + ID_OFFSET:011d}"


def lan(n):
    return f"ARGFLLAPMUM2{n:012d}"


GSTIN_RE = re.compile(r"\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z]")
GSTIN_CHARS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def gstin_check(first14):
    """The 15th character of a GSTIN (the GSTN check: each character's value
    times 1 and 2 in turn, folded in base 36)."""
    total = 0
    for i, ch in enumerate(first14):
        product = GSTIN_CHARS.index(ch) * (2 if i % 2 else 1)
        total += product // 36 + product % 36
    return GSTIN_CHARS[(36 - total % 36) % 36]


def with_gstin_check(ids):
    """ID documents with every GSTIN given its right check character (GST
    numbers also appear under other ID types, as in the real extract)."""
    return [(kind, value[:14] + gstin_check(value[:14]) if value and GSTIN_RE.fullmatch(value) else value)
            for kind, value in ids]


def person(n, name, ctype, dob, pan=None, *, ids=None, first=None, middle=None, last=None, mobile="auto",
           email=None, addresses=(), ckyc_addresses=None, phone=None, coll_email=None,
           applicant_type="individual", constitution="Individual", inc=None, risk="Low Risk", gender=None,
           parent=None):
    """An individual customer. addresses: [(ADDRESS_TYPE, text)] as in COLLATERAL;
    CKYC gets the same texts unless ckyc_addresses is given. gender and parent
    (the father's or mother's name) are not in the extracts: they go to the
    answers file (see ANSWERS_FILE)."""
    words = name.split()
    return {
        "number": n, "cif": cif(n), "name": name, "ctype": ctype, "applicant_type": applicant_type,
        "constitution": constitution, "dob": dob if applicant_type == "individual" else None,
        "inc": inc if applicant_type != "individual" else None,
        "first": first if first is not None else words[0],
        "middle": middle if middle is not None else (" ".join(words[1:-1]) or None),
        "last": last if last is not None else (words[-1] if len(words) > 1 else None),
        "ids": with_gstin_check(ids if ids is not None else ([("PAN", pan)] if pan else [])),
        "pan": pan, "mobile": f"+91--55{n % 100000000:08d}" if mobile == "auto" else mobile,
        "email": email, "addresses": list(addresses),
        "ckyc_addresses": ckyc_addresses if ckyc_addresses is not None else [a for _, a in addresses],
        "phone": phone, "coll_email": coll_email, "risk": risk, "gender": gender, "parent": parent,
    }


def firm(n, name, constitution, ctype, inc, pan=None, *, ids=None, mobile="auto", email=None, addresses=(),
         ckyc_addresses=None, applicant_type="non_individual", dob=None, risk="Low Risk", phone=None,
         coll_email=None):
    """A non-individual customer (firm, company, LLP, HUF, trust)."""
    return {
        "number": n, "cif": cif(n), "name": name, "ctype": ctype, "applicant_type": applicant_type,
        "constitution": constitution, "dob": dob if applicant_type == "individual" else None,
        "inc": inc if applicant_type != "individual" else None,
        "first": None, "middle": None, "last": None,
        "ids": with_gstin_check(ids if ids is not None else ([("PAN", pan)] if pan else [])),
        "pan": pan, "mobile": f"+91--56{n % 100000000:08d}" if mobile == "auto" else mobile,
        "email": email, "addresses": list(addresses),
        "ckyc_addresses": ckyc_addresses if ckyc_addresses is not None else [a for _, a in addresses],
        "phone": phone, "coll_email": coll_email, "risk": risk,
    }


def loan(number, status, date, sanctioned, disbursed=None):
    """A loan account. The dummy data has active loans only (Loan Status A),
    as the tool files only those."""
    if status != "A":
        raise ValueError(f"loan {number}: the dummy data has active loans only (status A), not {status}")
    return {"lan": lan(number), "status": status, "date": date, "sanctioned": sanctioned,
            "disbursed": sanctioned if disbursed is None else disbursed}


def collateral(coll, cms, ptype, address, pin, owners, *, built=None, carpet=None, classification="Others",
               nature="FREEHOLD", purpose="Self Occupied", ownership="Others", age=None, market=None, pdf=None,
               status=None):
    """owners: [(name as typed in the LMS, percentage share)]; pdf: a title-report spec or None;
    status: the Ownership Status (default Single for one owner, Joint for more)."""
    return {"coll": coll, "cms": cms, "ptype": ptype, "address": address, "pin": pin, "owners": owners,
            "built": built, "carpet": carpet, "classification": classification, "nature": nature,
            "purpose": purpose, "ownership": ownership, "age": (0 if ptype == "PLOT" else 8) if age is None else age,
            "market": market, "pdf": pdf,
            "status": status}


def application(app, branch, first_disbursal, customers, collaterals, loans, *, received="15-MAY-26",
                product=5000980, foir="48", program="1 Vanilla Income Program", case=""):
    a = {"app": app, "branch": branch, "first_disbursal": first_disbursal, "customers": customers,
         "collaterals": collaterals, "loans": loans, "received": received, "product": product, "foir": foir,
         "program": program, "case": case}
    for col in collaterals:
        if col["pdf"] and not col["pdf"]["initiated"]:
            col["pdf"]["initiated"], col["pdf"]["issued"] = report_dates(a)
    return a


def report_dates(a):
    """(initiated, issued) of a title report without dates of its own: 14 and
    9 days before the loan was disbursed."""
    day, month, year = a["loans"][0]["date"].split("-")
    disbursed = date(2000 + int(year), MONTHS.index(month) + 1, int(day))
    return tuple((disbursed - timedelta(days=d)).strftime("%d/%m/%Y") for d in (14, 9))


def report(template, applicant, owners, description, *, nature="Residential.", transaction="LAP",
           mortgage="Equitable Mortgage", owners_item4=None, owners_item12=None, status=None,
           file_name=None, initiated=None, issued=None):
    """A title report. template: 'A' (5-page style, no item 6, land history inside
    item 5), 'B' (continuation-sheet style with item 6 'Flow of Title' and the
    search section struck through) or 'scanned' (an image with no text)."""
    return {"template": template, "applicant": applicant, "owners": owners, "description": description,
            "nature": nature, "transaction": transaction, "mortgage": mortgage,
            "owners_item4": owners_item4 or owners, "owners_item12": owners_item12 or owners_item4 or owners,
            "status": status or ("POSITIVE" if template == "A" else "Positive subject to search"),
            "file_name": file_name, "initiated": initiated, "issued": issued}


# --------------------------------------------------------------------------
# The applications
# --------------------------------------------------------------------------
def build_applications():
    apps = []

    # D01 - the "golden" record: everything found, title report agrees.
    goregaon_office = "Office No. 14, Sunbeam Trade Centre, 1st Floor, Above Corner Bakery, Goregaon East, Mumbai, " \
                      "Maharashtra- 400063, India"
    apps.append(application(
        "APPL00090001", "Mumbai 2", "01-07-2026",
        [firm(90011, "Shree Sai Enterprises", "Partnership Firm", "primary", "01-JUL-04",
              ids=[("PAN", "AAHFS4821K"), ("GST_NO", "27AAHFS4821K1ZP"),
                   ("UDHYAM_REGISTRATION", "UDYAM-MH-18-0012345")],
              email="shreesai.ent@example.com", addresses=[("OfficeAddress", goregaon_office)]),
         person(90012, "RAMESH K PATIL", "Co-Applicant", "01-JUN-61", ids=[("PAN", "AKRPP5408N"), ("PAN", "AKRPP5408N")],
                pan="AKRPP5408N", email="ramesh.patil61@example.com",
                addresses=[("ResidentialAddress", "B-408 Neel Kamal Tower, Shanti Park Complex, Nr Aarey Road Signal, "
                                                  "Goregaon East, Mumbai, Maharashtra- 400063, India"),
                           ("PermanentAddress", "At Post Kasegaon, Tal Walwa, Sangli, Maharashtra- 415404, India")]),
         person(90013, "Ajit Dattatray Salunkhe", "Co-Applicant", "11-JAN-51", "BQSPS6215K",
                email="ad.salunkhe@example.net",
                addresses=[("ResidentialAddress", "Flat no C/305, Sagar Darshan CHS Ltd, Juhu Lane, Gulmohar "
                                                  "Road, Road No 4, Andheri West, Mumbai, Maharashtra- 400058, India")])],
        [collateral("COLL00090001", "CMS09090001", "COMMERCIAL",
                    "Office No. 21, 2nd Floor, Sunbeam Business CSL, Near Udupi Restaurant, Film City "
                    "Road, Dindoshi, Goregaon (E), Mumbai 400063, CTS No. 248/B/3/2pt, Mumbai, Maharashtra- 400063, "
                    "India", "400063", [(" MR. RAMESH KRISHNA PATIL", 100)], built=620, carpet=520,
                    purpose="Rented", market=12500000,
                    pdf=report("B", "Shree Sai Enterprises (through its partner Ramesh K. Patil)", "Mr. Ramesh K. Patil",
                               "Office No. 21 admeasuring 48.31 sq.mtr carpet area on the 2nd floor in the building "
                               "known as Sunbeam Business CSL of the project known as Sunbeam Business CSL "
                               "constructed on the land bearing CTS No. 248/B/3/2pt of Village Malad, Taluka Borivali, "
                               "Film City Road, Dindoshi, Goregaon (East), Mumbai 400063",
                               nature="Commercial", initiated="09/06/2026", issued="12/06/2026; 16/06/2026"))],
        [loan(900101, "A", "01-JUL-26", 8800000)],
        case="Golden record: partnership firm + 2 individuals, PAN listed twice, owner verified against a "
             "B-style title report (initials: 'Ramesh K. Patil' = 'RAMESH KRISHNA PATIL'), CTS number in both."))

    # D02 - joint owners with Mr/Mrs and initials; survey number only in the title report.
    thane_home = "Flat No. 1102, 11th Floor, C Wing, Shubh Laxmi Heights, Pokharan Road No. 1, Vartak Nagar, Thane West, " \
                 "Thane, Maharashtra- 400606, India"
    apps.append(application(
        "APPL00090002", "Thane", "18-08-2026",
        [person(90021, "SUNIL GANPAT KADAM", "primary", "23-SEP-78", "AKZPK4410M", email="sunil.kadam78@example.com",
                addresses=[("ResidentialAddress", thane_home)]),
         person(90022, "ANITA SUNIL KADAM", "Co-Applicant", "05-DEC-82", "BQRPK7788D", first="ANITA", middle="WO",
                last="KADAM", addresses=[("ResidentialAddress", thane_home)])],
        [collateral("COLL00090002", "CMS09090002", "RESIDENTIAL",
                    "Flat No. 1102, 11th Floor, C Wing, Shubh Laxmi Heights Co-op. Hsg. Soc. Ltd., Pokharan Road No. 1, "
                    "Vartak Nagar, Thane (West), Thane, Maharashtra- 400606, India", "400606",
                    [(" Mr. Sunil Ganpat Kadam", 50), ("Mrs. Anita Sunil Kadam", 50)], built=780, carpet=655.5,
                    classification="Residential Flat",
                    pdf=report("A", "Sunil Kadam", "Mr. Sunil G. Kadam & Mrs. Anita S. Kadam.",
                               "Flat No. 1102 on the 11th Floor, adm. 60.90 Sq. Mtrs., Carpet area, in the Society "
                               "known as ‘Shubh Laxmi Heights Co-op. Hsg. Soc. Ltd.’ constructed on land "
                               "bearing Survey No. 211/2/1, 214/3, 216/5, situated at Village Majiwade, Taluka and "
                               "District Thane, within the limits of Thane Municipal Corporation, Pokharan Road No. 1, "
                               "Vartak Nagar, Thane (West)400 606,", mortgage="Equitable  Mortgage"))],
        [loan(900201, "A", "18-AUG-26", 4500000)],
        case="Joint owners with Mr./Mrs. and a leading space; the title report writes middle names as initials; "
             "survey numbers only in the title report (A-style, PIN '400 606' inside item 1); decimal carpet "
             "area; middle name 'WO' (wife of)."))

    # D03 - Bangalore 'Property bearing No.'; proprietorship with no PAN; no title report.
    apps.append(application(
        "APPL00090003", "Bangalore", "22-07-2026",
        [firm(90031, "SRI LAKSHMI TRADERS", "Sole Proprietorship Firm", "primary", "10-OCT-18",
              ids=[("GST_NO", "29AGHPL5521K1Z5"), ("UDHYAM_REGISTRATION", "UDYAM-KR-03-0045678"),
                   ("Shop/Establishment Number", "123456789012")],
              email="srilakshmitraders@example.org",
              addresses=[("OfficeAddress", "No. 112, 4th Cross, APMC Yard, Yeshwanthpur, Bangalore, Karnataka- "
                                           "560022, India")]),
         person(90032, "N LOKESH", "Co-Applicant", "14-FEB-85", "AGHPL5521K", first="N LOKESH", middle=None, last=" ",
                addresses=[("ResidentialAddress", "H N 27/B, 5th Cross, Goraguntepalya, Bangalore, Karnataka- 560022, "
                                                  "India")]),
         person(90033, "Kavya L", "Co-Applicant", "30-AUG-90", "CQWPL3321A", email="kavya-l@example.net",
                addresses=[("ResidentialAddress", "H N 27/B, 5th Cross, Goraguntepalya, Bangalore, Karnataka- 560022, "
                                                  "India")])],
        [collateral("COLL00090003", "CMS09090003", "RESIDENTIAL",
                    "Property bearing No. 318, BBMP Katha No. 422/309, EPID No. 7104452196, Ward No. 38, 3rd Main "
                    "Road, Shivaganga Layout, Goraguntepalya, Yeshwanthpur Hobli, Bangalore North Taluk, Bangalore, "
                    "Karnataka- 560022, India", "560022", [("N Lokesh", 100)], built=1450, carpet=1200)],
        [loan(900301, "A", "22-JUL-26", 6500000)],
        case="Bangalore BBMP address (Katha, EPID, Ward, Hobli, Taluk): house number found, but not whether it is "
             "a flat or a house; no survey/plot number; no title report (SI type typed by hand, owners not verified); "
             "proprietorship with only GST/Udyam/Shop numbers (no PAN, so a CKYC number is needed); Last Name ' '; "
             "e-mail with a hyphen (not allowed by CERSAI)."))

    # D04 - Hyderabad house with two house numbers; NO PAN customer; registered mortgage.
    hyd_home = "H.No. 2-4-118, Kushaiguda, Kapra, Hyderabad, Telangana- 500062, India"
    apps.append(application(
        "APPL00090004", "Hyderabad", "03-08-2026",
        [person(90041, "MOHD IRFAN", "primary", "09-JUL-84", "BJUPI6612R", addresses=[("ResidentialAddress", hyd_home)]),
         person(90042, "SALMA BEGUM", "Co-Applicant", "17-MAR-88", constitution="NO PAN",
                ids=[("Voter's ID", "ZXQ1234567")], mobile=None, phone="5848012345",
                addresses=[("ResidentialAddress", hyd_home)])],
        [collateral("COLL00090004", "CMS09090004", "RESIDENTIAL",
                    "H.No. 2-4-118 and 2-4-119, Situated at Kushaiguda, Kapra Village, Kapra Mandal, Under GHMC "
                    "Circle, Survey No. 88/B, Hyderabad, Telangana- 500062, India", "500062",
                    [("Mohd Irfan", 50), ("Salma Begum", 50)], built=2400, carpet=2000,
                    pdf=report("B", "Mohd Irfan", "Mohd Irfan and Salma Begum",
                               "House No. 2-4-118 and 2-4-119 admeasuring 222.96 sq.mtr built up area consisting of "
                               "ground and first floors constructed on the land bearing Survey No. 88/B situated at "
                               "Kushaiguda, Kapra Village, Kapra Mandal, Medchal-Malkajgiri District, Hyderabad 500 062",
                               mortgage="Registered Mortgage", status="Positive")),
         collateral("COLL00090033", "CMS09090033", "PLOT",
                    "Plot No. 41, Survey No. 307/1, Lakshmi Enclave, Keesara Village, Keesara Mandal, Medchal-Malkajgiri "
                    "District, Hyderabad, Telangana- 501301, India", "501301", [("Salma Begum", 100)], carpet=2250,
                    purpose="Vacant",
                    pdf=report("A", "Mohd Irfan", "Smt. Salma Begum",
                               "Open Plot No. 41, adm. 209.03 Sq. Mtrs., Plot area, in the layout known as Lakshmi "
                               "Enclave, bearing Survey No. 307/1, situated at Keesara Village, Keesara Mandal, "
                               "Medchal-Malkajgiri District, Telangana 501 301.", nature="Residential Open Plot.",
                               mortgage="Registered Mortgage"))],
        [loan(900401, "A", "03-AUG-26", 7200000)],
        case="Two collaterals: the house is owned by both borrowers, the plot by one of them (Salma Begum) - so "
             "IS_ASSET_OWNER differs between the two SIs. Two house numbers ('H.No. 2-4-118 and 2-4-119' -> "
             "2-4-118/2-4-119); customer with constitution 'NO PAN' and only a Voter's ID (CKYC number needed); no "
             "mobile in CKYC, so COLLATERAL PHONE_NUMBER is used; both title reports say Registered Mortgage (SI "
             "type 18); the plot's title report says 'Residential', so the PLOT is asset type 1."))

    # D05 - PLOT with a title report that says it is residential.
    apps.append(application(
        "APPL00090005", "Jaipur", "26-06-2026",
        [person(90051, "RAJENDRA PRASAD SHARMA", "primary", "02-OCT-70", "AOLPS9087H",
                addresses=[("ResidentialAddress", "Plot No. 77, Gopalpura Bypass, Near Mahesh Nagar Phatak, Jaipur, "
                                                  "Rajasthan- 302018, India")]),
         person(90052, "priya sharma", "Co-Applicant", "21-APR-95", "FKLPS2231B", email="priya.sharma95@example.com",
                addresses=[("ResidentialAddress", "Plot No. 77, Gopalpura Bypass, Near Mahesh Nagar Phatak, Jaipur, "
                                                  "Rajasthan- 302018, India")]),
         person(90053, "MAHESH CHAND SHARMA", "Guarantor", "15-JAN-66", "AZXPS1122C",
                addresses=[("ResidentialAddress", "61, Shanti Nagar, Durgapura, Jaipur, Rajasthan- 302018, India")])],
        [collateral("COLL00090005", "CMS09090005", "PLOT",
                    "Plot No. 132, Shri Krishna Nagar Scheme, Near Water Tank, Jagatpura, Jaipur, Rajasthan- 302017, India",
                    "302017", [("Rajendra Prasad Sharma", 100)], carpet=1800, purpose="Vacant",
                    pdf=report("A", "Rajendra Prasad Sharma", "Shri Rajendra Prasad Sharma",
                               "Residential Plot No. 132, adm. 167.22 Sq. Mtrs., Plot area, in the scheme known as Shri "
                               "Krishna Nagar, situated at Village Jagatpura, Tehsil Sanganer, District Jaipur, Rajasthan "
                               "302 017.", nature="Residential Plot.", initiated="10/06/2026", issued="15/06/2026"))],
        [loan(900501, "A", "26-JUN-26", 3000000)],
        case="Property type PLOT: residential or commercial is taken from title report item 2 ('Residential Plot' -> "
             "1, land without structure -> 3); a guarantor (not a borrower, so no BOR row); an all-lower-case "
             "name; no project name anywhere (typed by hand)."))

    # D06 - PLOT with no title report; proprietorship tagged 'individual'.
    apps.append(application(
        "APPL00090006", "AHMEDABAD", "09-09-2026",
        [firm(90061, "PATEL AGRO INDUSTRIES", "Sole Proprietorship Firm", "primary", None,
              ids=[("PAN", "AMPPP4455G"), ("UDYAM", "UDYAM-GJ-17-0098765")], applicant_type="individual",
              dob="04-NOV-79",
              addresses=[("OfficeAddress", "Survey No. 418/2, Halol GIDC Road, Halol, Panchmahal, Gujarat- 389350, "
                                           "India")]),
         person(90062, "BHAVESH PATEL", "Co-Applicant", "04-NOV-79", "AMPPP4455G",
                addresses=[("ResidentialAddress", "21, Narmada Residency, Halol, Panchmahal, Gujarat- 389350, India")])],
        [collateral("COLL00090006", "CMS09090006", "PLOT",
                    "Plot No. 21 to 23, Narmada Residency, Nr. Shiv Shakti Apartment, Survey No. 418/2, Halol, "
                    "Panchmahal, Gujarat- 389350, India", "389350", [("Bhavesh Patel", 100)], carpet=2700.75,
                    nature="GRAM PANCHAYAT", purpose="Vacant")],
        [loan(900601, "A", "09-SEP-26", 2500000)],
        case="PLOT with no title report: the asset type is typed by hand (3 Other, 'Plot'); a proprietorship tagged "
             "'individual' in CKYC Applicant Type (so CKYC has a date of birth but no incorporation date: "
             "registration date needed); proprietor and firm share one PAN; Gram Panchayat land."))

    # D07 - private limited company with CIN; khasra number only in the title report.
    apps.append(application(
        "APPL00090007", "Ghaziabad", "14-08-2026",
        [firm(90071, "AARAV INFRA PROJECTS PVT. LTD.", "Private Limited Company", "primary", "15-MAR-12",
              ids=[("PAN", "AAMCA1234C"), ("CIN No", "U45200UP2012PTC012345"), ("GST_NO", "09AAMCA1234C1Z2")],
              email="accounts.aaravinfra@example.com",
              addresses=[("OfficeAddress", "C-21, 3rd Floor, RDC Raj Nagar, Ghaziabad, Uttar Pradesh- 201002, India"),
                         ("AlternateBusinessAddress", "Plot No. 7, Site IV, Sahibabad Industrial Area, Ghaziabad, "
                                                      "Uttar Pradesh- 201010, India")]),
         person(90072, "DEEPAK KUMAR VERMA", "Co-Applicant", "19-JUL-76", "ACVPV5566D",
                addresses=[("ResidentialAddress", "Flat No. 704, Silver Crest Avenue, Indirapuram, Ghaziabad, Uttar "
                                                  "Pradesh- 201014, India")]),
         person(90073, "POOJA VERMA", "Co-Applicant", "08-JAN-80", "BDTPV7788E", email="pooja-verma@example.net",
                coll_email="poojaverma80@example.com",
                addresses=[("ResidentialAddress", "Flat No. 704, Silver Crest Avenue, Indirapuram, Ghaziabad, Uttar "
                                                  "Pradesh- 201014, India")])],
        [collateral("COLL00090007", "CMS09090007", "RESIDENTIAL",
                    "Flat No. 704, 7th Floor, Tower B-2, Silver Crest Avenue, Plot No. GH-01, Sector 4, Abhay Khand, "
                    "Indirapuram, Ghaziabad, Uttar Pradesh- 201014, India", "201014",
                    [("Deepak Kumar Verma", 50), ("Pooja Verma", 50)], built=1150, carpet=890,
                    pdf=report("B", "Aarav Infra Projects Pvt. Ltd.", "Mr. Deepak Kumar Verma and Mrs. Pooja Verma",
                               "Flat No. 704 admeasuring 1150 Sq. Ft. super area on the 7th floor in Tower B-2 of the "
                               "project known as Silver Crest Avenue constructed on the land bearing Khasra No. 1245 "
                               "situated at Village Makanpur, Abhay Khand, Indirapuram, District Ghaziabad")),
         collateral("COLL00090034", "CMS09090034", "COMMERCIAL",
                    "Office No. 305, 3rd Floor, Vaibhav Trade Plaza, Plot No. 4, Vaishali Sector 1, Ghaziabad, Uttar "
                    "Pradesh- 201010, India", "201010", [("Aarav Infra Projects Pvt Ltd", 100)], built=720,
                    carpet=600, purpose="Self Occupied",
                    pdf=report("B", "Aarav Infra Projects Pvt. Ltd.", "M/s. Aarav Infra Projects Private Limited",
                               "Office No. 305 admeasuring 55.74 sq.mtr carpet area on the 3rd floor in the building "
                               "known as Vaibhav Trade Plaza of the project known as Vaibhav Trade Plaza constructed on the land "
                               "bearing Plot No. 4 situated at Vaishali Sector 1, District Ghaziabad",
                               nature="Commercial"))],
        [loan(900701, "A", "14-AUG-26", 5200000)],
        case="Two collaterals with different owners: the flat belongs to the two directors, the office to the "
             "company itself ('Private Limited' in the title report = 'Pvt Ltd' in COLLATERAL). Private limited "
             "company with its CIN in CKYC; office and alternate business addresses; e-mail with a hyphen in CKYC, "
             "valid one in COLLATERAL (used); plot number from the LMS address and khasra (survey) number from the "
             "title report; title-report area matches the built-up area, not the carpet area."))

    # D08 - Haryana revenue numbers; area only in the title report; PIN in two districts; COLLATERAL
    # misses an owner (shares add up to 50%), the title report names her.
    apps.append(application(
        "APPL00090008", "Gurugram", "29-08-2026",
        [person(90081, "AMIT MALHOTRA", "primary", "12-DEC-83", "AJKPM3344F",
                addresses=[("ResidentialAddress", "Tower 4, Flat 1203, Palm Court, Sector 42, Gurgaon, Haryana- "
                                                  "122002, India")]),
         person(90082, "NEHA MALHOTRA", "Co-Applicant", "27-MAY-86", "BLMPM5566G",
                addresses=[("ResidentialAddress", "Tower 4, Flat 1203, Palm Court, Sector 42, Gurgaon, Haryana- "
                                                  "122002, India")])],
        [collateral("COLL00090008", "CMS09090008", "RESIDENTIAL",
                    "Residential Apartment No. 9, on Second Floor, in Block-C, Greenwood City, Sohna Road, Khewat No. "
                    "123, Mustatil No. 45, Killa No. 12/2, Village Badshahpur, District Gurgaon, Gurgaon, Haryana- "
                    "122502, India", "122502", [("Amit Malhotra", 50)], status="Joint",
                    pdf=report("A", "Amit Malhotra", "Mr. Amit Malhotra and Mrs. Neha Malhotra",
                               "Residential Apartment No. 9 on the Second Floor, adm. 111.48 Sq. Mtrs., Carpet area, "
                               "in Block-C of the project known as Greenwood City, constructed on land bearing Khewat "
                               "No. 123, Mustatil No. 45, Killa No. 12/2, situated at Village Badshahpur, Tehsil and "
                               "District Gurgaon, Haryana."))],
        [loan(900801, "A", "29-AUG-26", 9500000)],
        case="Haryana khewat/mustatil/killa numbers; no area in COLLATERAL, so the title report's area in sq. m is "
             "used (unit 7); PIN 122502 belongs to two districts (Gurgaon, Rewari) and the address names Gurgaon. "
             "COLLATERAL lists one owner with a 50% share and Ownership Status Joint, so an owner is missing; the "
             "title report names Amit and his wife Neha (a co-applicant), so Neha is added from the title report "
             "and the Check and correct sheet says so."))

    # D09 - HUF; Delhi 'Property bearing No.', khasra; mixed usage; no title report.
    apps.append(application(
        "APPL00090009", "Delhi 1", "07-07-2026",
        [firm(90091, "SHREE BALAJI (HUF)", "Hindu Undivided Company", "primary", "01-APR-98", "AAHHS5678B",
              addresses=[("OfficeAddress", "B-45, Ground Floor, Mayur Vihar Phase 1, East Delhi, Delhi- 110091, "
                                           "India")]),
         person(90092, "SANJAY GUPTA", "Co-Applicant", "30-SEP-72", "AFGPG9900J",
                addresses=[("ResidentialAddress", "B-45, First Floor, Mayur Vihar Phase 1, East Delhi, Delhi- 110091, "
                                                  "India")])],
        [collateral("COLL00090009", "CMS09090009", "MIXED USAGE",
                    "Property bearing No. B-45, Ground Floor, Khasra No. 12/13, Block B, Mayur Vihar Phase 1, East "
                    "Delhi, Delhi- 110091, India", "110091", [("Sanjay Gupta", 100)], built=900, carpet=750,
                    purpose="Multi Tenanted")],
        [loan(900901, "A", "07-JUL-26", 6000000)],
        case="HUF with brackets in its name (constitution 'Hindu Undivided Company'); mixed usage (asset type 3); "
             "Delhi 'Property bearing No.' with a khasra number and a block."))

    # D10 - Surat shop, T.P. scheme / final plot; borrower address with a wrong ending.
    apps.append(application(
        "APPL00090010", "SURAT", "21-08-2026",
        [firm(90101, "M/S. KRISHNA TEXTILES", "Sole Proprietorship Firm", "primary", "10-AUG-15",
              ids=[("PAN", "BCDPA4321E"), ("GST_NO", "24BCDPA4321E1Z9")],
              addresses=[("OfficeAddress", "SHOP NO D 14, GROUND FLOOR, OM GANESH CHSL, S N ROAD NO 4, BORIVALI EAST, "
                                           "MUMBAI, Surat, Gujarat- 395007, India")]),
         person(90102, "KISHAN LAL AGARWAL", "Co-Applicant", "11-NOV-69", "BCDPA4321E",
                addresses=[("ResidentialAddress", "A-12, Shanti Kunj Society, Ghod Dod Road, Surat, Gujarat- 395007, "
                                                  "India")])],
        [collateral("COLL00090010", "CMS09090010", "COMMERCIAL",
                    "Shop No. G-9, Ground Floor, Krishna Textile Hub, Block No. 12, T. P. Scheme No. 7, Final Plot "
                    "No. 34 & 35, Nr. Krishna Hospital Circle, Dumbhal, Surat, Gujarat- 395010, India", "395010",
                    [("Kishan Lal Agarwal", 100)], built=300, carpet=240, classification="Commercial Flat",
                    purpose="Rented",
                    pdf=report("B", "Kishan Lal Agarwal (Krishna Textiles)", "Mr. Kishan Lal Agarwal",
                               "Shop No. G-9 admeasuring 22.30 sq.mtr carpet area on the ground floor in the building "
                               "known as Krishna Textile Hub of the project known as Krishna Textile Hub constructed on "
                               "the land bearing Block No. 12, T.P. Scheme No. 7, Final Plot No. 34 & 35 situated at "
                               "Village Dumbhal, Taluka Choryasi, District Surat", nature="Commercial",
                               mortgage="Mortgage by deposit of title deeds"))],
        [loan(901001, "A", "21-AUG-26", 3500000)],
        case="Gujarat T.P. scheme and final plot numbers (34/35) in both sources; shop (commercial, structure "
             "without land); the proprietorship's office address body is in Mumbai but ends with Surat (data "
             "quality: the ending decides city/PIN); 'Mortgage by deposit of title deeds' = equitable (17)."))

    # D11 - LLP (no LLPIN in CKYC); Door No.; usufructuary mortgage (SI type 21).
    apps.append(application(
        "APPL00090011", "Chennai", "30-07-2026",
        [firm(90111, "COROMANDEL FOODS LLP", "Limited Liability Partnership", "primary", "20-JAN-19",
              ids=[("PAN", "AAKFC7788L"), ("GST_NO", "33AAKFC7788L1Z1")],
              addresses=[("OfficeAddress", "No. 5, Industrial Estate, Guindy, Chennai, Tamil Nadu- 600032, India")]),
         person(90112, "SENTHIL KUMAR R", "Co-Applicant", "03-MAR-81", "BXRPR1234K",
                addresses=[("ResidentialAddress", "Door No. 12/4, 2nd Street, Anna Nagar East, Chennai, Tamil Nadu- "
                                                  "600040, India")]),
         person(90113, "MEENA SENTHIL", "Co-Applicant", "16-JUN-84", "CZYPS4321L",
                addresses=[("ResidentialAddress", "Door No. 12/4, 2nd Street, Anna Nagar East, Chennai, Tamil Nadu- "
                                                  "600040, India")])],
        [collateral("COLL00090011", "CMS09090011", "COMMERCIAL",
                    "Door No. 12/4, Old No. 7, 2nd Street, Anna Nagar East, Chennai, Tamil Nadu- 600040, India",
                    "600040", [("Senthil Kumar R", 50), ("Meena Senthil", 50)], built=1200, carpet=1000,
                    pdf=report("A", "Coromandel Foods LLP", "Mr. R. Senthil Kumar & Mrs. Meena Senthil",
                               "Door No. 12/4 (Old No. 7) on the Ground Floor, adm. 1200 Sq. Fts., Built up area, "
                               "situated at 2nd Street, Anna Nagar East, Chennai 600 040.", nature="Commerical.",
                               mortgage="Usufructuary Mortgage"))],
        [loan(901101, "A", "30-JUL-26", 4000000)],
        case="LLP with no LLPIN in CKYC (registration number needed); Tamil Nadu 'Door No.' (kind of unit "
             "unknown); owner initial written first in the title report ('R. Senthil Kumar'); 'Commerical' typo; "
             "'Usufructuary Mortgage' -> SI type 21 with SI_TYPE_OTHERS."))

    # D12 - two flats in one collateral; four owners; guarantor; the two disbursal dates differ.
    joshi_home = "Flat No. 1301, 13th Floor, A Wing, Shanti Sadan CHS Ltd, Veera Desai Road, Andheri West, Mumbai, " \
                 "Maharashtra- 400058, India"
    apps.append(application(
        "APPL00090012", "Mumbai 1", "05-08-2026",
        [person(90121, "MAHESH CHANDRA JOSHI", "primary", "25-JAN-62", "AEWPJ1111A",
                addresses=[("ResidentialAddress", joshi_home)]),
         person(90122, "REKHA MAHESH JOSHI", "Co-Applicant", "14-FEB-66", "AEWPJ2222B",
                addresses=[("ResidentialAddress", joshi_home)]),
         person(90123, "ROHAN MAHESH JOSHI", "Co-Applicant", "09-SEP-94", "BGHPJ3333C",
                addresses=[("ResidentialAddress", joshi_home)]),
         person(90124, "ISHA JOSHI", "Co-Applicant", "15-MAR-05", "CKLPJ4444D",
                addresses=[("ResidentialAddress", joshi_home)]),
         person(90125, "ANIL JOSHI", "Guarantor", "02-FEB-60", "AEWPJ5555E",
                addresses=[("ResidentialAddress", joshi_home)])],
        [collateral("COLL00090012", "CMS09090012", "RESIDENTIAL",
                    "Flat No. 1301 & 1302, 13th Floor, A Wing, Shanti Sadan CHS Ltd, Veera Desai Road, Azad Nagar, "
                    "Andheri West, Mumbai, Maharashtra- 400058, India", "400058",
                    [(" Mr. Mahesh Chandra Joshi", 25), ("Mrs. Rekha M. Joshi", 25), ("Mr. Rohan Joshi", 25),
                     ("Ms. Isha Joshi", 25)], built=1550, carpet=1300, market=42000000,
                    pdf=report("B", "Mahesh Chandra Joshi", "Mahesh Chandra Joshi, Rekha Mahesh Joshi, Rohan Mahesh "
                               "Joshi and Isha Mahesh Joshi",
                               "Flat No. 1301 and 1302 admeasuring 120.50 sq.mtr carpet area on the 13th floor in A "
                               "Wing of the building known as Shanti Sadan of the Society known as 'Shanti Sadan "
                               "Co-operative Housing Society Ltd.', constructed on land bearing C.T.S. No. 934, 934/1 "
                               "to 934/5 of Village Andheri, Taluka Andheri, Veera Desai Road, Azad Nagar, Andheri "
                               "(West), Mumbai 400058", status="Positive")),
         collateral("COLL00090035", "CMS09090035", "COMMERCIAL",
                    "Shop No. 4, Ground Floor, Joshi Arcade, S. V. Road, Andheri West, Mumbai, Maharashtra- 400058, "
                    "India", "400058", [], built=450, carpet=380, purpose="Rented",
                    pdf=report("A", "Mahesh Chandra Joshi", "Mr. Mahesh C. Joshi",
                               "Shop No. 4 on the Ground Floor, adm. 35.30 Sq. Mtrs., Carpet area, in the building "
                               "known as \u2018Joshi Arcade\u2019, constructed on land bearing C.T.S. No. 812, "
                               "situated at Village Andheri, Taluka Andheri, S. V. Road, Andheri (West), Mumbai "
                               "400 058.", nature="Commercial."))],
        [loan(901201, "A", "01-AUG-26", 15000000)],
        case="Two collaterals: the flats are owned by all four borrowers, the shop by one of them - and COLLATERAL "
             "has no owner rows for the shop, so its owner comes from its title report. Two flats in one "
             "collateral ('Flat No. 1301 & 1302' -> 1301/1302); four owners at 25% each with Mr/Mrs/Ms; a "
             "guarantor; CKYC First Disbursal Date differs from COLLATERAL Disbursal Date (CKYC used); a 2-digit "
             "birth year that is 20YY (15-MAR-05 -> 2005)."))

    # D13 - one application, two collaterals.
    pune_home = "Row House No. 7, Shubh Kalyan Society, Wakad, Pune, Maharashtra- 411057, India"
    apps.append(application(
        "APPL00090013", "Pune", "11-09-2026",
        [person(90131, "PRAKASH DESHMUKH", "primary", "18-AUG-75", "AHJPD6677K",
                addresses=[("ResidentialAddress", pune_home)]),
         person(90132, "SWATI DESHMUKH", "Co-Applicant", "02-APR-79", "BNMPD8899L",
                addresses=[("ResidentialAddress", pune_home)]),
         firm(90133, "DESHMUKH ENTERPRISES", "Sole Proprietorship Firm", "Co-Applicant", "01-JUN-10", "AHJPD6677K",
              addresses=[("OfficeAddress", "Shop No. 4, Wakad Chowk, Wakad, Pune, Maharashtra- 411057, India")])],
        [collateral("COLL00090013", "CMS09090013", "RESIDENTIAL",
                    "Row House No. 7, Shubh Kalyan Society, S. No. 45/2, Wakad, Pune, Maharashtra- 411057, India",
                    "411057", [("Prakash Deshmukh", 50), ("Swati Deshmukh", 50)], built=1800, carpet=1500,
                    classification="Row House",
                    pdf=report("A", "Prakash Deshmukh", "Mr. Prakash Deshmukh & Mrs. Swati Deshmukh",
                               "Row House No. 7, adm. 1800 Sq. Fts., Built up area, in the Society known as ‘Shubh "
                               "Kalyan Co-operative Housing Society Ltd.’, constructed on land bearing S. No. "
                               "45/2, situated at Village Wakad, Taluka Mulshi, District Pune.",
                               owners_item12="Mr. Prakash Deshmukh")),
         collateral("COLL00090014", "CMS09090014", "RESIDENTIAL",
                    "Flat No. B-302, 3rd Floor, Kingsway Grove, Survey No. 12/1A, Pimple Saudagar, Pune, Maharashtra- "
                    "411027, India", "411027", [("Prakash Deshmukh", 100)], built=1100, carpet=950, purpose="Rented")],
        [loan(901301, "A", "11-SEP-26", 9000000)],
        case="Two collaterals on one application (the borrowers are listed once in the required-details file); "
             "row house = structure with land; the title report's item 12 names fewer owners than item 4 "
             "(warning)."))

    # D14 + D15 - the same property (Global collateral No) secures two applications.
    express_zone = "Unit No. 305, 3rd Floor, Pinnacle Zone, Western Express Highway, Malad East, Mumbai, Maharashtra- " \
                   "400097, India"
    harish = person(90141, "HARISH MEHTA", "primary", "07-JUN-68", "AAQPM1212M", email="harish.mehta68@example.com",
                    addresses=[("ResidentialAddress", "B-1204, Harmony Heights, JVLR, Andheri East, Mumbai, "
                                                      "Maharashtra- 400060, India")])
    apps.append(application(
        "APPL00090014", "Mumbai 3", "15-07-2026", [harish],
        [collateral("COLL00090016", "CMS09090099", "COMMERCIAL", express_zone, "400097", [("Harish Mehta", 100)],
                    built=900, carpet=760, purpose="Rented")],
        [loan(901401, "A", "15-JUL-26", 11000000)],
        case="The same property (same Global collateral No) also secures APPL00090015, recorded there under its "
             "own Collateral ID (a Collateral ID belongs to one application only): TYPE_OF_CHARGE is "
             "left blank and typed by hand (FIRST CHARGE); single borrower; the city is MALAD (more specific than the general 'Mumbai')."))
    apps.append(application(
        "APPL00090015", "Mumbai 3", "20-08-2026",
        [firm(90151, "MEHTA TRADING CO.", "Partnership Firm", "primary", "12-DEC-01", "AAAFM3434N",
              email="accounts.mehtatrading@example.com", addresses=[("OfficeAddress", express_zone)]),
         dict(harish, ctype="Co-Applicant")],
        [collateral("COLL00090017", "CMS09090099", "COMMERCIAL", express_zone, "400097", [("Harish Mehta", 100)],
                    built=900, carpet=760, purpose="Rented")],
        [loan(901501, "A", "20-AUG-26", 5000000)],
        case="Second loan on the same property as APPL00090014, with its own Collateral ID (TYPE_OF_CHARGE "
             "left blank and typed by hand: SECOND CHARGE); the same customer (same Neo CIF ID) on two applications."))

    # D18 - no owners in COLLATERAL; the title report names them.
    kalyan_home = "Flat No. 3, Om Sai Apartment, Katemanivali, Kalyan East, Thane, Maharashtra- 421306, India"
    apps.append(application(
        "APPL00090018", "KALYAN", "25-08-2026",
        [person(90181, "GANESH PANDURANG MORE", "primary", "01-JAN-80", "ALKPM7878N",
                addresses=[("ResidentialAddress", kalyan_home)]),
         person(90182, "SARIKA GANESH MORE", "Co-Applicant", "13-MAY-85", "CMNPM8989P",
                addresses=[("ResidentialAddress", kalyan_home)])],
        [collateral("COLL00090021", "CMS09090021", "RESIDENTIAL",
                    "Flat No. 3, Ground Floor, Om Sai Apartment, Near Shivaji Chowk, Katemanivali, Kalyan East, Thane, "
                    "Maharashtra- 421306, India", "421306", [], built=520, carpet=450,
                    pdf=report("A", "Ganesh More", "Mr. Ganesh P. More",
                               "Flat No. 3 on the Ground Floor, adm. 41.81 Sq. Mtrs., Carpet area, in the building "
                               "known as ‘Om Sai Apartment’ and the Society known as ‘Om Sai Co-op. Hsg. "
                               "Soc. Ltd.’, constructed on land bearing Survey No. 56, Hissa No. 3, situated at "
                               "Village Katemanivali, Taluka Kalyan, District Thane."))],
        [loan(901801, "A", "25-AUG-26", 1800000)],
        case="No owner rows in COLLATERAL: the owner comes from the title report ('Mr. Ganesh P. More' = borrower "
             "GANESH PANDURANG MORE); survey number with Hissa (56/3) from the title report; ground floor = 0."))

    # D19 - no owners anywhere, no title report; PG; 'P.No.'; a customer with no address.
    apps.append(application(
        "APPL00090019", "Jodhpur", "16-07-2026",
        [person(90191, "OM PRAKASH CHOUDHARY", "primary", "05-MAY-71", "AQRPC5656Q",
                addresses=[("ResidentialAddress", "P.No. B-17, Bhagat Colony, Mahamandir, Jodhpur, Rajasthan- 342001, "
                                                  "India")]),
         person(90192, "KAMLA DEVI", "Co-Applicant", "01-JAN-75", "CRTPD6767R", mobile=None,
                addresses=[(None, None)], ckyc_addresses=[])],
        [collateral("COLL00090022", "CMS09090022", "PG",
                    "P.No. B-17, Bhagat Colony, Near Shiv Mandir, Mahamandir, Jodhpur, Rajasthan- 342001, India",
                    "342001", [], built=2200, carpet=1900, purpose="Multi Tenanted")],
        [loan(901901, "A", "16-JUL-26", 4200000)],
        case="No owners in COLLATERAL and no title report: owners must be typed; property type PG (residential); "
             "'P.No.' = plot number; a co-applicant with no address in CKYC or COLLATERAL (address needed)."))

    # D20 - the title report names an extra owner that COLLATERAL does not have.
    vashi = "Flat No. 504, Sai Darshan CHS, Sector 17, Vashi, Navi Mumbai, Thane, Maharashtra- 400703, India"
    apps.append(application(
        "APPL00090020", "VASHI BRANCH", "02-09-2026",
        [person(90201, "ARJUN NAIR", "primary", "22-OCT-88", "BSDPN2323S", addresses=[("ResidentialAddress", vashi)]),
         person(90202, "LAKSHMI NAIR", "Co-Applicant", "30-NOV-90", "CTFPN3434T",
                addresses=[("ResidentialAddress", vashi)])],
        [collateral("COLL00090023", "CMS09090023", "RESIDENTIAL",
                    "Flat No. 504, 5th Floor, Sai Darshan CHS, Sector 17, Vashi, Navi Mumbai, Thane, Maharashtra- "
                    "400703, India", "400703", [("Arjun Nair", 100)], built=850, carpet=700,
                    pdf=report("B", "Arjun Nair", "Mr. Arjun Nair and Mrs. Radha Nair",
                               "Flat No. 504 admeasuring 65.03 sq.mtr carpet area on the 5th floor in the building "
                               "known as Sai Darshan of the Society known as 'Sai Darshan Co-operative Housing Society "
                               "Ltd.', constructed on the land bearing Plot No. 23, Sector 17 of Village Vashi, Navi "
                               "Mumbai, Taluka and District Thane"))],
        [loan(902001, "A", "02-SEP-26", 6800000)],
        case="COLLATERAL names one owner (Arjun Nair, 100%, Single) and the title report two (Arjun and Radha "
             "Nair). COLLATERAL comes first and its shares add up, so Arjun alone is used (he is a borrower, so "
             "the SI is filed); the title report's names are shown in the Check and correct sheet, where the "
             "owners can be typed if the title report is right. Plot number only in the title report. Note: the "
             "law firm's own PIN (400703) matches this property's PIN by chance."))

    # D21 - an owner name that only resembles a borrower's name.
    apps.append(application(
        "APPL00090021", "Hyderabad", "19-08-2026",
        [person(90211, "SHARADA VELPURI", "primary", "12-JUN-87", "BUGPV4545U",
                addresses=[("ResidentialAddress", "Flat No. 202, Sri Sai Residency, Madhapur, Hyderabad, Telangana- "
                                                  "500081, India")]),
         person(90212, "RAVI VELPURI", "Co-Applicant", "25-DEC-84", "AVHPV5656V",
                addresses=[("ResidentialAddress", "Flat No. 202, Sri Sai Residency, Madhapur, Hyderabad, Telangana- "
                                                  "500081, India")])],
        [collateral("COLL00090024", "CMS09090024", "RESIDENTIAL",
                    "Flat No. 202, 2nd Floor, Sri Sai Residency, Plot No. 18 & 19, Survey No. 112, Madhapur, "
                    "Hyderabad, Telangana- 500081, India", "500081", [("Smt. Sarada V", 100)], built=1250, carpet=1050),
         collateral("COLL00090036", "CMS09090036", "RESIDENTIAL",
                    "Flat No. 305, 3rd Floor, Sri Sai Residency, Plot No. 18 & 19, Survey No. 112, Madhapur, "
                    "Hyderabad, Telangana- 500081, India", "500081", [("Smt. Sarada V", 100)], built=1250, carpet=1050,
                    purpose="Rented")],
        [loan(902101, "A", "19-AUG-26", 5600000)],
        case="Two flats with the same owner 'Smt. Sarada V', whose name only resembles borrower 'SHARADA VELPURI': "
             "the tool takes them to be the same person and files both SIs, and lists the decision once in the "
             "Check and correct sheet (an N typed there would apply to both collaterals); both survey and plot "
             "numbers in the address; no title reports."))

    # D22 - the title report belongs to a different property: SI held.
    apps.append(application(
        "APPL00090022", "Vadodara", "27-08-2026",
        [person(90221, "NILESH SHAH", "primary", "08-AUG-82", "AWJPS6767W",
                addresses=[("ResidentialAddress", "Flat No. 101, Shreeji Heights, Alkapuri, Vadodara, Gujarat- 390007, "
                                                  "India")]),
         person(90222, "HETAL SHAH", "Co-Applicant", "19-SEP-85", "CXKPS7878X",
                addresses=[("ResidentialAddress", "Flat No. 101, Shreeji Heights, Alkapuri, Vadodara, Gujarat- 390007, "
                                                  "India")])],
        [collateral("COLL00090025", "CMS09090025", "RESIDENTIAL",
                    "Flat No. 101, 1st Floor, Shreeji Heights, Near Genda Circle, Alkapuri, Vadodara, Gujarat- 390007, "
                    "India", "390007", [("Nilesh Shah", 50), ("Hetal Shah", 50)], built=1300, carpet=1100,
                    pdf=report("A", "Karan Mehta", "Mr. Karan Mehta & Mrs. Priya Mehta",
                               "Flat No. 808 on the 8th Floor, adm. 92.90 Sq. Mtrs., Carpet area, in the Society known "
                               "as ‘Sunrise Towers Co-op. Hsg. Soc. Ltd.’, constructed on land bearing Survey "
                               "No. 77/2, situated at Village Gotri, Taluka and District Vadodara."))],
        [loan(902201, "A", "27-AUG-26", 4800000)],
        case="The PDF named for this collateral is about another property (other applicant, owners and flat): "
             "the right-PDF check fails 3 times, so the SI is held until someone answers PDF OK / WRONG PDF."))

    # D23 - a scanned title report (no text layer); a name with a dot.
    indore = "House No. 213, Scheme No. 78, Vijay Nagar, Indore, Madhya Pradesh- 452010, India"
    apps.append(application(
        "APPL00090023", "INDORE", "12-08-2026",
        [person(90231, "VIKAS JAIN", "primary", "14-APR-83", "AYLPJ8989Y", addresses=[("ResidentialAddress", indore)]),
         person(90232, "RITU JAIN", "Co-Applicant", "03-JUL-86", "CZMPJ9090Z", addresses=[("ResidentialAddress", indore)]),
         person(90233, "RAKESH .", "Co-Applicant", "20-OCT-92", "DAPPR1010A", first="RAKESH", middle=None, last=".",
                addresses=[("ResidentialAddress", indore)])],
        [collateral("COLL00090026", "CMS09090026", "RESIDENTIAL", indore, "452010", [("Vikas Jain", 100)],
                    built=2000, carpet=1700,
                    pdf=report("scanned", "Vikas Jain", "Mr. Vikas Jain",
                               "House No. 213, Scheme No. 78, Vijay Nagar, Indore, adm. 157.93 Sq. Mtrs."))],
        [loan(902301, "A", "12-AUG-26", 5500000)],
        case="The title report is a scan with no text: it is reported and not used (SI type typed by hand, owners not "
             "verified); a customer name with a dot ('RAKESH .') and no Gender in CKYC (typed in the "
             "required-details file); house = structure with land."))

    # D24 - industrial gala; LMS PIN not in the master; area conflict; two active loans.
    apps.append(application(
        "APPL00090024", "Mumbai 2", "06-08-2026",
        [firm(90241, "DEWAN POLYMERS PRIVATE LIMITED", "Private Limited Company", "primary", None,
              ids=[("PAN", "AAJCD4455F"), ("UDHYAM_REGISTRATION", "UDYAM-MH-33-0123456")],
              applicant_type="individual", dob="12-MAY-10",
              addresses=[("OfficeAddress", "Gala No. 27, Mehta and Rao Industrial Estate No. 3, Navghar, Vasai East, "
                                           "Palghar, Maharashtra- 401208, India")]),
         person(90242, "FARHAN DEWAN", "Co-Applicant", "01-DEC-80", "BZNPD1212B",
                addresses=[("ResidentialAddress", "Flat No. 306, Greenfield City, Vasai East, Palghar, Maharashtra- "
                                                  "401208, India")])],
        [collateral("COLL00090027", "CMS09090027", "INDUSTRIAL",
                    "Gala No. 27, Ground Floor, Mehta and Rao Industrial Estate No. 3, Navghar, Vasai East, Palghar, "
                    "Maharashtra- 401208, India", "401299", [("Dewan Polymers Pvt Ltd", 100)], built=1100,
                    carpet=1000, purpose="Self Occupied",
                    pdf=report("B", "Dewan Plastics", "M/s. Dewan Polymers Pvt. Ltd.",
                               "Gala No. 27 admeasuring 60 sq.mtr built up area on the ground floor in the building "
                               "known as Mehta and Rao Industrial Estate No. 3 constructed on the land bearing Survey "
                               "No. 64, Hissa No. 7 situated at Village Navghar, Vasai East, Taluka Vasai, District "
                               "Palghar 401 208", nature="Industrial"))],
        [loan(902401, "A", "06-AUG-26", 5000000), loan(902402, "A", "20-SEP-26", 5500000, 2000000)],
        case="Industrial gala (asset type 3); COLLATERAL Property Pincode 401299 is not in the CERSAI master, so "
             "the PIN in the title report (401 208, also at the end of the LMS address) is used; LMS and title-report areas differ by more than 10%; "
             "two active loan accounts with different sanctioned amounts (the secured amount is typed by hand); "
             "a company tagged 'individual' (no "
             "incorporation date, no CIN); the company itself is the owner ('Pvt Ltd' = 'PRIVATE LIMITED'); the "
             "title report's loan applicant is an old company name (one right-PDF check fails: warning only)."))

    # D25 - individual tagged non_individual; very long addresses; several CKYC addresses.
    long_home = ("Bungalow No. 3, Plot No. 17, Gat No. 145/2, Shree Colony, Behind Shiv Mandir, Near Zilla Parishad "
                 "School, Opposite Kolhapur Road Bus Stop, Next to Gokul Dairy Collection Centre, Gokulshirgaon, Taluka "
                 "Karveer, District Kolhapur, Beside the Old Panchayat Office and the Water Tank on the Main Road, "
                 "Second Lane after the Sugar Factory Gate, Opposite Shree Mahalaxmi Kirana Stores and Ganesh "
                 "Tailors, Kolhapur, Maharashtra- 416234, India")
    apps.append(application(
        "APPL00090025", "Kolhapur", "23-07-2026",
        [person(90251, "SACHIN SHIVAJI JADHAV", "primary", None, "BAQPJ3434C", applicant_type="non_individual",
                inc="05-JUN-85", addresses=[("ResidentialAddress", long_home)],
                ckyc_addresses=[long_home, "Room No. 5, Shivaji Chawl, Laxmipuri, Kolhapur, Maharashtra- 416002, "
                                           "India"]),
         person(90252, "PRIYANKA SACHIN JADHAV", "Co-Applicant", "18-FEB-88", "CBRPJ4545D", mobile=None,
                phone="26543210", addresses=[(None, None)],
                ckyc_addresses=["Flat No. 9, Sai Krupa Apartment, Rajarampuri 5th Lane, Kolhapur, Maharashtra- "
                                "416008, India"])],
        [collateral("COLL00090028", "CMS09090028", "RESIDENTIAL",
                    "Bungalow No. 3, Ground and First Floor, Plot No. 17, Gat No. 145/2, Shree Colony, Behind Shiv "
                    "Mandir, Near Zilla Parishad School, Opposite Kolhapur Road Bus Stop, within the limits of "
                    "Gokulshirgaon Gram Panchayat, Gokulshirgaon, Taluka Karveer, District Kolhapur, bounded on the "
                    "East by Road, on the West by Plot No. 16, on the North by Plot No. 18 and on the South by Open "
                    "Space, Kolhapur, Maharashtra- 416234, India", "416234",
                    [("Sachin Shivaji Jadhav", 50), ("Priyanka Sachin Jadhav", 50)], built=2600, carpet=2200,
                    nature="GRAM PANCHAYAT")],
        [loan(902501, "A", "23-JUL-26", 6200000)],
        case="An individual tagged non_individual (CKYC has an incorporation date instead of a date of birth: "
             "birth date needed); property address of about 450 characters with a boundaries clause (the "
             "neighbours' plot numbers are ignored) and two floors (no floor number); a borrower address too "
             "long for three lines (warning); two CKYC addresses; no mobile and an 8-digit landline in "
             "COLLATERAL (not a mobile); an address row with no ADDRESS_TYPE."))

    # D26 - a guarantor who co-owns the property: third-party mortgage.
    rajkot = "Flat No. 5, Shivam Apartment, 150 Feet Ring Road, Rajkot, Gujarat- 360005, India"
    apps.append(application(
        "APPL00090026", "Rajkot", "10-08-2026",
        [person(90261, "JAYESH PARMAR", "primary", "15-AUG-90", "BCSPP5656E", addresses=[("ResidentialAddress", rajkot)]),
         person(90262, "MANJULA PARMAR", "Guarantor", "10-OCT-62", "ADTPP6767F",
                addresses=[("ResidentialAddress", rajkot)])],
        [collateral("COLL00090029", "CMS09090029", "RESIDENTIAL",
                    "Flat No. 5, 1st Floor, Shivam Apartment, 150 Feet Ring Road, Rajkot, Gujarat- 360005, India",
                    "360005", [("Jayesh Parmar", 50), ("Manjula Parmar", 50)], built=950, carpet=800)],
        [loan(902601, "A", "10-AUG-26", 2800000)],
        case="The co-owner is a guarantor (guarantors are not borrowers), so the SI is a third-party mortgage "
             "and is left out of the .DAT file."))

    # D27 - trust + public limited company; no first disbursal date in CKYC.
    apps.append(application(
        "APPL00090027", "Delhi 2", None,
        [firm(90271, "SANSKRITI EDUCATION TRUST", "Private Trust", "primary", "01-APR-99", "AAATS1234Q",
              addresses=[("OfficeAddress", "Property No. 12, Pocket C, Sector 8, Rohini, North West Delhi, Delhi- "
                                           "110085, India")]),
         firm(90272, "BHARAT SOFTECH LIMITED", "Public Limited Company", "Co-Applicant", "18-SEP-95",
              ids=[("PAN", "AABCB9988K"), ("CIN No", "L72200DL1995PLC067890")],
              addresses=[("OfficeAddress", "Plot No. 5, Okhla Industrial Area Phase 2, South Delhi, Delhi- 110020, "
                                           "India")])],
        [collateral("COLL00090030", "CMS09090030", "COMMERCIAL",
                    "Property No. 12, Pocket C, Sector 8, Rohini, North West Delhi, Delhi- 110085, India", "110085",
                    [("Sanskriti Education Trust", 100)], built=5200, carpet=4800, market=65000000)],
        [loan(902701, "A", "28-JUL-26", 25000000)],
        case="Private trust and public limited company (CIN found); CKYC has no First Disbursal Date, so "
             "COLLATERAL Disbursal Date is used; its title report file is misnamed (7-digit collateral), so it "
             "is reported and not used; a borrower PIN (110020) in two districts, named in the address."))

    # D28 - PIN in two districts, neither named; unusual ID documents.
    apps.append(application(
        "APPL00090028", "Faridabad", "13-08-2026",
        [person(90281, "RAMESH CHAND", "primary", "01-JAN-69", "AFYPC7878G",
                ids=[("PAN", "AFYPC7878G"), ("Aadhar No.", "4321"), ("PASSPORT", "K1234567")],
                addresses=[("ResidentialAddress", "H. No. 77-B, Gali No. 4, Adarsh Nagar, Ballabgarh, Haryana- 121004, India")]),
         firm(90282, "CHAND HARDWARE STORE", "Sole Proprietorship Firm", "Co-Applicant", "01-AUG-08",
              ids=[("Unsupported ", "06AFYPC7878G1Z5"), ("KYC Identification Number", "06AFYPC7878G1Z5"),
                   ("UDYAM", "UDYAM-HR-05-0012345")],
              addresses=[("OfficeAddress", "Shop No. 11, Main Bazar, Adarsh Nagar, Ballabgarh, Haryana- 121004, India")])],
        [collateral("COLL00090031", "CMS09090031", "RESIDENTIAL",
                    "H. No. 77-B, (EWS), area measuring 52.17 sq. yds., Gali No. 4, Adarsh Nagar, "
                    "Ballabgarh, Haryana- 121004, India", "121004", [("Ramesh Chand", 100)], built=450, carpet=400)],
        [loan(902801, "A", "13-AUG-26", 1500000)],
        case="PIN 121004 is in two districts (Faridabad, Palwal) and the address names neither: district and "
             "city needed (property and borrowers); ID types 'Unsupported ' (trailing space), Aadhar, passport, "
             "and a 'KYC Identification Number' that is really a GST number (not 14 digits, so not used)."))

    # D29 - old/new survey numbers with Hissa; three owners with full names in the title report.
    mira = "Flat No. C-2204, Tower 2, Sunrise Parkview, Mira Road East, Thane, Maharashtra- 401107, India"
    apps.append(application(
        "APPL00090029", "Thane", "04-09-2026",
        [person(90291, "IMRAN QURESHI", "primary", "28-FEB-82", "BGRPQ8989H", addresses=[("ResidentialAddress", mira)]),
         person(90292, "SALMA QURESHI", "Co-Applicant", "09-JUN-85", "CHSPQ9090J",
                addresses=[("ResidentialAddress", mira)]),
         person(90293, "YUSUF QURESHI", "Co-Applicant", "11-NOV-55", "AJTPQ1212K",
                addresses=[("ResidentialAddress", mira)])],
        [collateral("COLL00090032", "CMS09090032", "RESIDENTIAL",
                    "Flat No. C-2204, 22nd Floor, Tower 2, Maple, Sunrise Parkview, Old S. No. 318(Pt), New S. No. 27, "
                    "Hissa No. 5, Village Penkarpada, Mira Road East, Thane, Maharashtra- 401107, India", "401107",
                    [("Imran Qureshi", 33.33), ("Salma Qureshi", 33.33), ("Yusuf Qureshi", 33.34)], built=520,
                    carpet=452,
                    pdf=report("B", "Imran Qureshi", "Imran Yusuf Qureshi, Salma Imran Qureshi and Yusuf Ahmad Qureshi",
                               "Flat No. C-2204 admeasuring 41.99 sq.mtr RERA carpet area on the 22nd floor in "
                               "Tower-2(Maple) in the new building known as Sunrise Parkview 1 of the project known as "
                               "Sunrise Parkview constructed on the land bearing Old S. No. 318(Pt) and New S. No. 27 "
                               "and Hissa No. 5 of Village Penkarpada, Taluka and District Thane",
                               transaction="BT + Top Up"))],
        [loan(902901, "A", "04-SEP-26", 7800000)],
        case="Old and new survey numbers with Hissa (only the new one is used: 27/5); three owners at "
             "33.33/33.33/33.34%; the title report lists 'A, B and C' with full names."))

    # D30 - one application, five collaterals: the owners are the same on two, a subset on one, a different
    # combination on one, and one collateral has an owner who is not a borrower.
    sawant_home = "Flat No. 802, 8th Floor, Neelkanth Heights, Pokharan Road No. 2, Thane West, Thane, Maharashtra- " \
                  "400610, India"
    apps.append(application(
        "APPL00090030", "Thane", "16-09-2026",
        [person(90301, "VINOD RAMCHANDRA SAWANT", "primary", "14-JUL-72", "AXSPS3031K",
                email="vinod.sawant72@example.com", addresses=[("ResidentialAddress", sawant_home)]),
         person(90302, "MEGHA VINOD SAWANT", "Co-Applicant", "22-MAR-76", "BYSPS3032L",
                addresses=[("ResidentialAddress", sawant_home)]),
         firm(90303, "SAWANT ELECTRICALS", "Sole Proprietorship Firm", "Co-Applicant", "01-APR-05", "AXSPS3031K",
              addresses=[("OfficeAddress", "Shop No. 6, Ground Floor, Gokul Arcade, Station Road, Naupada, Thane West, "
                                           "Thane, Maharashtra- 400601, India")]),
         person(90304, "RAMCHANDRA GANPAT SAWANT", "Co-Applicant", "03-JAN-48", "ACSPS3034M",
                addresses=[("ResidentialAddress", sawant_home)])],
        [collateral("COLL00090037", "CMS09090037", "RESIDENTIAL",
                    "Flat No. 802, 8th Floor, B Wing, Neelkanth Heights CHS Ltd, Pokharan Road No. 2, Vasant Vihar, "
                    "Thane West, Thane, Maharashtra- 400610, India", "400610",
                    [(" Mr. Vinod Ramchandra Sawant", 50), ("Mrs. Megha Vinod Sawant", 50)], built=1000, carpet=840,
                    classification="Residential Flat",
                    pdf=report("B", "Vinod Ramchandra Sawant", "Mr. Vinod R. Sawant and Mrs. Megha V. Sawant",
                               "Flat No. 802 admeasuring 78.04 sq.mtr carpet area on the 8th floor in B Wing of the "
                               "building known as Neelkanth Heights of the Society known as 'Neelkanth Heights "
                               "Co-operative Housing Society Ltd.', constructed on the land bearing Survey No. 101/1 "
                               "of Village Majiwade, Taluka and District Thane", initiated="02/09/2026",
                               issued="05/09/2026; 09/09/2026")),
         collateral("COLL00090038", "CMS09090038", "COMMERCIAL",
                    "Shop No. 6, Ground Floor, Gokul Arcade, Station Road, Naupada, Thane West, Thane, Maharashtra- "
                    "400601, India", "400601", [("Vinod Ramchandra Sawant", 100)], built=300, carpet=250,
                    purpose="Self Occupied",
                    pdf=report("A", "Vinod Ramchandra Sawant", "Mr. Vinod R. Sawant",
                               "Shop No. 6 on the Ground Floor, adm. 23.23 Sq. Mtrs., Carpet area, in the building "
                               "known as \u2018Gokul Arcade\u2019, constructed on land bearing C.T.S. No. 45/2, "
                               "situated at Village Naupada, Taluka and District Thane, Station Road, Thane (West) "
                               "400 601.", nature="Commercial.")),
         collateral("COLL00090039", "CMS09090039", "RESIDENTIAL",
                    "Flat No. 101, 1st Floor, Sai Shraddha Apartment, Near Kopri Bridge, Kopri Colony, Thane East, "
                    "Thane, Maharashtra- 400603, India", "400603", [("Mr. Vinod Sawant", 50), ("Mrs. Megha Sawant", 50)],
                    built=650, carpet=540, purpose="Rented"),
         collateral("COLL00090040", "CMS09090040", "INDUSTRIAL",
                    "Gala No. 14, 1st Floor, Shree Ganesh Industrial Estate, Survey No. 36, Village Pimpri, Bhiwandi, "
                    "Thane, Maharashtra- 421302, India", "421302",
                    [("Megha Vinod Sawant", 50), ("Ramchandra Ganpat Sawant", 50)], built=1500, carpet=1350,
                    purpose="Rented",
                    pdf=report("A", "Vinod Ramchandra Sawant", "Mrs. Megha V. Sawant & Mr. Ramchandra G. Sawant",
                               "Gala No. 14 on the 1st Floor, adm. 125.42 Sq. Mtrs., Carpet area, in the building "
                               "known as \u2018Shree Ganesh Industrial Estate\u2019, constructed on land bearing "
                               "Survey No. 36, situated at Village Pimpri, Taluka Bhiwandi, District Thane.",
                               nature="Industrial.")),
         collateral("COLL00090041", "CMS09090041", "PLOT",
                    "Plot No. 23, Gat No. 118, Village Shahapur, Taluka Shahapur, Thane, Maharashtra- 421601, India",
                    "421601", [("Vinod Ramchandra Sawant", 50), ("Sunanda Ramchandra Sawant", 50)], carpet=5400,
                    nature="GRAM PANCHAYAT", purpose="Vacant")],
        [loan(903001, "A", "16-SEP-26", 18500000)],
        case="One application, five collaterals, one loan (each SI carries the full sanctioned amount). Owners: "
             "the same two borrowers on COLL00090037 and COLL00090039 (written differently in COLLATERAL); only "
             "Vinod on COLL00090038; Megha with her father-in-law Ramchandra on COLL00090040; Vinod with his mother "
             "Sunanda - not a borrower - on COLL00090041, which alone becomes a third-party mortgage and is left "
             "out, while the other four are filed. IS_ASSET_OWNER of each borrower differs from SI to SI. The "
             "proprietorship shares Vinod's PAN. Three of the five collaterals have title reports."))

    # D31 - two flats of one application whose title reports were saved under each other's names.
    kulkarni_home = "Flat No. 101, Sai Krupa Apartment, Karve Nagar, Pune, Maharashtra- 411052, India"
    report_101 = ("Flat No. 101 on the 1st Floor, adm. 60.39 Sq. Mtrs., Carpet area, in the Society known as "
                  "\u2018Sai Krupa Co-op. Hsg. Soc. Ltd.\u2019, constructed on land bearing S. No. 14/2, situated at "
                  "Karve Nagar, Taluka Haveli, District Pune.")
    report_202 = report_101.replace("Flat No. 101 on the 1st Floor", "Flat No. 202 on the 2nd Floor")
    apps.append(application(
        "APPL00090031", "Pune", "08-09-2026",
        [person(90311, "ASHOK BHIKAJI KULKARNI", "primary", "19-FEB-70", "ALKPK3111N",
                addresses=[("ResidentialAddress", kulkarni_home)]),
         person(90312, "SEEMA ASHOK KULKARNI", "Co-Applicant", "07-AUG-74", "BMKPK3112P",
                addresses=[("ResidentialAddress", kulkarni_home)])],
        [collateral("COLL00090042", "CMS09090042", "RESIDENTIAL",
                    "Flat No. 101, 1st Floor, Sai Krupa Apartment, S. No. 14/2, Karve Nagar, Pune, Maharashtra- "
                    "411052, India", "411052", [("Ashok Bhikaji Kulkarni", 50), ("Seema Ashok Kulkarni", 50)],
                    built=780, carpet=650,
                    pdf=report("A", "Ashok Kulkarni", "Mr. Ashok B. Kulkarni & Mrs. Seema A. Kulkarni", report_202)),
         collateral("COLL00090043", "CMS09090043", "RESIDENTIAL",
                    "Flat No. 202, 2nd Floor, Sai Krupa Apartment, S. No. 14/2, Karve Nagar, Pune, Maharashtra- "
                    "411052, India", "411052", [("Ashok Bhikaji Kulkarni", 50), ("Seema Ashok Kulkarni", 50)],
                    built=780, carpet=650, purpose="Rented",
                    pdf=report("A", "Ashok Kulkarni", "Mr. Ashok B. Kulkarni & Mrs. Seema A. Kulkarni", report_101))],
        [loan(903101, "A", "08-SEP-26", 6400000)],
        case="Two flats of one application whose title reports were swapped when saved (the report named for flat "
             "101 describes flat 202, and the other way round). The loan applicant and the owners match for both, "
             "so only the flat number gives it away: because it is the flat of the other collateral of the same "
             "application, both SIs are held until someone answers PDF OK / WRONG PDF (or renames the files and "
             "runs step 1 again)."))

    # D32 - husband and wife are both borrowers and both owners; the CKYC names are written surname
    # first, and each name contains the other's given name.
    desai_home = "B-12, Shreeji Bungalows, Near Thaltej Cross Roads, Thaltej, Ahmedabad, Gujarat- 380054, India"
    apps.append(application(
        "APPL00090032", "AHMEDABAD", "03-09-2026",
        [person(90321, "DESAI NIMESH JAYANTIBHAI", "primary", "21-JUN-74", "AGBPD3221Q", first="DESAI",
                middle="NIMESH", last="JAYANTIBHAI", addresses=[("ResidentialAddress", desai_home)]),
         person(90322, "DESAI HEMLATABEN NIMESHBHAI", "Co-Applicant", "09-JAN-78", "BHCPD3222R", first="DESAI",
                middle="HEMLATABEN", last="NIMESHBHAI", addresses=[("ResidentialAddress", desai_home)])],
        [collateral("COLL00090044", "CMS09090044", "RESIDENTIAL",
                    "Bungalow No. B-12, Shreeji Bungalows, Survey No. 118/2, Near Thaltej Cross Roads, Thaltej, "
                    "Ahmedabad, Gujarat- 380054, India", "380054",
                    [("Nimeshbhai Desai", 50), ("Hemlataben Nimeshbhai Desai", 50)], built=2400, carpet=2100,
                    pdf=report("B", "Nimesh Jayantibhai Desai", "Mr. Nimeshbhai J. Desai and Mrs. Hemlataben N. Desai",
                               "Bungalow No. B-12 admeasuring 195.10 sq.mtr built up area in the scheme known as "
                               "Shreeji Bungalows constructed on the land bearing Survey No. 118/2 situated at "
                               "Village Thaltej, Taluka Ghatlodia, District Ahmedabad")),
         collateral("COLL00090045", "CMS09090045", "COMMERCIAL",
                    "Office No. 407, 4th Floor, Sterling Business Centre, Near Rajpath Club, Bodakdev, Ahmedabad, "
                    "Gujarat- 380054, India", "380054", [("Nimeshbhai Desai", 100)], built=650, carpet=540,
                    purpose="Rented")],
        [loan(903201, "A", "03-SEP-26", 12500000)],
        case="Husband and wife are both borrowers, and CKYC writes their names surname first ('DESAI NIMESH "
             "JAYANTIBHAI', 'DESAI HEMLATABEN NIMESHBHAI'). The bungalow is owned by 'Nimeshbhai Desai' and "
             "'Hemlataben Nimeshbhai Desai': the wife's name contains her husband's given name, so each owner must "
             "be linked to the right spouse (each borrower is linked to one owner of a property at most). The "
             "office is the husband's alone: IS_ASSET_OWNER 1 for him and 0 for her on that SI."))

    # D33 - only the husband is a borrower, the wife co-owns: third-party mortgage; and a second
    # collateral with exactly the same address (different area and Global collateral No).
    kore_plot = ("Agriculture Plot No. 7, Gat No. 233 and 233/1, Kore Vasti, Hotgi Road, Taluka North Solapur, "
                 "Solapur, Maharashtra- 413002, India")
    apps.append(application(
        "APPL00090033", "Pune", "17-08-2026",
        [person(90331, "MAHADEV SHIVAPPA KORE", "primary", "12-APR-69", "ANKPK3331S",
                addresses=[("ResidentialAddress", "House No. 45, Kore Vasti, Hotgi Road, Solapur, Maharashtra- "
                                                  "413002, India")]),
         firm(90332, "KORE AGRO SERVICES", "Sole Proprietorship Firm", "Co-Applicant", "15-JUN-12", "ANKPK3331S",
              addresses=[("OfficeAddress", "Shop No. 2, Market Yard, Solapur, Maharashtra- 413002, India")])],
        [collateral("COLL00090046", "CMS09090046", "PLOT", kore_plot, "413002",
                    [("Mahadev Shivappa Kore", 50), ("Savita Mahadev Kore", 50)], carpet=16000,
                    nature="GRAM PANCHAYAT", purpose="Vacant"),
         collateral("COLL00090047", "CMS09090047", "PLOT", kore_plot, "413002", [("Mahadev Shivappa Kore", 100)],
                    carpet=1500, nature="GRAM PANCHAYAT", purpose="Vacant")],
        [loan(903301, "A", "17-AUG-26", 3500000)],
        case="Only the husband is a borrower. The first plot is co-owned by his wife 'Savita Mahadev Kore', whose "
             "name contains his given name and surname: she is not him, so that SI is a third-party mortgage. The "
             "second collateral has exactly the same address as the first (different area and Global collateral "
             "No, as in one real application): both are flagged to check that the same property is not recorded "
             "twice."))

    # D34 - an Owner Name cell holding two people; CKYC names with WO / DO (a relative's name follows).
    iyer_home = ("Flat No. 3B, Third Floor, Lakshmi Apartments, Door No. 12, Kamarajar Salai, Mylapore, Chennai, "
                 "Tamil Nadu- 600004, India")
    apps.append(application(
        "APPL00090034", "Chennai", "05-09-2026",
        [person(90341, "SURESH IYER", "primary", "14-FEB-71", "AKZPI3441T",
                addresses=[("ResidentialAddress", iyer_home)]),
         person(90342, "KALPANA WO SURESH IYER", "Co-Applicant", "02-MAR-75", "BMZPI3442U", first="KALPANA",
                middle="WO", last="SURESH IYER", addresses=[("ResidentialAddress", iyer_home)]),
         person(90343, "MEENA DO SURESH IYER", "Co-Applicant", "19-JUL-98", "CNZPI3443V", first="MEENA",
                middle="DO", last="SURESH IYER", addresses=[("ResidentialAddress", iyer_home)])],
        [collateral("COLL00090048", "CMS09090048", "RESIDENTIAL", iyer_home, "600004",
                    [("Suresh Iyer and Kalpana Iyer", 100)], built=1100, carpet=950,
                    pdf=report("A", "Suresh Iyer", "Mr. Suresh Iyer and Mrs. Kalpana Iyer",
                               "Flat No. 3B on the Third Floor admeasuring 1100 sq.ft. built up area in the building "
                               "known as Lakshmi Apartments, Door No. 12, Kamarajar Salai, Mylapore, constructed on "
                               "the land bearing Survey No. 1654/2 of Mylapore Village, Mylapore-Triplicane Taluk, "
                               "Chennai District"))],
        [loan(903401, "A", "05-SEP-26", 7200000)],
        case="COLLATERAL writes both owners in one Owner Name cell ('Suresh Iyer and Kalpana Iyer', 100%, Single), "
             "as in three real collaterals: each part is a borrower, so the cell is taken as two owners (a company "
             "name with 'and', such as 'Kaveri Pumps and Valves', stays whole). CKYC writes the wife and the "
             "daughter as 'KALPANA WO SURESH IYER' and 'MEENA DO SURESH IYER' (as in the real CKYC): the words "
             "after WO / DO name the husband or father and are ignored when matching, so Kalpana is linked to "
             "the owner 'Kalpana Iyer' (listed for review) and the daughter is not an owner."))

    # D35 - the owner's name is part of the borrower's name: her father or husband, not her.
    rs_puram = "Door No. 14/22, Ground Floor, Venkatasamy Road West, R.S. Puram, Coimbatore, Tamil Nadu- 641002, India"
    apps.append(application(
        "APPL00090035", "Coimbatore", "08-09-2026",
        [person(90351, "SANGEETHA RAMALINGAM", "primary", "11-NOV-84", "BQZPR3551W",
                addresses=[("ResidentialAddress", rs_puram)]),
         firm(90352, "SRI SANGEETHA TEXTILES", "Sole Proprietorship Firm", "Co-Applicant", "01-APR-15", "BQZPR3551W",
              addresses=[("OfficeAddress", "No. 45, Oppanakara Street, Town Hall, Coimbatore, Tamil Nadu- 641001, "
                                           "India")])],
        [collateral("COLL00090049", "CMS09090049", "RESIDENTIAL", rs_puram, "641002", [("Ramalingam", 100)],
                    built=1800, carpet=1500)],
        [loan(903501, "A", "08-SEP-26", 4800000)],
        case="The only owner, 'Ramalingam', is part of the borrower's name 'SANGEETHA RAMALINGAM' (her father's or "
             "husband's name, a pattern found in three real collaterals): the given names differ, so he is taken to "
             "be a different person and the SI is a third-party mortgage (left out). The decision is listed in the "
             "Check and correct sheet, where Y would link him to her."))
    return apps


# --------------------------------------------------------------------------
# What a person types in the required-details file for the hand-made cases
# --------------------------------------------------------------------------
# These are written to the answers file (ANSWERS_FILE) with the same values
# for the generated records, and dev_tools/fill_dummy_required_details.py
# types them in, so that the dummy data gives a complete .DAT file.
#
# Every individual: (gender, father's or mother's name). CKYC never has them.
HAND_MADE_PEOPLE = {
    90012: ("M", "KRISHNA PATIL"), 90013: ("M", "DATTATRAY SALUNKHE"), 90021: ("M", "GANPAT KADAM"),
    90022: ("F", "VISHNU SHINDE"), 90032: ("M", "NARAYANAPPA"), 90033: ("F", "N LOKESH"),
    90041: ("M", "MOHD SALIM"), 90042: ("F", "ABDUL RAHIM"), 90051: ("M", "SHIV PRASAD SHARMA"),
    90052: ("F", "RAJENDRA PRASAD SHARMA"), 90053: ("M", "GOPAL CHAND SHARMA"), 90062: ("M", "RAMANLAL PATEL"),
    90072: ("M", "SURESH CHANDRA VERMA"), 90073: ("F", "RAMESH SAXENA"), 90081: ("M", "VINOD MALHOTRA"),
    90082: ("F", "ASHOK KAPOOR"), 90092: ("M", "RAMESH CHANDRA GUPTA"), 90102: ("M", "MOHAN LAL AGARWAL"),
    90112: ("M", "RAMASAMY"), 90113: ("F", "SUBRAMANIAN"), 90121: ("M", "PRAKASH CHANDRA JOSHI"),
    90122: ("F", "DINESH PANDEY"), 90123: ("M", "MAHESH CHANDRA JOSHI"), 90124: ("F", "MAHESH CHANDRA JOSHI"),
    90125: ("M", "PRAKASH CHANDRA JOSHI"), 90131: ("M", "MADHUKAR DESHMUKH"), 90132: ("F", "ARVIND KULKARNI"),
    90141: ("M", "NARENDRA MEHTA"), 90181: ("M", "PANDURANG MORE"), 90182: ("F", "DATTU PAWAR"),
    90191: ("M", "BHANWAR LAL CHOUDHARY"), 90192: ("F", "GOPI RAM"), 90201: ("M", "GOPALAKRISHNAN NAIR"),
    90202: ("F", "SURESH MENON"), 90211: ("F", "PALLE VENKATESWARA RAO"), 90212: ("M", "VELPURI SATYANARAYANA"),
    90221: ("M", "PRAVIN SHAH"), 90222: ("F", "JAYANTILAL DOSHI"), 90231: ("M", "SURESH CHAND JAIN"),
    90232: ("F", "ANIL KHANDELWAL"), 90233: ("M", "MOHAN LAL"), 90242: ("M", "SALIM DEWAN"),
    90251: ("M", "SHIVAJI JADHAV"), 90252: ("F", "VILAS PATIL"), 90261: ("M", "MANSUKH PARMAR"),
    90262: ("F", "KANJI SOLANKI"), 90281: ("M", "HARI CHAND"), 90291: ("M", "YUSUF QURESHI"),
    90292: ("F", "ABDUL KARIM SHAIKH"), 90293: ("M", "IBRAHIM QURESHI"), 90301: ("M", "RAMCHANDRA GANPAT SAWANT"),
    90302: ("F", "SURESH NAIK"), 90304: ("M", "GANPAT SAWANT"), 90311: ("M", "BHIKAJI KULKARNI"),
    90312: ("F", "VASANT JOSHI"), 90321: ("M", "JAYANTIBHAI DESAI"), 90322: ("F", "RAMESHBHAI SHAH"),
    90331: ("M", "SHIVAPPA KORE"), 90341: ("M", "VENKATARAMAN IYER"), 90342: ("F", "SRINIVASAN"),
    90343: ("F", "SURESH IYER"), 90351: ("F", "RAMALINGAM"),
}
# Other values the hand-made cases leave out on purpose: (Neo CIF number) -> {field: value}
BORROWER_ANSWERS = {
    # proprietorships with no PAN in CKYC use their proprietor's PAN (it is inside their GSTIN)
    90031: {"PAN": "AGHPL5521K"},                                               # its proprietor N LOKESH
    90282: {"PAN": "AFYPC7878G", "CITY": "2532", "DISTRICT": "158"},           # Ballabgarh, Faridabad district
    90042: {"CKYC": "50000000090042"},                                          # no PAN: a (made-up) KYC number
    90061: {"REG_DATE": "2012-04-01"},
    90111: {"REG_NUMBER": "AAC4521"},                                           # LLPIN
    90241: {"REG_DATE": "2010-05-12", "REG_NUMBER": "U22209MH2010PTC204512"},   # CIN
    90251: {"DATE_OF_BIRTH": "1985-06-05"},          # CKYC shows it as the incorporation date
    90192: {"ADDRESS": "P.No. B-17, Bhagat Colony, Mahamandir", "PIN_CODE": "342001", "CITY": "5804",
            "DISTRICT": "488"},                       # she lives with the primary borrower, in Jodhpur
    90281: {"CITY": "2532", "DISTRICT": "158"},
}
# (Collateral ID) -> {field: value}
COLLATERAL_ANSWERS = {
    "COLL00090003": {"ASSET_SUB_TYPE_ID": "1", "SURVEY_NUMBER": "422/309", "PROJECT_NAME": "Shivaganga Layout"},
    "COLL00090004": {"PROJECT_NAME": "Kushaiguda"},
    "COLL00090005": {"PROJECT_NAME": "Shri Krishna Nagar Scheme"},
    "COLL00090006": {"ASSET_TYPE_ID": "3"},                     # a plot with no title report: Other
    "COLL00090009": {"ASSET_SUB_TYPE_ID": "7"},
    "COLL00090011": {"ASSET_SUB_TYPE_ID": "4", "SURVEY_NUMBER": "1542/3", "PROJECT_NAME": "Anna Nagar East"},
    "COLL00090014": {"PROJECT_NAME": "Kingsway Grove"},
    "COLL00090016": {"SURVEY_NUMBER": "520/1", "PROJECT_NAME": "Pinnacle Zone",
                     "TYPE_OF_CHARGE": "FIRST CHARGE"},     # the first loan on the property
    "COLL00090017": {"SURVEY_NUMBER": "520/1", "PROJECT_NAME": "Pinnacle Zone",
                     "TYPE_OF_CHARGE": "SECOND CHARGE"},     # the second loan on the same property
    "COLL00090022": {"ASSET_SUB_TYPE_ID": "1", "HOUSE_NUMBER": "B-17", "PROJECT_NAME": "Bhagat Colony",
                     "OWNERS": "Om Prakash Choudhary"},
    "COLL00090025": {"SURVEY_NUMBER": "512/2"},
    "COLL00090026": {"SURVEY_NUMBER": "233/4", "PROJECT_NAME": "Scheme No 78"},
    "COLL00090028": {"PROJECT_NAME": "Shree Colony"},
    "COLL00090029": {"SURVEY_NUMBER": "412/1", "LOCALITY": "Kalawad Road"},
    "COLL00090030": {"ASSET_SUB_TYPE_ID": "4", "SURVEY_NUMBER": "45/12", "PROJECT_NAME": "Rohini Sector 8"},
    "COLL00090031": {"SURVEY_NUMBER": "77/2", "PROJECT_NAME": "Adarsh Nagar", "CITY": "2532",
                     "DISTRICT": "158"},
    "COLL00090039": {"SURVEY_NUMBER": "67/3"},
    "COLL00090041": {"PROJECT_NAME": "Shahapur Village", "ASSET_TYPE_ID": "3"},
    "COLL00090044": {"PROJECT_NAME": "Shreeji Bungalows"},
    "COLL00090045": {"SURVEY_NUMBER": "284/1"},
    "COLL00090046": {"PROJECT_NAME": "Kore Vasti", "ASSET_TYPE_ID": "3"},      # agricultural plots: Other
    "COLL00090047": {"PROJECT_NAME": "Kore Vasti", "ASSET_TYPE_ID": "3"},
    "COLL00090049": {"ASSET_SUB_TYPE_ID": "1", "SURVEY_NUMBER": "8/112", "PROJECT_NAME": "R S Puram"},
}
# The type of mortgage (SI_TYPE_ID) where the title report is missing, scanned,
# misnamed or about another property, so the tool asks for it: all equitable
TYPED_SI_TYPES = {coll: "17" for coll in (
    "COLL00090003", "COLL00090006", "COLL00090009", "COLL00090014", "COLL00090016", "COLL00090017",
    "COLL00090022", "COLL00090024", "COLL00090025", "COLL00090026", "COLL00090028", "COLL00090029",
    "COLL00090030", "COLL00090031", "COLL00090036", "COLL00090039", "COLL00090041", "COLL00090042",
    "COLL00090043", "COLL00090045", "COLL00090046", "COLL00090047", "COLL00090049")}
# Answers in the Check and correct sheet: (Collateral ID, row, field, value, why)
CHECK_ANSWERS = [
    ("COLL00090025", "SI", "PDF_CHECK", "WRONG PDF", "the report is about another property"),
    ("COLL00090042", "SI", "PDF_CHECK", "WRONG PDF", "the two reports were saved under each other's names"),
    ("COLL00090043", "SI", "PDF_CHECK", "WRONG PDF", "the two reports were saved under each other's names"),
    ("COLL00090027", "SI", "TOTAL_SECURED_AMOUNT", "10500000", "the gala secures both active loans"),
]


# Title reports saved under a wrong or odd file name
EXTRA_PDFS = [
    ("APPL00090027_COLL0009030.pdf", "APPL00090027", "COLL00090030",
     "misnamed: the collateral ID has 7 digits (reported; the collateral is treated as having no title report)"),
    ("Title report APPL00090021.pdf", "APPL00090021", "COLL00090024",
     "a name that does not follow <APPL>_<COLL>.pdf (reported)"),
]


# --------------------------------------------------------------------------
# Regular records: generated to give the dummy data its size
# --------------------------------------------------------------------------
# The applications above are hand-made: each one tests a kind of data found in
# the real extracts. The ones below are ordinary records, generated until the
# COLLATERAL extract has exactly COLLATERALS collaterals (distinct Collateral
# IDs, the hand-made ones included). 304 gives a .DAT file of 300 SIs: the 4
# third-party mortgages among the hand-made cases are always left out. Each one
# also gets its CKYC rows and, for every collateral, a title report that
# agrees with the LMS. The proportions follow the real extracts: mostly 3 or 4
# customers per application; one CKYC row per customer, plus one per extra ID
# document (a firm's GST, Udyam or CIN number); one collateral per application
# in nine cases out of ten, otherwise two or three; one or two owners, all of
# them borrowers. A fixed seed makes every run produce the same records.
COLLATERALS = 304
SEED = 20261008

# Real places: the PIN is in the CERSAI masters with a single district, and the
# tool finds the city in the address. (branch, locality, city, state, PIN,
# district, the land-number label used there - "Plot No." where land is
# allotted in plots -, streets, leasehold land)
PLACES = [
    ("Mumbai 1", "Andheri East", "Mumbai", "Maharashtra", "400069", "Mumbai Suburban", "CTS No.",
     ("Sahar Road", "Chakala Road"), False),
    ("Mumbai 2", "Malad West", "Mumbai", "Maharashtra", "400064", "Mumbai Suburban", "CTS No.",
     ("Marve Road", "Liberty Garden Road"), False),
    ("Mumbai 2", "Borivali West", "Mumbai", "Maharashtra", "400092", "Mumbai Suburban", "CTS No.",
     ("S. V. Road", "L. T. Road"), False),
    ("Mumbai 1", "Chembur", "Mumbai", "Maharashtra", "400071", "Mumbai Suburban", "CTS No.",
     ("R. C. Marg", "Central Avenue Road"), False),
    ("Mumbai 1", "Mulund West", "Mumbai", "Maharashtra", "400080", "Mumbai Suburban", "CTS No.",
     ("L. B. S. Marg", "Zaver Road"), False),
    ("Thane", "Majiwada", "Thane", "Maharashtra", "400601", "Thane", "Survey No.",
     ("Majiwada Road", "Kapurbawdi Road"), False),
    ("Thane", "Dombivli East", "Thane", "Maharashtra", "421201", "Thane", "Survey No.",
     ("Manpada Road", "Tilak Road"), False),
    ("VASHI BRANCH", "Kharghar", "Navi Mumbai", "Maharashtra", "410210", "Raigad", "Plot No.",
     ("Sector 12 Road", "Utsav Chowk Road"), True),
    ("Pune", "Kothrud", "Pune", "Maharashtra", "411038", "Pune", "Survey No.", ("Paud Road", "Karve Road"), False),
    ("Pune", "Hadapsar", "Pune", "Maharashtra", "411028", "Pune", "Survey No.",
     ("Solapur Road", "Magarpatta Road"), False),
    ("Pune", "Wakad", "Pune", "Maharashtra", "411057", "Pune", "Survey No.",
     ("Datta Mandir Road", "Kaspate Vasti Road"), False),
    ("Nashik", "Gangapur Road", "Nashik", "Maharashtra", "422013", "Nashik", "Survey No.",
     ("College Road", "Anandvalli Road"), False),
    ("Nagpur", "Dharampeth", "Nagpur", "Maharashtra", "440010", "Nagpur", "Khasra No.",
     ("West High Court Road", "Ram Nagar Road"), False),
    ("Bangalore", "Jayanagar", "Bangalore", "Karnataka", "560041", "Bengaluru Urban", "Sy. No.",
     ("9th Main Road", "30th Cross Road"), False),
    ("Bangalore 2", "Whitefield", "Bangalore", "Karnataka", "560066", "Bengaluru Urban", "Sy. No.",
     ("ITPL Main Road", "Varthur Road"), False),
    ("Bangalore", "Rajajinagar", "Bangalore", "Karnataka", "560010", "Bengaluru Urban", "Sy. No.",
     ("Dr Rajkumar Road", "West of Chord Road"), False),
    ("Mysore", "Vijayanagar", "Mysore", "Karnataka", "570017", "Mysuru", "Sy. No.",
     ("Hunsur Road", "Vijayanagar Main Road"), False),
    ("Hyderabad", "Kukatpally", "Hyderabad", "Telangana", "500072", "Medchal-Malkajgiri", "Survey No.",
     ("JNTU Road", "Kukatpally Main Road"), False),
    ("Hyderabad", "Miyapur", "Hyderabad", "Telangana", "500049", "Medchal-Malkajgiri", "Survey No.",
     ("Allwyn Colony Road", "Miyapur Main Road"), False),
    ("Hyderabad", "Dilsukhnagar", "Hyderabad", "Telangana", "500060", "Ranga Reddy", "Survey No.",
     ("Chaitanyapuri Road", "Kothapet Road"), False),
    ("Chennai", "Anna Nagar", "Chennai", "Tamil Nadu", "600040", "Chennai", "S. No.",
     ("2nd Avenue", "Shanthi Colony Main Road"), False),
    ("Chennai", "Velachery", "Chennai", "Tamil Nadu", "600042", "Chennai", "S. No.",
     ("Velachery Main Road", "Taramani Link Road"), False),
    ("Coimbatore", "Saibaba Colony", "Coimbatore", "Tamil Nadu", "641011", "Coimbatore", "S. No.",
     ("NSR Road", "Bharathi Park Road"), False),
    ("AHMEDABAD", "Navrangpura", "Ahmedabad", "Gujarat", "380009", "Ahmedabad", "Survey No.",
     ("C. G. Road", "University Road"), False),
    ("AHMEDABAD", "Bopal", "Ahmedabad", "Gujarat", "380058", "Ahmedabad", "Survey No.",
     ("Bopal Ghuma Road", "South Bopal Road"), False),
    ("Surat", "Adajan", "Surat", "Gujarat", "395009", "Surat", "R.S. No.",
     ("Adajan Patiya Road", "Honey Park Road"), False),
    ("Vadodara", "Gotri", "Vadodara", "Gujarat", "390021", "Vadodara", "Survey No.",
     ("Gotri Road", "Sevasi Road"), False),
    ("Jaipur", "Malviya Nagar", "Jaipur", "Rajasthan", "302017", "Jaipur", "Khasra No.",
     ("Malviya Nagar Main Road", "Jawahar Lal Nehru Marg"), False),
    ("Jaipur", "Vaishali Nagar", "Jaipur", "Rajasthan", "302021", "Jaipur", "Khasra No.",
     ("Queens Road", "Amrapali Marg"), False),
    ("Delhi 1", "Rohini", "North West Delhi", "Delhi", "110085", "North West Delhi", "Khasra No.",
     ("Sector 7 Road", "Rohini Road"), False),
    ("Delhi 2", "Dwarka", "South West Delhi", "Delhi", "110075", "South West Delhi", "Khasra No.",
     ("Sector 6 Road", "Dwarka Road"), False),
    ("Delhi 2", "Janakpuri", "West Delhi", "Delhi", "110058", "West Delhi", "Khasra No.",
     ("Pankha Road", "Janakpuri Road"), False),
    ("Gurugram", "Sector 56", "Gurgaon", "Haryana", "122011", "Gurgaon", "Khasra No.",
     ("Golf Course Extension Road", "Sector 56 Road"), False),
    ("Noida", "Sector 62", "Noida", "Uttar Pradesh", "201309", "Gautam Buddha Nagar", "Plot No.",
     ("Sector 62 Road", "NH 24 Service Road"), True),
    ("Lucknow", "Gomti Nagar", "Lucknow", "Uttar Pradesh", "226010", "Lucknow", "Khasra No.",
     ("Vibhuti Khand Road", "Patrakarpuram Road"), False),
    ("Indore", "Vijay Nagar", "Indore", "Madhya Pradesh", "452010", "Indore", "Khasra No.",
     ("A. B. Road", "Ring Road"), False),
    ("Bhopal", "Arera Colony", "Bhopal", "Madhya Pradesh", "462016", "Bhopal", "Khasra No.",
     ("E-5 Road", "Arera Hills Road"), False),
    ("Kolkata", "Salt Lake", "Kolkata", "West Bengal", "700091", "North 24 Parganas", "Plot No.",
     ("Sector V Road", "Broadway Road"), True),
    ("Kochi", "Kakkanad", "Kochi", "Kerala", "682030", "Ernakulam", "Sy. No.",
     ("Seaport Airport Road", "Infopark Road"), False),
    ("Ludhiana", "Model Town", "Ludhiana", "Punjab", "141002", "Ludhiana", "Khasra No.",
     ("Model Town Road", "Pakhowal Road"), False),
]
# state -> (GST state code, Udyam state code, CIN state code), and how its
# families write their names (see Family.full_name)
STATE_CODES = {"Maharashtra": ("27", "MH", "MH"), "Karnataka": ("29", "KR", "KA"), "Telangana": ("36", "TS", "TG"),
               "Tamil Nadu": ("33", "TN", "TN"), "Gujarat": ("24", "GJ", "GJ"), "Rajasthan": ("08", "RJ", "RJ"),
               "Delhi": ("07", "DL", "DL"), "Haryana": ("06", "HR", "HR"), "Uttar Pradesh": ("09", "UP", "UP"),
               "Madhya Pradesh": ("23", "MP", "MP"), "West Bengal": ("19", "WB", "WB"),
               "Kerala": ("32", "KL", "KL"), "Punjab": ("03", "PB", "PB")}
NAMING = {"Maharashtra": "west", "Gujarat": "gujarat", "Rajasthan": "north", "Delhi": "north", "Haryana": "north",
          "Uttar Pradesh": "north", "Madhya Pradesh": "north", "Punjab": "punjab", "Karnataka": "karnataka",
          "Telangana": "telugu", "Tamil Nadu": "tamil", "Kerala": "kerala", "West Bengal": "bengal"}
# given names of men and women, surnames ("" = none: the initial is used), and
# the names of the older generation (fathers)
NAMES = {
    "west": {"men": ["RAHUL", "SACHIN", "SANDEEP", "PRASHANT", "NITIN", "MAHESH", "SANTOSH", "MANGESH", "TUSHAR",
                     "YOGESH", "SWAPNIL", "AMOL"],
             "women": ["PRIYA", "SNEHA", "SWATI", "MANISHA", "KAVITA", "RUPALI", "SHWETA", "VAISHALI", "MADHURI",
                       "ASHWINI", "PALLAVI", "TEJASWINI"],
             "surnames": ["PATIL", "JADHAV", "PAWAR", "SHINDE", "KULKARNI", "DESHPANDE", "GAIKWAD", "CHAVAN",
                          "SALUNKHE", "KALE", "DESHMUKH", "MANE", "LONDHE", "WAGH"],
             "elders": ["VIJAY", "SURESH", "RAMESH", "DATTATRAY", "GANPAT", "SHANKAR", "VASANT", "PRAKASH", "ASHOK",
                        "SHRIKANT", "DILIP", "ARUN"]},
    "gujarat": {"men": ["KETAN", "JIGNESH", "HITESH", "PARESH", "NILESH", "MEHUL", "CHIRAG", "DHAVAL", "KALPESH",
                        "BHAVIN", "TEJAS", "VIRAL"],
                "women": ["HETAL", "NISHA", "KOMAL", "BHAVNA", "JAGRUTI", "DHARA", "MITAL", "RINKU", "PAYAL",
                          "KINJAL"],
                "surnames": ["PATEL", "SHAH", "MEHTA", "DESAI", "PARIKH", "MODI", "TRIVEDI", "PANCHAL", "SONI",
                             "THAKKAR"],
                "elders": ["RAMESH", "DAHYA", "KANTI", "MANU", "NAVIN", "JAYANTI", "ARVIND", "HASMUKH", "PRAVIN"]},
    "north": {"men": ["ROHIT", "VIKAS", "MANOJ", "SANJAY", "RAKESH", "DEEPAK", "PANKAJ", "ANKIT", "GAURAV", "VIVEK",
                      "NARESH", "ANUJ"],
              "women": ["NEHA", "RITU", "ANITA", "SHALINI", "PREETI", "MEENAKSHI", "SEEMA", "AARTI", "SONIA",
                        "RASHMI"],
              "surnames": ["SHARMA", "VERMA", "GUPTA", "AGARWAL", "YADAV", "JAIN", "SAXENA", "MISHRA", "TIWARI",
                           "RATHORE", "KHANDELWAL", "MALHOTRA", "ARORA", "SRIVASTAVA"],
              "elders": ["RAMESH", "SURESH", "MAHESH", "OM PRAKASH", "RAJENDRA", "SATISH", "KRISHAN", "VINOD"]},
    "punjab": {"men": ["GURPREET", "BALWINDER", "JASPAL", "KULDEEP", "AMRIK", "TEJINDER", "HARVINDER"],
               "women": ["HARLEEN", "NAVNEET", "SIMRAN", "MANDEEP", "PARMINDER"],
               "surnames": ["GILL", "SIDHU", "GREWAL", "DHILLON", "SANDHU", "BAJWA"],
               "elders": ["GURDEV", "SURJIT", "HARBANS", "MOHINDER", "JOGINDER"]},
    "karnataka": {"men": ["MANJUNATH", "RAGHAVENDRA", "NAGARAJ", "GIRISH", "SHIVAKUMAR", "HARISH", "PRAVEEN",
                          "SANTHOSH", "CHANDRASHEKAR"],
                  "women": ["SHILPA", "SOWMYA", "PAVITHRA", "ANITHA", "RASHMI", "DEEPA", "VANITHA", "NAYANA"],
                  "surnames": ["GOWDA", "HEGDE", "SHETTY", "BHAT", "", "", ""],
                  "elders": ["KRISHNAPPA", "NANJUNDAPPA", "RAMAIAH", "SIDDAPPA", "VENKATARAMANA", "BASAVARAJ"]},
    "telugu": {"men": ["SRINIVAS", "VENKATESH", "NAVEEN", "MAHESH", "KIRAN", "RAVI TEJA", "SURESH"],
               "women": ["PADMA", "SWATHI", "SRAVANI", "ANUSHA", "MADHAVI", "SIRISHA", "LAVANYA"],
               "surnames": ["REDDY", "NAIDU", "CHOWDARY", "GOUD", "RAO", "VARMA"],
               "elders": ["NARAYANA", "SUBBA RAO", "RAMA RAO", "VENKATA RAMANA", "SATYANARAYANA"]},
    "tamil": {"men": ["KARTHIK", "SENTHIL", "SARAVANAN", "BALAJI", "VIGNESH", "MURUGAN", "DINESH", "PRABHU",
                      "ASHOK KUMAR"],
              "women": ["PRIYA", "KAVITHA", "REVATHI", "SARANYA", "DIVYA", "MEENA", "GAYATHRI", "NITHYA"],
              "surnames": [""],
              "elders": ["MUTHU", "RAMASAMY", "PERUMAL", "SUBRAMANI", "GOVINDAN", "KANNAN", "PALANISAMY"]},
    "kerala": {"men": ["ANIL", "SAJEEV", "BIJU", "SHAJI", "RAJEEV", "SUNIL", "PRADEEP"],
               "women": ["ANJU", "BINDU", "RESHMA", "LEKHA", "SREEJA", "DIVYA"],
               "surnames": ["NAIR", "MENON", "PILLAI", "KURIAN", "VARGHESE", "THOMAS"],
               "elders": ["GOPINATHAN", "RAMACHANDRAN", "BALAKRISHNAN", "JOSEPH", "MATHEW"]},
    "bengal": {"men": ["SOURAV", "ARINDAM", "DEBASISH", "PARTHA", "SANJIB", "ABHIJIT", "SUBRATA"],
               "women": ["MOUMITA", "SUTAPA", "ANANYA", "RINA", "PAPIYA", "SWAGATA"],
               "surnames": ["BANERJEE", "CHATTERJEE", "MUKHERJEE", "GHOSH", "BOSE", "DUTTA", "SEN"],
               "elders": ["SUBHASH", "DILIP", "PRADIP", "TAPAN", "ASHIM"]},
}
FIRM_PREFIXES = {"west": ["Shree Samarth", "Mangalmurti", "Shree Ganesh", "Jai Malhar", "Siddhivinayak"],
                 "gujarat": ["Shreeji", "Jalaram", "Umiya", "Swastik", "Parshwa"],
                 "north": ["Shree Balaji", "Shiv Shakti", "Bajrang", "Shree Shyam", "Radhe Krishna"],
                 "punjab": ["Guru Kirpa", "Sat Kartar", "Khalsa", "Mehar", "Waheguru"],
                 "karnataka": ["Sri Venkateswara", "Sri Lakshmi", "Annapoorna", "Sri Sai Ram", "Sri Manjunatha"],
                 "telugu": ["Sri Venkateswara", "Sri Lakshmi", "Annapoorna", "Sri Sai Ram", "Sri Durga"],
                 "tamil": ["Sri Murugan", "Sri Lakshmi", "Annapoorna", "Sri Sai Ram", "Sri Balaji"],
                 "kerala": ["Sree Krishna", "Sree Bhadra", "Malabar", "Kairali", "Sree Lakshmi"],
                 "bengal": ["Maa Tara", "Joy Guru", "Loknath", "Ma Durga", "Annapurna"]}
TRADES = ["Traders", "Enterprises", "Hardware", "Electricals", "Medical Stores", "Agencies", "Textiles", "Auto Parts",
          "Furniture", "Packaging", "Plastics", "Engineering Works", "Opticals", "Sweets", "Caterers", "Printers",
          "Tyres", "Paints", "Sanitary Stores", "Steel Traders"]
INDUSTRIES = {"Polymers": "22209", "Infratech": "45200", "Foods": "15400", "Packaging": "17029", "Pharma": "21002",
              "Logistics": "63090", "Textiles": "17111", "Agro Industries": "01403", "Engineering": "29299"}
BUILDING_PREFIXES = ["Sai Krupa", "Om Shanti", "Gokul", "Siddhi", "Sunrise", "Green Valley", "Silver Oak", "Lotus",
                     "Orchid", "Anand", "Vrindavan", "Pushpak", "Nandanvan", "Ashirwad", "Sukh Shanti", "Navratna",
                     "Sapphire", "Emerald", "Mangal", "Shubham"]
FLAT_SUFFIXES = ["Apartments", "Residency", "Heights", "Towers", "Enclave", "Homes"]
HOUSE_SUFFIXES = ["Enclave", "Gardens", "Vihar", "Homes", "Avenue"]
SHOP_SUFFIXES = ["Plaza", "Arcade", "Business Centre", "Trade Centre", "Chambers", "Complex", "Shopping Centre"]
LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
# market rate per square foot of built-up area (residential; commercial premises a third more), by city
RATE_PER_SQFT = {"Mumbai": (18000, 32000), "Thane": (9000, 15000), "Navi Mumbai": (9000, 14000),
                 "Pune": (7000, 12000), "Nashik": (4500, 7000), "Nagpur": (4500, 7500), "Bangalore": (6500, 12000),
                 "Mysore": (4500, 7000), "Hyderabad": (5500, 9500), "Chennai": (6500, 11000),
                 "Coimbatore": (4500, 7500), "Ahmedabad": (4500, 8000), "Surat": (4500, 7500),
                 "Vadodara": (4000, 6500), "Jaipur": (4000, 7000), "North West Delhi": (9000, 15000),
                 "South West Delhi": (9000, 16000), "West Delhi": (10000, 17000), "Gurgaon": (8000, 14000),
                 "Noida": (6000, 9500), "Lucknow": (4000, 6500), "Indore": (3500, 6000), "Bhopal": (3500, 5500),
                 "Kolkata": (5500, 9500), "Kochi": (5000, 8000), "Ludhiana": (4000, 6500)}
# loan eligibility programmes, about as common as in the real extract
PROGRAMS = [("1 Vanilla Income Program", 75), ("3 Low LTV Program", 5), ("15 SENP-Assessed Income Program", 5),
            ("11 Bank Salaried Method", 3), ("9 Debt Consolidation Program", 3), ("8 Banking Program", 3),
            ("4 Liquid Income Program", 3), ("20 Salaried Income Program LAP", 3)]


def mon_date(d):
    """DD-MON-YY, as the LMS writes most dates."""
    return f"{d.day:02d}-{MONTHS[d.month - 1]}-{d.year % 100:02d}"


def ordinal(n):
    if n == 0:
        return "Ground"
    return f"{n}{'th' if 10 <= n % 100 <= 20 else {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')}"


def title(name):
    return " ".join(word.capitalize() for word in name.split())


def similar(a, b):
    """Given names too alike for two members of one family (the tool would
    take them for spellings of one name)."""
    return a[:3] == b[:3] or difflib.SequenceMatcher(None, a, b).ratio() >= 0.6


class Family:
    """The members of one family, named the way their region writes names."""

    def __init__(self, rng, style):
        self.rng, self.style, self.pool = rng, style, NAMES[style]
        self.surname = rng.choice(self.pool["surnames"])
        self.initial = rng.choice("BCGKMNPRSV")                     # a Telugu family-name initial
        self.initial_first = rng.random() < 0.5                    # Tamil: "S KARTHIK" or "KARTHIK S"
        self.surname_first = style == "gujarat" and rng.random() < 0.4   # CKYC "PATEL KETAN RAMESHBHAI"
        self.given_names = []

    def pick(self, gender, elder=False):
        names = self.pool["elders"] if elder else self.pool["men" if gender == "M" else "women"]
        free = [n for n in names if not any(similar(n, used) for used in self.given_names)] \
            or [n for n in names if n not in self.given_names] or names
        name = self.rng.choice(free)
        self.given_names.append(name)
        return name

    def member(self, role, gender, birth_year, given=None, father="", husband=""):
        """father / husband: the given name the member's own name carries after theirs."""
        return {"role": role, "gender": gender, "given": given or self.pick(gender), "father": father,
                "husband": husband, "birth_year": birth_year,
                "kumar": self.style == "north" and gender == "M" and self.rng.random() < 0.25}

    def full_name(self, p):
        """The name given name first, the way the family writes it."""
        s, g, link, man = self.surname, p["given"], p["husband"] or p["father"], p["gender"] == "M"
        if self.style == "west":                                    # RAHUL VIJAY PATIL, PRIYA RAHUL PATIL
            return f"{g} {link} {s}"
        if self.style == "gujarat":                                 # KETAN RAMESHBHAI PATEL
            return f"{g} {link}BHAI {s}"
        if self.style == "north":
            return f"{g} KUMAR {s}" if p["kumar"] else f"{g} {s}"
        if self.style == "punjab":
            return f"{g} {'SINGH' if man else 'KAUR'} {s}"
        if self.style == "karnataka":                               # MANJUNATH GOWDA or MANJUNATH K
            return f"{g} {s}" if s else f"{g} {link[0]}"
        if self.style == "telugu":                                  # K SRINIVAS REDDY, K PADMA
            return f"{self.initial} {g} {s}" if man else f"{self.initial} {g}"
        if self.style == "tamil":                                   # S KARTHIK or KARTHIK S
            return f"{link[0]} {g}" if self.initial_first else f"{g} {link[0]}"
        if self.style == "kerala":                                  # ANIL NAIR, BINDU ANIL
            return f"{g} {s}" if man else f"{g} {link}"
        return f"{g} {s}"                                           # bengal

    def parent_name(self, p, rng):
        """The member's father's name, for FATHER_MOTHER_NAME. A wife's or a
        mother's own father is from another family of the same region."""
        if p["father"]:
            given, surname = p["father"], self.surname
        else:
            given = rng.choice(self.pool["elders"])
            surname = rng.choice([s for s in self.pool["surnames"] if s != self.surname] or self.pool["surnames"])
        if self.style == "punjab":
            return f"{given} SINGH {surname}"
        return f"{given} {surname}".strip()

    def ckyc_name(self, p):
        name = self.full_name(p)
        return f"{self.surname} {name.rsplit(' ', 1)[0]}" if self.surname_first else name

    def owner_cell(self, p):
        """How COLLATERAL's Owner Name writes the member: in capitals, or with a title."""
        name, form = self.full_name(p), self.rng.random()
        if form < 0.4:
            return name
        if form < 0.85:
            mr, mrs = ("Shri", "Smt.") if self.style == "north" else ("Mr.", "Mrs.")
            return f"{mr if p['gender'] == 'M' else mrs} {title(name)}"
        return title(name)

    def report_name(self, p):
        """How a title report writes the member: a title, and in the west the
        middle name as an initial ("Mr. Rahul V. Patil")."""
        words = title(self.full_name(p)).split()
        if self.style in ("west", "gujarat") and len(words) == 3:
            words[1] = words[1][0] + "."
        words = [w + "." if len(w) == 1 else w for w in words]
        honorific = "Mr." if p["gender"] == "M" else "Mrs." if p["husband"] else "Ms."
        return f"{honorific} {' '.join(words)}"

    def short_name(self, p):
        """The loan applicant as a title report names it ("Rahul Patil")."""
        words = title(self.full_name(p)).split()
        return f"{words[0]} {words[-1]}" if self.style in ("west", "gujarat") else " ".join(words)


def new_pan(rng, holder, initial, used):
    """A made-up PAN: 4th letter P (person), F (firm) or C (company), 5th the
    first letter of the surname or name."""
    while True:
        pan = "".join(rng.choice(LETTERS) for _ in range(3)) + holder + initial + f"{rng.randrange(1, 10000):04d}" \
            + rng.choice(LETTERS)
        if pan not in used["pans"]:
            used["pans"].add(pan)
            return pan


def building(rng, suffixes, locality, used):
    """A building name not used before in this locality (two properties of one
    application must not share an address)."""
    for attempt in range(500):
        name = f"{rng.choice(BUILDING_PREFIXES)} {rng.choice(suffixes)}"
        if attempt >= 100:                             # the common names are taken here
            name += f" Phase {attempt // 100 + 1}"
        if (locality, name) not in used["buildings"]:
            used["buildings"].add((locality, name))
            return name
    raise RuntimeError(f"no building name left in {locality}")


def land_number(rng, label):
    if label == "Plot No.":
        return str(rng.randint(1, 160))
    if label == "CTS No.":
        return f"{rng.randint(100, 1999)}{rng.choice(['', '/A', '/B', '/1'])}"
    return f"{rng.randint(10, 450)}/{rng.randint(1, 9)}{rng.choice(['', '', 'A', 'B'])}"


def premises(rng, place, kind, used):
    """A flat, house, office or shop in this place: the parts of its address,
    its areas (LMS square feet), its address as the property's (sometimes with
    the land number) and as a customer's postal address."""
    floor = 0 if kind == "shop" else rng.randint(1, 12) if kind in ("flat", "office") else None
    if kind == "house":
        unit, name = str(rng.randint(1, 250)), building(rng, HOUSE_SUFFIXES, place["locality"], used)
        built = rng.randint(110, 320) * 10
        carpet = round(built * 0.85)
    else:
        unit = str(rng.randint(1, 40)) if kind == "shop" else f"{floor}{rng.randint(1, 8):02d}"
        name = building(rng, FLAT_SUFFIXES if kind == "flat" else SHOP_SUFFIXES, place["locality"], used)
        carpet = rng.randint(*{"flat": (45, 160), "office": (25, 90), "shop": (15, 60)}[kind]) * 10
        built = round(carpet * 1.2)
    word = {"flat": "Flat", "house": "House", "office": "Office", "shop": "Shop"}[kind]
    head = f"{word} No. {unit}" + (f", {ordinal(floor)} Floor" if floor is not None else "")
    wing = rng.choice(["A Wing, ", "B Wing, "]) if kind == "flat" and rng.random() < 0.3 else ""
    land = land_number(rng, place["label"])
    street = rng.choice(place["streets"])
    tail = f"{street}, {place['locality']}, {place['city']}, {place['state']}- {place['pin']}, India"
    with_land = place["label"] == "Plot No." or rng.random() < 0.4      # many LMS addresses leave it out
    return {"kind": kind, "unit": unit, "floor": floor, "name": name, "land": land, "built": built,
            "carpet": carpet, "address": f"{head}, {wing}{name}, " + (f"{place['label']} {land}, " if with_land else "")
            + tail, "postal": f"{head}, {wing}{name}, {tail}"}


def report_description(template, prop, place):
    """Title report item 1 for these premises, worded like the real reports."""
    word = {"flat": "Flat", "house": "House", "office": "Office", "shop": "Shop"}[prop["kind"]]
    land = f"{place['label']} {prop['land']}"
    where = f"{place['locality']}, {place['city']}, District {place['district']}"
    pin = f"{place['state']} {place['pin'][:3]} {place['pin'][3:]}"
    if prop["kind"] == "house":
        area = f"{prop['built'] * 0.092903:.2f}"
        if template == "A":
            return (f"{word} No. {prop['unit']}, adm. {area} Sq. Mtrs., Built up area, consisting of ground and first "
                    f"floors, in the layout known as ‘{prop['name']}’, constructed on land bearing {land}, situated "
                    f"at {where}, {pin}.")
        return (f"{word} No. {prop['unit']} admeasuring {area} sq.mtr built up area consisting of ground and first "
                f"floors in the layout known as {prop['name']} constructed on the land bearing {land} situated at "
                f"{where}")
    area = f"{prop['carpet'] * 0.092903:.2f}"
    floor = ordinal(prop["floor"])
    if template == "A":
        return (f"{word} No. {prop['unit']} on the {floor} Floor, adm. {area} Sq. Mtrs., Carpet area, in the "
                f"building known as ‘{prop['name']}’ constructed on land bearing {land}, situated at {where}, {pin}.")
    return (f"{word} No. {prop['unit']} admeasuring {area} sq.mtr carpet area on the {floor} floor in the building "
            f"known as {prop['name']} of the project known as {prop['name']} constructed on the land bearing {land} "
            f"situated at {where}")


def collateral_count(apps):
    return len({c["coll"] for a in apps for c in a["collaterals"]})


def regular_applications(apps, target):
    """Generated regular applications that bring the COLLATERAL extract to
    `target` collaterals."""
    count = collateral_count(apps)
    if count > target:
        raise SystemExit(f"The hand-made cases already have {count} collaterals: COLLATERALS ({target}) cannot be "
                         "lower.")
    rng = random.Random(SEED)
    used = {"pans": {n for a in apps for c in a["customers"] for t, n in c["ids"] if t == "PAN"}
            | {c["pan"] for a in apps for c in a["customers"] if c["pan"]}, "buildings": set(),
            "firms": {c["name"].upper() for a in apps for c in a["customers"] if c["applicant_type"] != "individual"}}
    number = max(int(a["app"][4:]) for a in apps) - 90000 + 1
    coll = max(int(c["coll"][4:]) for a in apps for c in a["collaterals"]) - 90000 + 1
    made = []
    while count < target:
        app = regular_application(rng, number, coll, target - count, used)
        made.append(app)
        count += len(app["collaterals"])
        number += 1
        coll += len(app["collaterals"])
    return made


def regular_application(rng, number, coll, room, used):
    """One generated application with at most `room` collaterals."""
    place = dict(zip(("branch", "locality", "city", "state", "pin", "district", "label", "streets", "leasehold"),
                     rng.choice(PLACES)))
    gst_code, udyam_code, cin_code = STATE_CODES[place["state"]]
    style = NAMING[place["state"]]
    kind = rng.choices(["individual", "proprietorship", "partnership", "company"], [45, 35, 10, 10])[0]
    size = rng.choices([2, 3, 4, 5, 6], [17, 40, 27, 13, 3])[0]       # customers, as in the real CKYC
    disbursal = date(2026, 6, 1) + timedelta(days=rng.randrange(115))
    received = disbursal - timedelta(days=rng.randint(25, 60))        # the application, before any title search
    fam = Family(rng, style)

    # the family: the head, the wife, then other relatives in a random order
    head_age = rng.randint(30, 60)
    head_year = 2026 - head_age
    grandfather = fam.pick("M", elder=True)
    head_man = rng.random() < 0.8
    head = fam.member("head", "M" if head_man else "F", head_year, father=grandfather)
    wife = fam.member("wife", "F", head_year + rng.randint(0, 6), husband=head["given"]) \
        if head_man and rng.random() < 0.85 else None
    others = [fam.member("brother", "M", head_year + rng.choice([-5, -3, 2, 4]), father=grandfather)]
    if head_age <= 50:
        father = fam.member("father", "M", head_year - rng.randint(24, 32), given=grandfather,
                            father=fam.pick("M", elder=True))
        others += [father, fam.member("mother", "F", father["birth_year"] + rng.randint(2, 6), husband=grandfather)]
    if head_age >= 48:
        gender = rng.choice("MF")
        others.append(fam.member("son" if gender == "M" else "daughter", gender,
                                 min(head_year + rng.randint(24, 30), 2004),
                                 father=head["given"] if head_man else fam.pick("M")))
    brother = others[0]
    rng.shuffle(others)

    home = premises(rng, place, "house" if style in ("karnataka", "telugu", "tamil", "north", "punjab")
                    and rng.random() < 0.3 else "flat", used)
    office = premises(rng, place, rng.choice(["office", "shop"]), used) if kind != "individual" else None

    # the customers, the primary first: (customer, family member or None for the business)
    entries = []
    numbers = iter(range((9000 + number) * 10 + 1, (9000 + number) * 10 + 10))     # their Neo CIF numbers

    def add_person(p, ctype, address, family=fam):
        words = family.full_name(p).split()
        if "pan" not in p:                             # a proprietor already has the firm's PAN
            p["pan"] = new_pan(rng, "P", words[-1][0], used)
        ids = [("PAN", p["pan"])]
        if rng.random() < 0.03:                        # a second ID document: a second CKYC row
            ids.append(rng.choice([("Voter's ID", "".join(rng.choice(LETTERS) for _ in range(3))
                                    + f"{rng.randrange(10 ** 7):07d}"),
                                   ("PASSPORT", f"{rng.choice(LETTERS)}{rng.randrange(10 ** 7):07d}")]))
        dob = date(p["birth_year"], rng.randint(1, 12), rng.randint(1, 28))
        email = (f"{p['given'].split()[0].lower()}.{words[-1].lower()}{dob.year % 100:02d}@example.com"
                 if rng.random() < 0.9 and len(words[-1]) > 1 else None)
        risk = rng.choices(["Low Risk", "Medium Risk", "High Risk"], [83, 15, 2])[0]
        n = next(numbers)
        # a separate generator for the parent's name, so the records stay as they were
        parent = family.parent_name(p, random.Random(f"{SEED}:{n}"))
        entries.append((person(n, family.ckyc_name(p), ctype, mon_date(dob), p["pan"], ids=ids, email=email,
                               addresses=[("ResidentialAddress", address)], risk=risk, gender=p["gender"],
                               parent=parent, phone=f"55{n % 100000000:08d}", coll_email=email), p))

    business, role, second = None, "", None
    if kind == "individual":
        add_person(head, "primary", home["postal"])
    else:
        prefix = rng.choice(FIRM_PREFIXES[style])
        if kind == "company":
            industry = rng.choice(list(INDUSTRIES))
            while (name := f"{prefix} {industry} {rng.choice(['Private Limited', 'Pvt. Ltd.', 'Pvt Ltd'])}").upper() \
                    in used["firms"]:
                prefix, industry = rng.choice(FIRM_PREFIXES[style]), rng.choice(list(INDUSTRIES))
            constitution, role = "Private Limited Company", "director"
            inc = date(rng.randint(2004, 2021), rng.randint(1, 12), rng.randint(1, 28))
            pan = new_pan(rng, "C", name[0], used)
            ids = [("PAN", pan), ("CIN No", f"U{INDUSTRIES[industry]}{cin_code}{inc.year}PTC"
                                            f"{rng.randrange(1, 10 ** 6):06d}")]
        else:
            while (name := f"{prefix} {rng.choice(TRADES)}").upper() in used["firms"]:
                prefix = rng.choice(FIRM_PREFIXES[style])
            constitution = "Sole Proprietorship Firm" if kind == "proprietorship" else "Partnership Firm"
            role = "proprietor" if kind == "proprietorship" else "partner"
            inc = date(rng.randint(2004, 2022), rng.randint(1, 12), rng.randint(1, 28))
            if kind == "proprietorship":               # a proprietorship uses its proprietor's PAN
                head["pan"] = pan = new_pan(rng, "P", fam.full_name(head).split()[-1][0], used)
            else:
                pan = new_pan(rng, "F", name[0], used)
            ids = [("PAN", pan)]
        if rng.random() < 0.6:
            ids.append(("UDHYAM_REGISTRATION", f"UDYAM-{udyam_code}-{rng.randint(1, 35):02d}-"
                                               f"{rng.randrange(10 ** 7):07d}"))
        if rng.random() < 0.4:
            ids.append(("GST_NO", f"{gst_code}{pan}1Z{rng.choice(LETTERS + '0123456789')}"))
        used["firms"].add(name.upper())
        name = name.upper() if rng.random() < 0.6 else name
        email = f"{''.join(ch for ch in name.lower() if ch.isalpha())[:18]}@example.com" if rng.random() < 0.5 else None
        n = next(numbers)
        business = firm(n, name, constitution, "primary", mon_date(inc), pan, ids=ids, email=email,
                        addresses=[("OfficeAddress", office["postal"])],
                        risk=rng.choices(["Low Risk", "Medium Risk"], [85, 15])[0], phone=f"56{n % 100000000:08d}",
                        coll_email=email)
        entries.append((business, None))
        add_person(head, "Co-Applicant", home["postal"])
        if kind in ("partnership", "company"):         # a second partner or director
            if kind == "company" and wife is not None:
                second, wife = wife, None
            else:
                second = brother
                others.remove(brother)
            add_person(second, "Co-Applicant", home["postal"])
    if wife is not None and len(entries) < size:
        add_person(wife, "Co-Applicant", home["postal"])
    for other in others:
        if len(entries) >= size:
            break
        add_person(other, "Co-Applicant", home["postal"])
    if rng.random() < 0.06:                            # a friend who stands guarantee: not a borrower
        friend = Family(rng, style)
        add_person(friend.member("guarantor", "M", 2026 - rng.randint(35, 60), father=friend.pick("M", elder=True)),
                   "Guarantor", premises(rng, place, "flat", used)["postal"], friend)

    customers = [c for c, _ in entries]
    family = [p for _, p in entries if p is not None and p["role"] != "guarantor"]
    wife = next((p for p in family if p["role"] == "wife"), None)
    guarantor = any(p is not None and p["role"] == "guarantor" for _, p in entries)

    # the collaterals: the home and/or the business premises, owned by borrowers
    def owners_of(prop):
        if prop["kind"] in ("office", "shop") and kind == "company":
            return [business]
        if prop["kind"] in ("office", "shop") and kind == "partnership":
            return [business] if second is None or rng.random() < 0.4 else [head, second]
        if prop["kind"] in ("flat", "house") and wife is not None and rng.random() < 0.45:
            return [head, wife]
        return [head]

    # one collateral in nine cases out of ten, otherwise two or three: the home
    # and the business premises, then flats let out (never more than `room`)
    props = [home] if kind == "individual" or rng.random() < 0.5 else [office]
    spare = [p for p in (home, office) if p is not None and p is not props[0]]
    for _ in range(min(rng.choices([1, 2, 3], [90, 8, 2])[0], room) - 1):
        props.append(spare.pop() if spare else premises(rng, place, "flat", used))
    applicant = fam.short_name(head) if business is None else \
        f"{title(business['name'])} (through its {role} {fam.short_name(head)})"
    collaterals, notes = [], []
    for k, prop in enumerate(props):
        owners = owners_of(prop)
        cells = [title(o["name"]) if o is business else fam.owner_cell(o) for o in owners]
        on_report = [f"M/s. {title(o['name'])}" if o is business else fam.report_name(o) for o in owners]
        commercial = prop["kind"] in ("office", "shop")
        if prop is not home and not commercial:
            purpose = "Rented"                         # a second flat, let out
        else:
            purpose = rng.choice(["Self Occupied", "Self Occupied", "Rented"]) if commercial else "Self Occupied"
        template = rng.choice("AB")
        initiated = received + timedelta(days=rng.randint(3, (disbursal - received).days - 12))
        # market value: built-up area at the city's rate per square foot
        low, high = RATE_PER_SQFT[place["city"]]
        market = round(prop["built"] * rng.randint(low, high) * (4 / 3 if commercial else 1) / 50_000) * 50_000
        number_id = 90000 + coll + k
        # Global collateral Nos of their own (CMS0919xxxx): the hand-made ones are CMS090900xx
        collaterals.append(collateral(
            f"COLL{number_id:08d}", f"CMS{number_id + ID_OFFSET + 100000:08d}", "COMMERCIAL" if commercial else "RESIDENTIAL",
            prop["address"], place["pin"], list(zip(cells, {1: [100], 2: [50, 50]}[len(owners)])),
            built=prop["built"], carpet=prop["carpet"],
            nature="LEASEHOLD" if place["leasehold"] or rng.random() < 0.05 else "FREEHOLD", purpose=purpose,
            age=rng.randint(1, 25), market=market,
            pdf=report(template, applicant, (" & " if template == "A" else " and ").join(on_report),
                       report_description(template, prop, place),
                       nature="Commercial" if commercial else "Residential.",
                       mortgage="Registered Mortgage" if place["state"] == "Telangana" and rng.random() < 0.5
                       else "Equitable Mortgage",
                       initiated=initiated.strftime("%d/%m/%Y"),
                       issued=(initiated + timedelta(days=rng.randint(2, 6))).strftime("%d/%m/%Y"))))
        owned = " and ".join(title(o["name"]) if o is business else title(fam.full_name(o)) for o in owners)
        notes.append(f"{prop['kind']} owned by {owned}" + (" (50% each)" if len(owners) == 2 else ""))

    who = {"individual": "an individual", "proprietorship": "a sole proprietorship firm with its proprietor",
           "partnership": "a partnership firm with its partner" + ("s" if second else ""),
           "company": "a private limited company with its director" + ("s" if second else "")}[kind]
    relatives = len(family) - 1 - (1 if second else 0)
    if relatives:
        who += f" and {relatives} family member{'s' if relatives > 1 else ''}"
    if guarantor:
        who += ", and a guarantor"
    # a loan against property: 35-65% of the market value of all its collaterals; about a quarter of the
    # loans are only partly disbursed so far, as in the real extract
    total = sum(c["market"] for c in collaterals)
    sanctioned = max(500_000, round(total * rng.uniform(0.35, 0.65) / 50_000) * 50_000)
    disbursed = sanctioned if rng.random() < 0.77 else round(sanctioned * rng.uniform(0.4, 0.9) / 10_000) * 10_000
    return application(
        f"APPL{90000 + number:08d}", place["branch"], disbursal.strftime("%d-%m-%Y"), customers, collaterals,
        [loan(900000 + number * 100 + 1, "A", mon_date(disbursal), sanctioned, disbursed)],
        received=mon_date(received), product=rng.choices([5000980, 5001360, 5001380], [90, 5, 5])[0],
        foir=str(rng.randint(25, 75)),
        program=rng.choices([name for name, _ in PROGRAMS], [weight for _, weight in PROGRAMS])[0],
        case=f"Regular record (generated): {who} in {place['locality']}, {place['city']}; " + "; ".join(notes)
             + ("; the title reports agree" if len(props) > 1 else "; the title report agrees") + " with the LMS.")


# --------------------------------------------------------------------------
# The LMS extracts
# --------------------------------------------------------------------------
def dod(first_disbursal):
    if not first_disbursal:
        return None
    day, month, year = first_disbursal.split("-")
    return f"{day}-{MONTHS[int(month) - 1]}-{year[2:]}"


# hand-made individuals whose CKYC rows have no Gender (the tool then asks for it)
NO_CKYC_GENDER = {90233}


def ckyc_gender(c):
    """CKYC "Gender": Male or Female for individuals; empty for firms and for
    the people in NO_CKYC_GENDER."""
    if c["number"] in NO_CKYC_GENDER:
        return None
    gender = c.get("gender") or HAND_MADE_PEOPLE.get(c["number"], (None, None))[0]
    return {"M": "Male", "F": "Female"}.get(gender)


def ckyc_rows(apps):
    rows = []
    for a in apps:
        account = a["loans"][0]
        for c in a["customers"]:
            for id_type, id_number in (c["ids"] or [(None, None)]):
                for address in (c["ckyc_addresses"] or [None]):
                    rows.append([a["first_disbursal"], a["branch"], dod(a["first_disbursal"]), a["app"],
                                 account["lan"], c["ctype"], c["applicant_type"], c["name"],
                                 c["constitution"], c["cif"], c["dob"], c["inc"], c["first"], c["middle"], c["last"],
                                 id_type, id_number, c["mobile"], c["email"], address, c["risk"],
                                 "Yes" if c["mobile"] else None, a["product"], a["received"], ckyc_gender(c)])
    return rows


def collateral_rows(apps):
    rows = []
    for a in apps:
        for l in a["loans"]:
            for col in a["collaterals"]:
                owners = col["owners"] or [(None, None)]
                status = col["status"] or (None if not col["owners"] else
                                           "Single" if len(col["owners"]) == 1 else "Joint")
                for c in a["customers"]:
                    for owner, share in owners:
                        for address_type, address in (c["addresses"] or [(None, None)]):
                            rows.append([
                                a["app"], l["lan"], l["status"], l["date"], a["branch"], c["name"], c["ctype"],
                                c["cif"], col["purpose"], col["coll"], col["cms"], col["ownership"], status, owner,
                                share, col["ptype"], col["classification"], col["nature"], col["address"], col["pin"],
                                a["foir"], a["program"], col["built"], col["carpet"], col["age"],
                                col["market"] or int(a["loans"][0]["sanctioned"] * 1.8), l["sanctioned"], l["disbursed"],
                                c["phone"], c["coll_email"], c["pan"], address_type, address])
    return rows


def check_ids(apps):
    """Identifiers that must not repeat by accident: Collateral IDs, loan
    accounts, Global collateral Nos (one property, CMS09090099, deliberately
    secures two applications), Neo CIF IDs (one customer per ID) and PANs (only
    a sole proprietorship shares its proprietor's PAN)."""
    def repeated(values):
        seen, twice = set(), set()
        for v in values:
            (twice if v in seen else seen).add(v)
        return twice
    colls = [c for a in apps for c in a["collaterals"]]
    problems = {
        "Collateral ID": repeated(c["coll"] for c in colls),
        "loan account": repeated(l["lan"] for a in apps for l in a["loans"]),
        "Global collateral No": repeated(c["cms"] for c in colls) - {"CMS09090099"},
    }
    if sum(c["cms"] == "CMS09090099" for c in colls) != 2:
        problems["Global collateral No"].add("CMS09090099 (meant for exactly two collaterals)")
    names, holders = {}, {}
    for a in apps:
        for c in a["customers"]:
            names.setdefault(c["cif"], set()).add(c["name"])
            for kind, number in c["ids"]:
                if kind == "PAN":
                    holders.setdefault(number, set()).add((c["name"], c["constitution"]))
    problems["Neo CIF ID"] = {cif for cif, n in names.items() if len(n) > 1}
    problems["PAN"] = {pan for pan, who in holders.items() if len({n for n, _ in who}) > 1
                       and not any(k == "Sole Proprietorship Firm" for _, k in who)}
    problems = {k: v for k, v in problems.items() if v}
    if problems:
        raise SystemExit(f"Repeated identifiers in the dummy data: {problems}")


def write_answers(apps, path):
    """The values a person would type in the required-details file for this
    data: every individual's gender and father's or mother's name (CKYC never
    has them), the values the hand-made cases leave out on purpose, and the
    answers to the questions in the Check and correct sheet. With them,
    dev_tools/fill_dummy_required_details.py makes the dummy data give a
    complete .DAT file."""
    book = openpyxl.Workbook()
    notes = book.active
    notes.title = "Read me"
    for line in ("Made by dev_tools/make_dummy_data.py together with the dummy extracts.",
                 "The values a person would type in the required-details file for the dummy data. After step 1 "
                 "(--dummy), run dev_tools/fill_dummy_required_details.py to type them in, then steps 2 and 3.",
                 "Gender and father's / mother's name are never in CKYC; the other values are the ones the "
                 "hand-made test cases leave out on purpose. All of it is made up."):
        notes.append([line])
    people = book.create_sheet("Borrowers")
    people_keys = ["GENDER", "FATHER_MOTHER_NAME", "DATE_OF_BIRTH", "PAN", "CKYC", "REG_NUMBER", "REG_DATE",
                   "ADDRESS", "PIN_CODE", "CITY", "DISTRICT"]
    people.append(["Application number", "Neo CIF ID", "Borrower name"] + people_keys)
    for a in apps:
        for c in a["customers"]:
            if c["ctype"] == "Guarantor":                  # not a borrower
                continue
            n = c["number"]
            values = dict(BORROWER_ANSWERS.get(n, {}))
            gender, parent = (c["gender"], c["parent"]) if c.get("gender") else HAND_MADE_PEOPLE.get(n, (None, None))
            if gender:
                values.update(GENDER=gender, FATHER_MOTHER_NAME=parent)
            if values:
                people.append([a["app"], c["cif"], c["name"]] + [values.get(k) for k in people_keys])
    places = book.create_sheet("Collaterals")
    answers = {coll: dict(COLLATERAL_ANSWERS.get(coll, {}), **({"SI_TYPE_ID": code} if code else {}))
               for coll in set(COLLATERAL_ANSWERS) | set(TYPED_SI_TYPES) for code in [TYPED_SI_TYPES.get(coll)]}
    place_keys = sorted({k for v in answers.values() for k in v})
    places.append(["Collateral ID", "Application number"] + place_keys)
    for a in apps:
        for col in a["collaterals"]:
            values = answers.get(col["coll"])
            if values:
                places.append([col["coll"], a["app"]] + [values.get(k) for k in place_keys])
    checks = book.create_sheet("Check and correct")
    checks.append(["Collateral ID", "Row", "Field", "Your value", "Why"])
    for row in CHECK_ANSWERS:
        checks.append(list(row))
    for sheet in book.worksheets:
        for cell in sheet[1]:
            cell.font = Font(bold=True)
    book.save(path)
    return people.max_row - 1


def write_extract(path, headers, rows, note):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Export Worksheet"
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    sql = book.create_sheet("SQL")
    sql["A2"] = note
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)


# --------------------------------------------------------------------------
# The title reports
# --------------------------------------------------------------------------
SIZE, LEADING = 10.5, 13.2
# A TrueType serif font gives real dashes and curly quotes in the text layer, as in
# the real reports. Georgia is used because its text extracts cleanly (with Times
# New Roman, spaces come back as non-breaking spaces). Without it, the built-in
# Times is used with plain ASCII punctuation.
FONT_FILES = {
    "regular": ["/System/Library/Fonts/Supplemental/Georgia.ttf", "C:/Windows/Fonts/georgia.ttf"],
    "bold": ["/System/Library/Fonts/Supplemental/Georgia Bold.ttf", "C:/Windows/Fonts/georgiab.ttf"],
}
_FILES = {style: next((f for f in files if Path(f).exists()), None) for style, files in FONT_FILES.items()}
FONT, BOLD = ("dtr", "dtb") if all(_FILES.values()) else ("tiro", "tibo")
FONT_FILE = {FONT: _FILES["regular"], BOLD: _FILES["bold"]} if FONT == "dtr" else {}
MEASURE = {name: pymupdf.Font(fontfile=path) for name, path in FONT_FILE.items()} or \
    {FONT: pymupdf.Font(FONT), BOLD: pymupdf.Font(BOLD)}
ASCII_ONLY = {"\u2013": "-", "\u2014": "-", "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"'}


def plain(text):
    if FONT_FILE:
        return text
    for a, b in ASCII_ONLY.items():
        text = text.replace(a, b)
    return text
PAGE_W, PAGE_H, TOP, BOTTOM = 595.32, 841.92, 70, 780
STYLES = {
    "A": {"columns": [54, 95, 207, 539], "continuation": False,
          "letterhead": ["Demo_______", "Associates, Advocates and Solicitors",
                         "_______________________________www.demo-associates.example_____________________",
                         "Office No. 3, Example Chambers, Sector 50, Nerul,",
                         "Navi Mumbai- 400706. Tel. 022-0000 0000. demo.associates@example.com.",
                         "_________________________________________________________________________"],
          "firm": "Demo Associates", "advocate": "(Adv. A. B. Sample)"},
    "B": {"columns": [54, 95, 201, 527], "continuation": True,
          "letterhead": ["SAMPLE LAW ASSOCIATES", "Legal Advisors",
                         "Office at: 12, Ground floor, Sample House, Sector 29, Vashi, Navi Mumbai- 400703,",
                         "Email- samplelaw@example.com Ph. 022-0000 0001."],
          "firm": "Sample Law Associates", "advocate": "(Adv. C. D. Example)"},
}


def wrap(text, width, font=FONT, size=SIZE):
    lines = []
    for paragraph in text.split("\n"):
        current = ""
        for word in paragraph.split():
            trial = f"{current} {word}".strip()
            if MEASURE[font].text_length(trial, fontsize=size) <= width:
                current = trial
            else:
                if current:
                    lines.append(current)
                current = word
        lines.append(current)
    return lines


class Writer:
    """Lays out one report: header lines, then a bordered 3-column table that
    continues across pages, then closing text."""

    def __init__(self, style):
        self.style = STYLES[style]
        self.doc = pymupdf.open()
        self.page, self.y, self.number = None, TOP, 0
        self.strikes = []                      # (page index, top, bottom, left edge)

    def new_page(self):
        self.page = self.doc.new_page(width=PAGE_W, height=PAGE_H)
        for name, path in FONT_FILE.items():
            self.page.insert_font(fontname=name, fontfile=path)
        self.number += 1
        self.page.insert_text((PAGE_W / 2 - 12, 40), f"...{self.number}...", fontname=FONT, fontsize=SIZE)
        self.y = TOP
        if self.number > 1 and self.style["continuation"]:
            self.line("Continuation Sheet...", font=BOLD)

    def line(self, text, x=54, font=FONT, size=SIZE, strike=False):
        if self.page is None or self.y + LEADING > BOTTOM:
            self.new_page()
        self.page.insert_text((x, self.y + size), plain(text), fontname=font, fontsize=size)
        if strike and text.strip():
            self.strikes.append((self.number - 1, self.y, self.y + LEADING, x))
        self.y += LEADING

    def paragraph(self, text, strike=False, font=FONT):
        for text_line in wrap(text, PAGE_W - 54 - 60, font):
            self.line(text_line, font=font, strike=strike)

    def row(self, number, label, value, strike=False):
        cols = self.style["columns"]
        labels = wrap(label, cols[2] - cols[1] - 8, font=BOLD)
        values = wrap(value, cols[3] - cols[2] - 8)
        total, done, first = max(len(labels), len(values), 1), 0, True
        while done < total:
            room = int((BOTTOM - self.y - 8) // LEADING)
            if room < 2:
                self.new_page()
                continue
            take = min(room, total - done)
            top, height = self.y, take * LEADING + 8
            for x in cols:
                self.page.draw_line((x, top), (x, top + height), width=0.6)
            self.page.draw_line((cols[0], top), (cols[-1], top), width=0.6)
            self.page.draw_line((cols[0], top + height), (cols[-1], top + height), width=0.6)
            if first:
                self.page.insert_text((cols[0] + 4, top + 4 + SIZE), number, fontname=FONT, fontsize=SIZE)
            for k in range(take):
                y = top + 4 + SIZE + k * LEADING
                if done + k < len(labels):
                    self.page.insert_text((cols[1] + 4, y), plain(labels[done + k]), fontname=BOLD,
                                          fontsize=SIZE)
                if done + k < len(values):
                    self.page.insert_text((cols[2] + 4, y), plain(values[done + k]), fontname=FONT,
                                          fontsize=SIZE)
                    if strike and values[done + k].strip():
                        self.strikes.append((self.number - 1, y - SIZE, y + 3, cols[2] + 2))
            self.y = top + height
            done += take
            first = False
            if done < total:
                self.new_page()

    def finish(self, path):
        for page_index, top, bottom, left in self.strikes:
            page = self.doc[page_index]
            words = [w for w in page.get_text("words")
                     if top - 1 <= (w[1] + w[3]) / 2 <= bottom + 1 and w[0] >= left - 1]
            if words:
                x0, x1 = min(w[0] for w in words), max(w[2] for w in words)
                middle = sum((w[1] + w[3]) / 2 for w in words) / len(words)
                page.draw_line((x0, middle), (x1, middle), width=0.7)
        # embed only the letters used: about 85 KB a report instead of 800 KB
        # (the text the tool reads is the same)
        self.doc.subset_fonts()
        self.doc.save(str(path), garbage=4, deflate=True)


DOCUMENTS = ("1. Photocopy of Agreement for Sale executed between the Promoters of the One part and the Purchasers "
             "of the Other part, registered with the Sub-Registrar.\n2. Photocopy of Registration Receipt.\n"
             "3. Photocopy of Share Certificate issued by the Society.\n4. Photocopy of Commencement Certificate "
             "issued by the Municipal Corporation.\n5. Photocopy of Occupation Certificate.\n6. Photocopy of the "
             "latest Property Tax Paid Receipt.\n7. Photocopy of 7/12 extract.")
LAND_HISTORY = ("LAND HISTORY\nFrom the perusal of documents it is observed that the land described above was "
                "originally agricultural land converted to non-agricultural use by the Collector. The Promoter "
                "obtained the Commencement Certificate and constructed the building, and by an Agreement for Sale "
                "the Promoter sold the said premises to the Purchasers.")
FLOW = ("We have perused the documents and observed that the Promoter was the owner of the larger land. By a "
        "registered Agreement for Sale the Promoter agreed to sell and transfer the said premises to the "
        "Purchasers on the terms and conditions mentioned therein. The Promoter registered the project with RERA. "
        "Occupation certificate issued by the competent authority is not shared with us. Lender to verify the "
        "same through technical team.")
TITLE_DEEDS = ("I - Required Document to be collected Prior to disbursal:-\nOriginals:-\n1. Registered Agreement "
               "for Sale along with stamp duty paid receipt and Index II.\n2. Share Certificate.\n3. NOC for Mortgage "
               "issued by the Society in favour of Anand Rathi Global Finance Ltd. [ARGFL].\nCopies:-\n1. Latest "
               "Maintenance Paid Receipts.\n2. Latest Property Tax Paid Receipts.")


def make_report(path, spec):
    if spec["template"] == "scanned":
        text_version = Writer("A")
        write_report_body(text_version, dict(spec, template="A"))
        page = text_version.doc[0]
        picture = page.get_pixmap(dpi=90)
        scan = pymupdf.open()
        scan_page = scan.new_page(width=PAGE_W, height=PAGE_H)
        scan_page.insert_image(scan_page.rect, pixmap=picture)
        scan.save(str(path), garbage=4, deflate=True)
        return
    writer = Writer(spec["template"])
    write_report_body(writer, spec)
    writer.finish(path)


def write_report_body(w, spec):
    b_style = spec["template"] == "B"
    w.new_page()
    for text in w.style["letterhead"]:
        w.line(text, font=BOLD if text == w.style["letterhead"][0] else FONT)
    w.y += 20
    w.line("TITLE REPORT", x=250, font=BOLD, size=12)
    w.line(f"Status: {spec['status']}", font=BOLD)
    w.y += 6
    w.line(f"Date of Initiation – {spec['initiated']}")
    w.line("Ref. No. –ARGFL/SL/2026-27/____" if b_style else "Ref no. –")
    w.line(f"Date of Issuance – {spec['issued']}")
    for text in ("To,", "Anand Rathi Global Finance Ltd. [ARGFL]", "Express Zone, 8th Floor, A Wing,",
                 "Western Express Highway, Goregaon (E)", "Mumbai 400063", ""):
        w.line(text)
    w.line("Please find enclosed herewith title opinion in respect of the following property:")
    w.line(f"Loan applicant – {spec['applicant']}")
    w.y += 6
    w.line(f"Property Owner– {spec['owners']}")
    w.y += 10
    owners12 = re.sub(r"(?<=[A-Za-z]{2})\.$", "", spec["owners_item12"].strip())   # keep an initial's dot
    verb = "were" if any(sep in owners12 for sep in ("&", " and ", ",")) else "was"
    items = [
        ("1.", "Full Description of the property", spec["description"]),
        ("2", "Nature of immovable property", spec["nature"]),
        ("3", "Type of Transaction", spec["transaction"]),
        ("4", "Name of the owner", spec["owners_item4"]),
        ("5", "List of documents/ deed provided to the advocate and perused by him",
         DOCUMENTS + ("" if b_style else "\n" + LAND_HISTORY)),
    ]
    if b_style:
        items.append(("6", "Flow of Title", FLOW))
    items += [
        ("7", "Title Deeds/ documents to be obtained", TITLE_DEEDS),
        ("8", "Nature of Mortgage", spec["mortgage"]),
        ("9", "Search / Certificate of Non- Encumbrance",
         "As per Search Report issued by the Title Investigator for the searches taken online for the last 13 "
         "years, there are no charges registered against the said property." if b_style else
         "We have caused e-search for the above referred property in the Official IGR Website for the period of "
         "21 years and it is observed that there are no charges/mortgage registered against the said property."),
        ("10", "Views on Encumbrance", "Search awaited" if b_style else "NA."),
        ("11", "Whether the documents examined are duly stamped / registered as per the stamp rule.", "YES"),
        ("12", "Certificate of title",
         f"I/We certify that {owners12} {verb} the owners of the said property and has a valid, clear and "
         "marketable title of the immovable property stated above and is mortgage able and enforceable."),
        ("13.", "Any other suggestion / observations or Advise to protect the security interest of the ARGFL",
         "Please make physical verification for ensuring the possession of the property."),
    ]
    for number, label, value in items:
        w.row(number, label, value, strike=b_style and number == "9")
    w.y += 8
    w.line("REQUIREMENTS:- NIL.", font=BOLD)
    w.line("FINAL CERTIFICATE:-", font=BOLD)
    mortgage = " ".join(spec["mortgage"].split())       # the mortgage named in item 8
    w.paragraph("We have gone through the complete chain of documents, which were produced by the ARGFL before me. "
                "All documents are proper and in order. After perusal of all the documents as per Serial No.5, our "
                f"{mortgage} will be valid with regard to said property, which shall be enforceable under the "
                "SARFAESI ACT.", strike=b_style)
    w.line("Thanking you.")
    w.line(f"For {w.style['firm']}", x=380)
    w.line(w.style["advocate"], x=380)


# --------------------------------------------------------------------------
def main():
    apps = build_applications()
    apps += regular_applications(apps, COLLATERALS)
    assert collateral_count(apps) == COLLATERALS, collateral_count(apps)
    check_ids(apps)
    for folder in (CKYC_DIR, COLL_DIR, LEGAL_DIR):
        folder.mkdir(parents=True, exist_ok=True)
        for old in folder.iterdir():
            if old.is_file():
                old.unlink()
    note = "Dummy data made by dev_tools/make_dummy_data.py - not an LMS export."
    write_extract(CKYC_DIR / "CKYC.xlsx", CKYC_HEADERS, ckyc_rows(apps), note)
    write_extract(COLL_DIR / "COLLATERAL.xlsx", COLL_HEADERS, collateral_rows(apps), note)
    answered = write_answers(apps, ANSWERS_FILE)

    specs = {}
    for a in apps:
        for col in a["collaterals"]:
            if col["pdf"]:
                specs[col["coll"]] = (a, col)
                name = col["pdf"]["file_name"] or f"{a['app']}_{col['coll']}.pdf"
                make_report(LEGAL_DIR / name, col["pdf"])
    for name, app, coll, _ in EXTRA_PDFS:
        a, col = next((a, c) for a in apps for c in a["collaterals"] if c["coll"] == coll)
        initiated, issued = report_dates(a)
        spec = col["pdf"] or report("A", a["customers"][0]["name"].title(),
                                    "; ".join(o.strip() for o, _ in col["owners"]) or "-",
                                    col["address"].rsplit(",", 3)[0], initiated=initiated, issued=issued)
        make_report(LEGAL_DIR / name, spec)
    (LEGAL_DIR / "notes.txt").write_text("Not a PDF: the tool ignores this file.\n", encoding="utf-8")

    lines = ["# Dummy data cases", "",
             "Made by `dev_tools/make_dummy_data.py`. All people, companies, buildings, PANs, GSTINs, phone "
             "numbers, e-mails and law firms are made up; PIN codes and place names are real. Phone numbers start "
             "with 55 or 56, so they cannot be real Indian mobile numbers; e-mails use the reserved example.com, "
             ".net and .org domains; Neo CIF IDs and Global collateral Nos are far above the range the LMS uses. "
             "Run the tool on this data with `--dummy`.", "",
             "As in the LMS, an application can have several collaterals (one loan secured by several properties), "
             "but a Collateral ID belongs to one application only. Borrowers are per application; owners are per "
             "collateral and can be the same, different, or overlap between the collaterals of one application.", "",
             "Every loan is active (Loan Status A); there are no cancelled (X) or closed (C) loans.", "",
             "`Dummy required details answers.xlsx` holds what a person would type in the required-details file "
             "(every individual's gender and father's / mother's name, and the values the hand-made cases leave "
             "out); `dev_tools/fill_dummy_required_details.py` types it in after step 1.", ""]
    generated = [a for a in apps if a["case"].startswith("Regular record")]
    if generated:
        lines += [f"Applications {generated[0]['app']} to {generated[-1]['app']} are ordinary records, generated so that "
                  f"COLLATERAL has exactly {COLLATERALS} collaterals (`COLLATERALS` in the script). Each one has its "
                  "CKYC rows and a title report for every collateral that agrees with the LMS.", ""]
    lines += ["| Application | Collateral(s) | What it tests |", "|---|---|---|"]
    for a in apps:
        colls = ", ".join(c["coll"] for c in a["collaterals"])
        lines.append(f"| {a['app']} | {colls} | {a['case']} |")
    lines += ["", "## Title report files with problems", "", "| File | What it tests |", "|---|---|"]
    lines += [f"| {name} | {why} |" for name, _, _, why in EXTRA_PDFS]
    lines += ["| notes.txt | not a PDF: ignored |", ""]
    (OUT / "Dummy data cases.md").write_text("\n".join(lines), encoding="utf-8")

    customers = sum(len(a["customers"]) for a in apps)
    print(f"{len(apps)} applications, {customers} customers, "
          f"{sum(len(a['collaterals']) for a in apps)} collaterals, "
          f"{len(ckyc_rows(apps))} CKYC rows, {len(collateral_rows(apps))} COLLATERAL rows, "
          f"{len(list(LEGAL_DIR.glob('*.pdf')))} PDFs")
    print(f"Answers for the required-details file: {answered} borrower(s) -> {ANSWERS_FILE.name}")
    print(f"Written to {OUT}")


if __name__ == "__main__":
    main()
