"""Creates the dummy input data used while developing and testing the tool.

    python dev_tools/make_dummy_data.py

Writes (replacing earlier dummy files):
    Dummy input data/Dummy CKYC/CKYC.xlsx
    Dummy input data/Dummy COLLATERAL/COLLATERAL.xlsx
    Dummy input data/Dummy Legal/*.pdf
    Dummy input data/Dummy data cases.md    (what each application tests)

Every person, company, PAN, phone number, e-mail and law firm is made up.
PIN codes and place names are real, because the tool looks them up in the
CERSAI masters. The files copy the layout of the real LMS extracts: the same
sheets ("Export Worksheet" and "SQL"), columns, and value formats (DD-MM-YYYY
and DD-MON-YY dates, "+91--" mobiles, numbers stored as numbers). Rows are
built the way the LMS queries build them: CKYC has one row per customer x ID
document x address; COLLATERAL has one row per loan account x customer x
owner x customer address.
"""
from pathlib import Path

import openpyxl
import pymupdf

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Dummy input data"
CKYC_DIR, COLL_DIR, LEGAL_DIR = OUT / "Dummy CKYC", OUT / "Dummy COLLATERAL", OUT / "Dummy Legal"

CKYC_HEADERS = ["First Disbursal Date", "Branch Name", "DOD", "Application Number", "Loan Account No",
                "Customer Type", "Applicant Type", "Customer Name", "CONSTITUTION", "Neo CIF ID", "Date Of Birth",
                "Incorporation Date", "First Name", "Middle Name", "Last Name", "Identification Type",
                "Identification Number", "Mobile No", "Email ID", "Full Address", "Customer Risk Category",
                "Primary Mobile", "Product ID", "Application Received Date"]
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
def cif(n):
    return f"GLBCUST{n:011d}"


def lan(n):
    return f"ARGFLLAPMUM2{n:012d}"


def person(n, name, ctype, dob, pan=None, *, ids=None, first=None, middle=None, last=None, mobile="auto",
           email=None, addresses=(), ckyc_addresses=None, phone=None, coll_email=None,
           applicant_type="individual", constitution="Individual", inc=None, risk="Low Risk"):
    """An individual customer. addresses: [(ADDRESS_TYPE, text)] as in COLLATERAL;
    CKYC gets the same texts unless ckyc_addresses is given."""
    words = name.split()
    return {
        "cif": cif(n), "name": name, "ctype": ctype, "applicant_type": applicant_type,
        "constitution": constitution, "dob": dob if applicant_type == "individual" else None,
        "inc": inc if applicant_type != "individual" else None,
        "first": first if first is not None else words[0],
        "middle": middle if middle is not None else (" ".join(words[1:-1]) or None),
        "last": last if last is not None else (words[-1] if len(words) > 1 else None),
        "ids": ids if ids is not None else ([("PAN", pan)] if pan else []),
        "pan": pan, "mobile": f"+91--98{n % 100000000:08d}" if mobile == "auto" else mobile,
        "email": email, "addresses": list(addresses),
        "ckyc_addresses": ckyc_addresses if ckyc_addresses is not None else [a for _, a in addresses],
        "phone": phone, "coll_email": coll_email, "risk": risk,
    }


def firm(n, name, constitution, ctype, inc, pan=None, *, ids=None, mobile="auto", email=None, addresses=(),
         ckyc_addresses=None, applicant_type="non_individual", dob=None, risk="Low Risk"):
    """A non-individual customer (firm, company, LLP, HUF, trust)."""
    return {
        "cif": cif(n), "name": name, "ctype": ctype, "applicant_type": applicant_type,
        "constitution": constitution, "dob": dob if applicant_type == "individual" else None,
        "inc": inc if applicant_type != "individual" else None,
        "first": None, "middle": None, "last": None,
        "ids": ids if ids is not None else ([("PAN", pan)] if pan else []),
        "pan": pan, "mobile": f"+91--99{n % 100000000:08d}" if mobile == "auto" else mobile,
        "email": email, "addresses": list(addresses),
        "ckyc_addresses": ckyc_addresses if ckyc_addresses is not None else [a for _, a in addresses],
        "phone": None, "coll_email": None, "risk": risk,
    }


def loan(number, status, date, sanctioned, disbursed=None):
    if disbursed is None:
        disbursed = 0 if status == "X" else sanctioned
    return {"lan": lan(number), "status": status, "date": date, "sanctioned": sanctioned, "disbursed": disbursed}


def collateral(coll, cms, ptype, address, pin, owners, *, built=None, carpet=None, classification="Others",
               nature="FREEHOLD", purpose="Self Occupied", ownership="Others", age=8, market=None, pdf=None):
    """owners: [(name as typed in the LMS, percentage share)]; pdf: a title-report spec or None."""
    return {"coll": coll, "cms": cms, "ptype": ptype, "address": address, "pin": pin, "owners": owners,
            "built": built, "carpet": carpet, "classification": classification, "nature": nature,
            "purpose": purpose, "ownership": ownership, "age": age, "market": market, "pdf": pdf}


def application(app, branch, first_disbursal, customers, collaterals, loans, *, received="15-MAY-26",
                product=5000980, case=""):
    return {"app": app, "branch": branch, "first_disbursal": first_disbursal, "customers": customers,
            "collaterals": collaterals, "loans": loans, "received": received, "product": product, "case": case}


def report(template, applicant, owners, description, *, nature="Residential.", transaction="LAP",
           mortgage="Equitable Mortgage", owners_item4=None, owners_item12=None, status=None,
           file_name=None, initiated="18/08/2026", issued="22/08/2026"):
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
    goregaon_office = "Office No. 20, Yashodham Shopping Centre, 1st Floor, Above HDFC Bank, Goregaon East, Mumbai, " \
                      "Maharashtra- 400063, India"
    apps.append(application(
        "APPL00090001", "Mumbai 2", "01-07-2026",
        [firm(90011, "Shree Sai Enterprises", "Partnership Firm", "primary", "01-JUL-04",
              ids=[("PAN", "AAHFS4821K"), ("GST_NO", "27AAHFS4821K1ZP"),
                   ("UDHYAM_REGISTRATION", "UDYAM-MH-18-0012345")],
              email="shreesai.ent@gmail.com", addresses=[("OfficeAddress", goregaon_office)]),
         person(90012, "RAMESH K PATIL", "Co-Applicant", "01-JUN-61", ids=[("PAN", "AFMPP1762L"), ("PAN", "AFMPP1762L")],
                pan="AFMPP1762L", email="ramesh.patil61@gmail.com",
                addresses=[("ResidentialAddress", "112 Raag Tower, Megh Malhar Complex, Nr Royal Challenge Hotel, "
                                                  "Goregaon East, Mumbai, Maharashtra- 400063, India"),
                           ("PermanentAddress", "At Post Kasegaon, Tal Walwa, Sangli, Maharashtra- 415404, India")]),
         person(90013, "Vijay Dattatray Bhosale", "Co-Applicant", "11-JAN-51", "ABLPB4977X",
                email="vd.bhosale@yahoo.com",
                addresses=[("ResidentialAddress", "Flat no A/602, Rukmini Vaibhav CHS Ltd, Veera Desai Road, Azad "
                                                  "Nagar, Road No 3, Andheri West, Mumbai, Maharashtra- 400058, India")])],
        [collateral("COLL00090001", "CMS00090001", "COMMERCIAL",
                    "Office No. 21, 2nd Floor, Yashodham Business CSL, Near Royal Challenge Restaurant, Aba Karmarkar "
                    "Road, Dindoshi, Goregaon (E), Mumbai 400063, CTS No. 156/A/7/1pt, Mumbai, Maharashtra- 400063, "
                    "India", "400063", [(" MR. RAMESH KRISHNA PATIL", 100)], built=620, carpet=520,
                    purpose="Rented", market=12500000,
                    pdf=report("B", "Shree Sai Enterprises (through its partner Ramesh K. Patil)", "Mr. Ramesh K. Patil",
                               "Office No. 21 admeasuring 48.31 sq.mtr carpet area on the 2nd floor in the building "
                               "known as Yashodham Business CSL of the project known as Yashodham Business CSL "
                               "constructed on the land bearing CTS No. 156/A/7/1pt of Village Malad, Taluka Borivali, "
                               "Aba Karmarkar Road, Dindoshi, Goregaon (East), Mumbai 400063",
                               nature="Commercial", initiated="09/06/2026", issued="12/06/2026; 16/06/2026"))],
        [loan(900101, "A", "01-JUL-26", 8800000)],
        case="Golden record: partnership firm + 2 individuals, PAN listed twice, owner verified against a "
             "B-style title report (initials: 'Ramesh K. Patil' = 'RAMESH KRISHNA PATIL'), CTS number in both."))

    # D02 - joint owners with Mr/Mrs and initials; survey number only in the title report.
    thane_home = "Flat No. 1102, 11th Floor, C Wing, Devpriya Heights, Pokharan Road No. 1, Vartak Nagar, Thane West, " \
                 "Thane, Maharashtra- 400606, India"
    apps.append(application(
        "APPL00090002", "Thane", "18-08-2026",
        [person(90021, "SUNIL GANPAT KADAM", "primary", "23-SEP-78", "AKZPK4410M", email="sunil.kadam78@gmail.com",
                addresses=[("ResidentialAddress", thane_home)]),
         person(90022, "ANITA SUNIL KADAM", "Co-Applicant", "05-DEC-82", "BQRPK7788D", first="ANITA", middle="WO",
                last="KADAM", addresses=[("ResidentialAddress", thane_home)])],
        [collateral("COLL00090002", "CMS00090002", "RESIDENTIAL",
                    "Flat No. 1102, 11th Floor, C Wing, Devpriya Heights Co-op. Hsg. Soc. Ltd., Pokharan Road No. 1, "
                    "Vartak Nagar, Thane (West), Thane, Maharashtra- 400606, India", "400606",
                    [(" Mr. Sunil Ganpat Kadam", 50), ("Mrs. Anita Sunil Kadam", 50)], built=780, carpet=655.5,
                    classification="Residential Flat",
                    pdf=report("A", "Sunil Kadam", "Mr. Sunil G. Kadam & Mrs. Anita S. Kadam.",
                               "Flat No. 1102 on the 11th Floor, adm. 60.90 Sq. Mtrs., Carpet area, in the Society "
                               "known as ‘Devpriya Heights Co-op. Hsg. Soc. Ltd.’ constructed on land "
                               "bearing Survey No. 149/1/3, 150/6, 152/5, situated at Village Panchpakhadi, Taluka and "
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
              ids=[("GST_NO", "29ABCPL1234M1Z5"), ("UDHYAM_REGISTRATION", "UDYAM-KR-03-0045678"),
                   ("Shop/Establishment Number", "123456789012")],
              email="srilakshmitraders@yahoo.co.in",
              addresses=[("OfficeAddress", "No. 45, 2nd Cross, APMC Yard, Yeshwanthpur, Bangalore, Karnataka- "
                                           "560022, India")]),
         person(90032, "N LOKESH", "Co-Applicant", "14-FEB-85", "AGHPL5521K", first="N LOKESH", middle=None, last=" ",
                addresses=[("ResidentialAddress", "H N 8/C, 1st Cross, Goraguntepalya, Bangalore, Karnataka- 560022, "
                                                  "India")]),
         person(90033, "Kavya L", "Co-Applicant", "30-AUG-90", "CQWPK3321A", email="kavya-l@outlook.com",
                addresses=[("ResidentialAddress", "H N 8/C, 1st Cross, Goraguntepalya, Bangalore, Karnataka- 560022, "
                                                  "India")])],
        [collateral("COLL00090003", "CMS00090003", "RESIDENTIAL",
                    "Property bearing No. 126, BBMP Katha No. 196/187, EPID No. 6538284977, Ward No. 42, 1st Main "
                    "Road, Hoysala Nagara, Goraguntepalya, Yeshwanthpur Hobli, Bangalore North Taluk, Bangalore, "
                    "Karnataka- 560022, India", "560022", [("N Lokesh", 100)], built=1450, carpet=1200)],
        [loan(900301, "A", "22-JUL-26", 6500000)],
        case="Bangalore BBMP address (Katha, EPID, Ward, Hobli, Taluk): house number found, but not whether it is "
             "a flat or a house; no survey/plot number; no title report (SI type default, owners not verified); "
             "proprietorship with only GST/Udyam/Shop numbers (no PAN, so a CKYC number is needed); Last Name ' '; "
             "e-mail with a hyphen (not allowed by CERSAI)."))

    # D04 - Hyderabad house with two house numbers; NO PAN customer; registered mortgage.
    hyd_home = "H.No. 1-9-374, Kushaiguda, Kapra, Hyderabad, Telangana- 500062, India"
    apps.append(application(
        "APPL00090004", "Hyderabad", "03-08-2026",
        [person(90041, "MOHD IRFAN", "primary", "09-JUL-84", "BJUPI6612R", addresses=[("ResidentialAddress", hyd_home)]),
         person(90042, "SALMA BEGUM", "Co-Applicant", "17-MAR-88", constitution="NO PAN",
                ids=[("Voter's ID", "ZXQ1234567")], mobile=None, phone="9848012345",
                addresses=[("ResidentialAddress", hyd_home)])],
        [collateral("COLL00090004", "CMS00090004", "RESIDENTIAL",
                    "H.No. 1-9-374 and 1-9-375, Situated at Kushaiguda, Kapra Village, Kapra Mandal, Under GHMC "
                    "Circle, Survey No. 45/A, Hyderabad, Telangana- 500062, India", "500062",
                    [("Mohd Irfan", 50), ("Salma Begum", 50)], built=2400, carpet=2000,
                    pdf=report("B", "Mohd Irfan", "Mohd Irfan and Salma Begum",
                               "House No. 1-9-374 and 1-9-375 admeasuring 222.96 sq.mtr built up area consisting of "
                               "ground and first floors constructed on the land bearing Survey No. 45/A situated at "
                               "Kushaiguda, Kapra Village, Kapra Mandal, Medchal-Malkajgiri District, Hyderabad 500 062",
                               mortgage="Registered Mortgage", status="Positive"))],
        [loan(900401, "A", "03-AUG-26", 7200000)],
        case="Two house numbers ('H.No. 1-9-374 and 1-9-375' -> 1-9-374/1-9-375); customer with constitution "
             "'NO PAN' and only a Voter's ID (CKYC number needed); no mobile in CKYC, so COLLATERAL PHONE_NUMBER is "
             "used; title report says Registered Mortgage (SI type 18); area in sq. m close to the built-up area."))

    # D05 - PLOT with a title report that says it is residential.
    apps.append(application(
        "APPL00090005", "Jaipur", "26-06-2026",
        [person(90051, "RAJENDRA PRASAD SHARMA", "primary", "02-OCT-70", "AOLPS9087H",
                addresses=[("ResidentialAddress", "Plot No. 12, Gopalpura Bypass, Near Triveni Nagar, Jaipur, "
                                                  "Rajasthan- 302018, India")]),
         person(90052, "priya sharma", "Co-Applicant", "21-APR-95", "FKLPS2231B", email="priya.sharma95@gmail.com",
                addresses=[("ResidentialAddress", "Plot No. 12, Gopalpura Bypass, Near Triveni Nagar, Jaipur, "
                                                  "Rajasthan- 302018, India")]),
         person(90053, "MAHESH CHAND SHARMA", "Guarantor", "15-JAN-66", "AZXPS1122C",
                addresses=[("ResidentialAddress", "22, Shanti Nagar, Durgapura, Jaipur, Rajasthan- 302018, India")])],
        [collateral("COLL00090005", "CMS00090005", "PLOT",
                    "Plot No. 45, Shri Ram Nagar Scheme, Near Water Tank, Jagatpura, Jaipur, Rajasthan- 302017, India",
                    "302017", [("Rajendra Prasad Sharma", 100)], carpet=1800, purpose="Vacant",
                    pdf=report("A", "Rajendra Prasad Sharma", "Shri Rajendra Prasad Sharma",
                               "Residential Plot No. 45, adm. 167.22 Sq. Mtrs., Plot area, in the scheme known as Shri "
                               "Ram Nagar, situated at Village Jagatpura, Tehsil Sanganer, District Jaipur, Rajasthan "
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
              addresses=[("OfficeAddress", "Survey No. 233/1, Halol GIDC Road, Halol, Panchmahal, Gujarat- 389350, "
                                           "India")]),
         person(90062, "BHAVESH PATEL", "Co-Applicant", "04-NOV-79", "AMPPP4455G",
                addresses=[("ResidentialAddress", "12, Ravji Residency, Halol, Panchmahal, Gujarat- 389350, India")])],
        [collateral("COLL00090006", "CMS00090006", "PLOT",
                    "Plot No. 12 to 14, Ravji Residency, Nr. Gokuldham Apartment, Survey No. 233/1, Halol, "
                    "Panchmahal, Gujarat- 389350, India", "389350", [("Bhavesh Patel", 100)], carpet=2700.75,
                    nature="GRAM PANCHAYAT", purpose="Vacant")],
        [loan(900601, "A", "09-SEP-26", 2500000)],
        case="PLOT with no title report: asset type defaults to 3 (Other, 'Plot'); a proprietorship tagged "
             "'individual' in CKYC Applicant Type (so CKYC has a date of birth but no incorporation date: "
             "registration date needed); proprietor and firm share one PAN; Gram Panchayat land."))

    # D07 - private limited company with CIN; khasra number only in the title report.
    apps.append(application(
        "APPL00090007", "Ghaziabad", "14-08-2026",
        [firm(90071, "AARAV INFRA PROJECTS PVT. LTD.", "Private Limited Company", "primary", "15-MAR-12",
              ids=[("PAN", "AAMCA1234C"), ("CIN No", "U45200UP2012PTC012345"), ("GST_NO", "09AAMCA1234C1Z2")],
              email="accounts@aaravinfra.co.in",
              addresses=[("OfficeAddress", "C-21, 3rd Floor, RDC Raj Nagar, Ghaziabad, Uttar Pradesh- 201002, India"),
                         ("AlternateBusinessAddress", "Plot No. 7, Site IV, Sahibabad Industrial Area, Ghaziabad, "
                                                      "Uttar Pradesh- 201010, India")]),
         person(90072, "DEEPAK KUMAR VERMA", "Co-Applicant", "19-JUL-76", "ACVPV5566D",
                addresses=[("ResidentialAddress", "Flat No. 704, Gaur Green Avenue, Indirapuram, Ghaziabad, Uttar "
                                                  "Pradesh- 201014, India")]),
         person(90073, "POOJA VERMA", "Co-Applicant", "08-JAN-80", "BDTPV7788E", email="pooja-verma@rediffmail.com",
                coll_email="poojaverma80@gmail.com",
                addresses=[("ResidentialAddress", "Flat No. 704, Gaur Green Avenue, Indirapuram, Ghaziabad, Uttar "
                                                  "Pradesh- 201014, India")])],
        [collateral("COLL00090007", "CMS00090007", "RESIDENTIAL",
                    "Flat No. 704, 7th Floor, Tower B-2, Gaur Green Avenue, Plot No. GH-01, Sector 4, Abhay Khand, "
                    "Indirapuram, Ghaziabad, Uttar Pradesh- 201014, India", "201014",
                    [("Deepak Kumar Verma", 50), ("Pooja Verma", 50)], built=1150, carpet=890,
                    pdf=report("B", "Aarav Infra Projects Pvt. Ltd.", "Mr. Deepak Kumar Verma and Mrs. Pooja Verma",
                               "Flat No. 704 admeasuring 1150 Sq. Ft. super area on the 7th floor in Tower B-2 of the "
                               "project known as Gaur Green Avenue constructed on the land bearing Khasra No. 1245 "
                               "situated at Village Makanpur, Abhay Khand, Indirapuram, District Ghaziabad"))],
        [loan(900701, "A", "14-AUG-26", 5200000)],
        case="Private limited company with its CIN in CKYC; office and alternate business addresses; e-mail with a "
             "hyphen in CKYC, valid one in COLLATERAL (used); plot number from the LMS address and khasra (survey) "
             "number from the title report; title-report area matches the built-up area, not the carpet area."))

    # D08 - Haryana revenue numbers; area only in the title report; PIN in two districts.
    apps.append(application(
        "APPL00090008", "Gurugram", "29-08-2026",
        [person(90081, "AMIT MALHOTRA", "primary", "12-DEC-83", "AJKPM3344F",
                addresses=[("ResidentialAddress", "Tower 4, Flat 1203, Central Park, Sector 42, Gurgaon, Haryana- "
                                                  "122002, India")]),
         person(90082, "NEHA MALHOTRA", "Co-Applicant", "27-MAY-86", "BLMPM5566G",
                addresses=[("ResidentialAddress", "Tower 4, Flat 1203, Central Park, Sector 42, Gurgaon, Haryana- "
                                                  "122002, India")])],
        [collateral("COLL00090008", "CMS00090008", "RESIDENTIAL",
                    "Residential Apartment No. 9, on Second Floor, in Block-C, Rosewood City, Sohna Road, Khewat No. "
                    "123, Mustatil No. 45, Killa No. 12/2, Village Badshahpur, District Gurgaon, Gurgaon, Haryana- "
                    "122502, India", "122502", [("Amit Malhotra", 100)],
                    pdf=report("A", "Amit Malhotra", "Mr. Amit Malhotra",
                               "Residential Apartment No. 9 on the Second Floor, adm. 111.48 Sq. Mtrs., Carpet area, "
                               "in Block-C of the project known as Rosewood City, constructed on land bearing Khewat "
                               "No. 123, Mustatil No. 45, Killa No. 12/2, situated at Village Badshahpur, Tehsil and "
                               "District Gurgaon, Haryana."))],
        [loan(900801, "A", "29-AUG-26", 9500000)],
        case="Haryana khewat/mustatil/killa numbers; no area in COLLATERAL, so the title report's area in sq. m is "
             "used (unit 7); PIN 122502 belongs to two districts (Gurgaon, Rewari) and the address names Gurgaon."))

    # D09 - HUF; Delhi 'Property bearing No.', khasra; mixed usage; no title report.
    apps.append(application(
        "APPL00090009", "Delhi 1", "07-07-2026",
        [firm(90091, "SHREE BALAJI (HUF)", "Hindu Undivided Company", "primary", "01-APR-98", "AAHHS5678B",
              addresses=[("OfficeAddress", "B-45, Ground Floor, Mayur Vihar Phase 1, East Delhi, Delhi- 110091, "
                                           "India")]),
         person(90092, "SANJAY GUPTA", "Co-Applicant", "30-SEP-72", "AFGPG9900J",
                addresses=[("ResidentialAddress", "B-45, First Floor, Mayur Vihar Phase 1, East Delhi, Delhi- 110091, "
                                                  "India")])],
        [collateral("COLL00090009", "CMS00090009", "MIXED USAGE",
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
              ids=[("PAN", "BCDPK4321E"), ("GST_NO", "24BCDPK4321E1Z9")],
              addresses=[("OfficeAddress", "SHOP NO C 9, GROUND FLOOR, LAXMI CHSL, C S ROAD NO 9, DAHISAR EAST, "
                                           "MUMBAI, Surat, Gujarat- 395007, India")]),
         person(90102, "KISHAN LAL AGARWAL", "Co-Applicant", "11-NOV-69", "BCDPK4321E",
                addresses=[("ResidentialAddress", "A-12, Shanti Kunj Society, Ghod Dod Road, Surat, Gujarat- 395007, "
                                                  "India")])],
        [collateral("COLL00090010", "CMS00090010", "COMMERCIAL",
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
         person(90112, "SENTHIL KUMAR R", "Co-Applicant", "03-MAR-81", "BXRPS1234K",
                addresses=[("ResidentialAddress", "Door No. 12/4, 2nd Street, Anna Nagar East, Chennai, Tamil Nadu- "
                                                  "600040, India")]),
         person(90113, "MEENA SENTHIL", "Co-Applicant", "16-JUN-84", "CZYPS4321L",
                addresses=[("ResidentialAddress", "Door No. 12/4, 2nd Street, Anna Nagar East, Chennai, Tamil Nadu- "
                                                  "600040, India")])],
        [collateral("COLL00090011", "CMS00090011", "COMMERCIAL",
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
    joshi_home = "Flat No. 1301, 13th Floor, A Wing, Rukmini Vaibhav CHS Ltd, Veera Desai Road, Andheri West, Mumbai, " \
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
        [collateral("COLL00090012", "CMS00090012", "RESIDENTIAL",
                    "Flat No. 1301 & 1302, 13th Floor, A Wing, Rukmini Vaibhav CHS Ltd, Veera Desai Road, Azad Nagar, "
                    "Andheri West, Mumbai, Maharashtra- 400058, India", "400058",
                    [(" Mr. Mahesh Chandra Joshi", 25), ("Mrs. Rekha M. Joshi", 25), ("Mr. Rohan Joshi", 25),
                     ("Ms. Isha Joshi", 25)], built=1550, carpet=1300, market=42000000,
                    pdf=report("B", "Mahesh Chandra Joshi", "Mahesh Chandra Joshi, Rekha Mahesh Joshi, Rohan Mahesh "
                               "Joshi and Isha Mahesh Joshi",
                               "Flat No. 1301 and 1302 admeasuring 120.50 sq.mtr carpet area on the 13th floor in A "
                               "Wing of the building known as Rukmini Vaibhav of the Society known as 'Rukmini Vaibhav "
                               "Co-operative Housing Society Ltd.', constructed on land bearing C.T.S. No. 781, 781/1 "
                               "to 781/5 of Village Andheri, Taluka Andheri, Veera Desai Road, Azad Nagar, Andheri "
                               "(West), Mumbai 400058", status="Positive"))],
        [loan(901201, "A", "01-AUG-26", 15000000)],
        case="Two flats in one collateral ('Flat No. 1301 & 1302' -> 1301/1302); four owners at 25% each with "
             "Mr/Mrs/Ms; a guarantor; CKYC First Disbursal Date differs from COLLATERAL Disbursal Date (CKYC "
             "used); a 2-digit birth year that is 20YY (15-MAR-05 -> 2005)."))

    # D13 - one application, two collaterals, an active and a cancelled loan account.
    pune_home = "Row House No. 7, Shubh Kalyan Society, Wakad, Pune, Maharashtra- 411057, India"
    apps.append(application(
        "APPL00090013", "Pune", "11-09-2026",
        [person(90131, "PRAKASH DESHMUKH", "primary", "18-AUG-75", "AHJPD6677K",
                addresses=[("ResidentialAddress", pune_home)]),
         person(90132, "SWATI DESHMUKH", "Co-Applicant", "02-APR-79", "BNMPD8899L",
                addresses=[("ResidentialAddress", pune_home)]),
         firm(90133, "DESHMUKH ENTERPRISES", "Sole Proprietorship Firm", "Co-Applicant", "01-JUN-10", "AHJPD6677K",
              addresses=[("OfficeAddress", "Shop No. 4, Wakad Chowk, Wakad, Pune, Maharashtra- 411057, India")])],
        [collateral("COLL00090013", "CMS00090013", "RESIDENTIAL",
                    "Row House No. 7, Shubh Kalyan Society, S. No. 45/2, Wakad, Pune, Maharashtra- 411057, India",
                    "411057", [("Prakash Deshmukh", 50), ("Swati Deshmukh", 50)], built=1800, carpet=1500,
                    classification="Row House",
                    pdf=report("A", "Prakash Deshmukh", "Mr. Prakash Deshmukh & Mrs. Swati Deshmukh",
                               "Row House No. 7, adm. 1800 Sq. Fts., Built up area, in the Society known as ‘Shubh "
                               "Kalyan Co-operative Housing Society Ltd.’, constructed on land bearing S. No. "
                               "45/2, situated at Village Wakad, Taluka Mulshi, District Pune.",
                               owners_item12="Mr. Prakash Deshmukh")),
         collateral("COLL00090014", "CMS00090014", "RESIDENTIAL",
                    "Flat No. B-302, 3rd Floor, Ganga Kingston, Survey No. 12/1A, Pimple Saudagar, Pune, Maharashtra- "
                    "411027, India", "411027", [("Prakash Deshmukh", 100)], built=1100, carpet=950, purpose="Rented")],
        [loan(901301, "A", "11-SEP-26", 9000000), loan(901302, "X", "05-SEP-26", 9000000)],
        case="Two collaterals on one application (the borrowers are listed once in the required-details file); "
             "a cancelled loan account (status X, nothing disbursed) next to the active one - its rows are "
             "ignored; row house = structure with land; the title report's item 12 names fewer owners than item 4 "
             "(warning)."))

    # D14 + D15 - the same property (Global collateral No) secures two applications.
    express_zone = "Unit No. 305, 3rd Floor, Express Zone, Western Express Highway, Malad East, Mumbai, Maharashtra- " \
                   "400097, India"
    harish = person(90141, "HARISH MEHTA", "primary", "07-JUN-68", "AAQPM1212M", email="harish.mehta68@gmail.com",
                    addresses=[("ResidentialAddress", "B-1204, Oberoi Splendor, JVLR, Andheri East, Mumbai, "
                                                      "Maharashtra- 400060, India")])
    apps.append(application(
        "APPL00090014", "Mumbai 3", "15-07-2026", [harish],
        [collateral("COLL00090016", "CMS00090099", "COMMERCIAL", express_zone, "400097", [("Harish Mehta", 100)],
                    built=900, carpet=760, purpose="Rented")],
        [loan(901401, "A", "15-JUL-26", 11000000)],
        case="Property also secures APPL00090015 (same Global collateral No): TYPE_OF_CHARGE needs confirming; "
             "single borrower; the city is MALAD (more specific than the general 'Mumbai')."))
    apps.append(application(
        "APPL00090015", "Mumbai 3", "20-08-2026",
        [firm(90151, "MEHTA TRADING CO.", "Partnership Firm", "primary", "12-DEC-01", "AAAFM3434N",
              email="accounts@mehtatrading.in", addresses=[("OfficeAddress", express_zone)]),
         dict(harish, ctype="Co-Applicant")],
        [collateral("COLL00090017", "CMS00090099", "COMMERCIAL", express_zone, "400097", [("Harish Mehta", 100)],
                    built=900, carpet=760, purpose="Rented")],
        [loan(901501, "A", "20-AUG-26", 5000000)],
        case="Second loan on the same property as APPL00090014 (TYPE_OF_CHARGE to confirm); the same customer "
             "(same Neo CIF ID) on two applications."))

    # D16 - cancelled only; D17 - closed only: both left out.
    apps.append(application(
        "APPL00090016", "Pune", None,
        [person(90161, "SURESH NAIK", "primary", "12-MAR-77", "AKLPN2323M",
                addresses=[("ResidentialAddress", "Flat No. 11, Sunshine Apartment, Baner, Pune, Maharashtra- 411045, "
                                                  "India")])],
        [collateral("COLL00090019", "CMS00090019", "RESIDENTIAL",
                    "Flat No. 11, 1st Floor, Sunshine Apartment, Baner, Pune, Maharashtra- 411045, India", "411045",
                    [("Suresh Naik", 100)], built=700, carpet=600,
                    pdf=report("A", "Suresh Naik", "Mr. Suresh Naik",
                               "Flat No. 11 on the 1st Floor, adm. 55.74 Sq. Mtrs., Carpet area, in the Society known "
                               "as 'Sunshine Apartment Co-op. Hsg. Soc. Ltd.', situated at Village Baner, Taluka "
                               "Haveli, District Pune."))],
        [loan(901601, "X", "02-SEP-26", 3000000)],
        case="Only a cancelled loan (X): the whole application is left out, and its title report is ignored."))
    apps.append(application(
        "APPL00090017", "Chennai", "10-07-2026",
        [person(90171, "LATA IYER", "primary", "29-JUL-58", "AMNPI3434P",
                addresses=[("ResidentialAddress", "Flat No. 3, Lakshmi Apartments, Adyar, Chennai, Tamil Nadu- 600020, "
                                                  "India")])],
        [collateral("COLL00090020", "CMS00090020", "RESIDENTIAL",
                    "Flat No. 3, Lakshmi Apartments, Adyar, Chennai, Tamil Nadu- 600020, India", "600020",
                    [("Lata Iyer", 100)], built=650, carpet=560)],
        [loan(901701, "C", "10-JUL-26", 2000000)],
        case="Only a closed loan (C): the application is left out."))

    # D18 - no owners in COLLATERAL; the title report names them.
    kalyan_home = "Flat No. 3, Om Sai Apartment, Katemanivali, Kalyan East, Thane, Maharashtra- 421306, India"
    apps.append(application(
        "APPL00090018", "KALYAN", "25-08-2026",
        [person(90181, "GANESH PANDURANG MORE", "primary", "01-JAN-80", "ALKPM7878N",
                addresses=[("ResidentialAddress", kalyan_home)]),
         person(90182, "SARIKA GANESH MORE", "Co-Applicant", "13-MAY-85", "CMNPM8989P",
                addresses=[("ResidentialAddress", kalyan_home)])],
        [collateral("COLL00090021", "CMS00090021", "RESIDENTIAL",
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
                addresses=[("ResidentialAddress", "P.No. A-9, Marwar Nagar, Mahamandir, Jodhpur, Rajasthan- 342001, "
                                                  "India")]),
         person(90192, "KAMLA DEVI", "Co-Applicant", "01-JAN-75", "CRTPD6767R", mobile=None,
                addresses=[(None, None)], ckyc_addresses=[])],
        [collateral("COLL00090022", "CMS00090022", "PG",
                    "P.No. A-9, Marwar Nagar, Near Bheel Basti, Mahamandir, Jodhpur, Rajasthan- 342001, India",
                    "342001", [], built=2200, carpet=1900, purpose="Multi Tenanted")],
        [loan(901901, "A", "16-JUL-26", 4200000)],
        case="No owners in COLLATERAL and no title report: owners must be typed; property type PG (residential); "
             "'P.No.' = plot number; a co-applicant with no address in CKYC or COLLATERAL (address needed)."))

    # D20 - the title report names an extra owner who is not a borrower: third-party mortgage.
    vashi = "Flat No. 504, Sai Darshan CHS, Sector 17, Vashi, Navi Mumbai, Thane, Maharashtra- 400703, India"
    apps.append(application(
        "APPL00090020", "VASHI BRANCH", "02-09-2026",
        [person(90201, "ARJUN NAIR", "primary", "22-OCT-88", "BSDPN2323S", addresses=[("ResidentialAddress", vashi)]),
         person(90202, "LAKSHMI NAIR", "Co-Applicant", "30-NOV-90", "CTFPN3434T",
                addresses=[("ResidentialAddress", vashi)])],
        [collateral("COLL00090023", "CMS00090023", "RESIDENTIAL",
                    "Flat No. 504, 5th Floor, Sai Darshan CHS, Sector 17, Vashi, Navi Mumbai, Thane, Maharashtra- "
                    "400703, India", "400703", [("Arjun Nair", 100)], built=850, carpet=700,
                    pdf=report("B", "Arjun Nair", "Mr. Arjun Nair and Mrs. Radha Nair",
                               "Flat No. 504 admeasuring 65.03 sq.mtr carpet area on the 5th floor in the building "
                               "known as Sai Darshan of the Society known as 'Sai Darshan Co-operative Housing Society "
                               "Ltd.', constructed on the land bearing Plot No. 23, Sector 17 of Village Vashi, Navi "
                               "Mumbai, Taluka and District Thane"))],
        [loan(902001, "A", "02-SEP-26", 6800000)],
        case="COLLATERAL and the title report name different owners: the title report's owners are used and "
             "flagged; its extra owner (Radha Nair) is not a borrower, so this is a third-party mortgage and the "
             "SI is left out of the .DAT file; plot number only in the title report. Note: the law firm's own "
             "PIN (400703) matches this property's PIN by chance."))

    # D21 - an owner name that only probably matches a borrower.
    apps.append(application(
        "APPL00090021", "Hyderabad", "19-08-2026",
        [person(90211, "BHANU KATAM", "primary", "12-JUN-87", "BUGPK4545U",
                addresses=[("ResidentialAddress", "Flat No. 202, Sri Sai Residency, Madhapur, Hyderabad, Telangana- "
                                                  "500081, India")]),
         person(90212, "RAVI KATAM", "Co-Applicant", "25-DEC-84", "AVHPK5656V",
                addresses=[("ResidentialAddress", "Flat No. 202, Sri Sai Residency, Madhapur, Hyderabad, Telangana- "
                                                  "500081, India")])],
        [collateral("COLL00090024", "CMS00090024", "RESIDENTIAL",
                    "Flat No. 202, 2nd Floor, Sri Sai Residency, Plot No. 18 & 19, Survey No. 112, Madhapur, "
                    "Hyderabad, Telangana- 500081, India", "500081", [("Smt. Banu K", 100)], built=1250, carpet=1050)],
        [loan(902101, "A", "19-AUG-26", 5600000)],
        case="Owner 'Smt. Banu K' only probably matches borrower 'BHANU KATAM': the Check and correct sheet asks "
             "Y/N; both survey and plot numbers in the address; no title report."))

    # D22 - the title report belongs to a different property: SI held.
    apps.append(application(
        "APPL00090022", "Vadodara", "27-08-2026",
        [person(90221, "NILESH SHAH", "primary", "08-AUG-82", "AWJPS6767W",
                addresses=[("ResidentialAddress", "Flat No. 101, Shreeji Heights, Alkapuri, Vadodara, Gujarat- 390007, "
                                                  "India")]),
         person(90222, "HETAL SHAH", "Co-Applicant", "19-SEP-85", "CXKPS7878X",
                addresses=[("ResidentialAddress", "Flat No. 101, Shreeji Heights, Alkapuri, Vadodara, Gujarat- 390007, "
                                                  "India")])],
        [collateral("COLL00090025", "CMS00090025", "RESIDENTIAL",
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
    indore = "House No. 45, Scheme No. 54, Vijay Nagar, Indore, Madhya Pradesh- 452010, India"
    apps.append(application(
        "APPL00090023", "INDORE", "12-08-2026",
        [person(90231, "VIKAS JAIN", "primary", "14-APR-83", "AYLPJ8989Y", addresses=[("ResidentialAddress", indore)]),
         person(90232, "RITU JAIN", "Co-Applicant", "03-JUL-86", "CZMPJ9090Z", addresses=[("ResidentialAddress", indore)]),
         person(90233, "RAKESH .", "Co-Applicant", "20-OCT-92", "DAPPR1010A", first="RAKESH", middle=None, last=".",
                addresses=[("ResidentialAddress", indore)])],
        [collateral("COLL00090026", "CMS00090026", "RESIDENTIAL", indore, "452010", [("Vikas Jain", 100)],
                    built=2000, carpet=1700,
                    pdf=report("scanned", "Vikas Jain", "Mr. Vikas Jain",
                               "House No. 45, Scheme No. 54, Vijay Nagar, Indore, adm. 157.93 Sq. Mtrs."))],
        [loan(902301, "A", "12-AUG-26", 5500000)],
        case="The title report is a scan with no text: it is reported and not used (SI type default, owners not "
             "verified); a customer name with a dot ('RAKESH .'); house = structure with land."))

    # D24 - industrial gala; LMS PIN not in the master; area conflict; two active loans.
    apps.append(application(
        "APPL00090024", "Mumbai 2", "06-08-2026",
        [firm(90241, "DEWAN POLYMERS PRIVATE LIMITED", "Private Limited Company", "primary", None,
              ids=[("PAN", "AAJCD4455F"), ("UDHYAM_REGISTRATION", "UDYAM-MH-33-0123456")],
              applicant_type="individual", dob="12-MAY-10",
              addresses=[("OfficeAddress", "Gala No. 12, Dewan and Shah Industrial Estate No. 2, Navghar, Vasai East, "
                                           "Palghar, Maharashtra- 401208, India")]),
         person(90242, "FARHAN DEWAN", "Co-Applicant", "01-DEC-80", "BZNPD1212B",
                addresses=[("ResidentialAddress", "Flat No. 9, Evershine City, Vasai East, Palghar, Maharashtra- "
                                                  "401208, India")])],
        [collateral("COLL00090027", "CMS00090027", "INDUSTRIAL",
                    "Gala No. 12, Ground Floor, Dewan and Shah Industrial Estate No. 2, Navghar, Vasai East, Palghar, "
                    "Maharashtra- 401208, India", "401299", [("Dewan Polymers Pvt Ltd", 100)], built=1100,
                    carpet=1000, purpose="Self Occupied",
                    pdf=report("B", "Dewan Plastics", "M/s. Dewan Polymers Pvt. Ltd.",
                               "Gala No. 12 admeasuring 60 sq.mtr built up area on the ground floor in the building "
                               "known as Dewan and Shah Industrial Estate No. 2 constructed on the land bearing Survey "
                               "No. 88, Hissa No. 4 situated at Village Navghar, Vasai East, Taluka Vasai, District "
                               "Palghar 401 208", nature="Industrial"))],
        [loan(902401, "A", "06-AUG-26", 5000000), loan(902402, "A", "20-SEP-26", 5500000, 2000000)],
        case="Industrial gala (asset type 3); COLLATERAL Property Pincode 401299 is not in the CERSAI master, so "
             "the PIN in the title report (401 208) is used; LMS and title-report areas differ by more than 10%; "
             "two active loan accounts with different sanctioned amounts; a company tagged 'individual' (no "
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
        [collateral("COLL00090028", "CMS00090028", "RESIDENTIAL",
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
        [collateral("COLL00090029", "CMS00090029", "RESIDENTIAL",
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
        [collateral("COLL00090030", "CMS00090030", "COMMERCIAL",
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
                addresses=[("ResidentialAddress", "House No. 1018, Sector-55, Ballabgarh, Haryana- 121004, India")]),
         firm(90282, "CHAND HARDWARE STORE", "Sole Proprietorship Firm", "Co-Applicant", "01-AUG-08",
              ids=[("Unsupported ", "06AFYPC7878G1Z5"), ("KYC Identification Number", "06AFYPC7878G1Z5"),
                   ("UDYAM", "UDYAM-HR-05-0012345")],
              addresses=[("OfficeAddress", "Shop No. 3, Main Market, Sector-55, Ballabgarh, Haryana- 121004, India")])],
        [collateral("COLL00090031", "CMS00090031", "RESIDENTIAL",
                    "House No. 1018, (EWS), area measuring 49.94 sq. yds., in Sector-55, Housing Board Colony, "
                    "Ballabgarh, Haryana- 121004, India", "121004", [("Ramesh Chand", 100)], built=450, carpet=400)],
        [loan(902801, "A", "13-AUG-26", 1500000)],
        case="PIN 121004 is in two districts (Faridabad, Palwal) and the address names neither: district and "
             "city needed (property and borrowers); ID types 'Unsupported ' (trailing space), Aadhar, passport, "
             "and a 'KYC Identification Number' that is really a GST number (not 14 digits, so not used)."))

    # D29 - old/new survey numbers with Hissa; three owners with full names in the title report.
    mira = "Flat No. C-3002, Tower 3, Aaradhya Parkwood, Mira Road East, Thane, Maharashtra- 401107, India"
    apps.append(application(
        "APPL00090029", "Thane", "04-09-2026",
        [person(90291, "IMRAN QURESHI", "primary", "28-FEB-82", "BGRPQ8989H", addresses=[("ResidentialAddress", mira)]),
         person(90292, "SALMA QURESHI", "Co-Applicant", "09-JUN-85", "CHSPQ9090J",
                addresses=[("ResidentialAddress", mira)]),
         person(90293, "YUSUF QURESHI", "Co-Applicant", "11-NOV-55", "AJTPQ1212K",
                addresses=[("ResidentialAddress", mira)])],
        [collateral("COLL00090032", "CMS00090032", "RESIDENTIAL",
                    "Flat No. C-3002, 30th Floor, Tower 3, Clove, Aaradhya Parkwood, Old S. No. 260(Pt), New S. No. 12, "
                    "Hissa No. 2, Village Mahajanwadi, Mira Road East, Thane, Maharashtra- 401107, India", "401107",
                    [("Imran Qureshi", 33.33), ("Salma Qureshi", 33.33), ("Yusuf Qureshi", 33.34)], built=520,
                    carpet=452,
                    pdf=report("B", "Imran Qureshi", "Imran Yusuf Qureshi, Salma Imran Qureshi and Yusuf Ahmad Qureshi",
                               "Flat No. C-3002 admeasuring 41.99 sq.mtr RERA carpet area on the 30th floor in "
                               "Tower-3(Clove) in the new building known as Aaradhya Parkwood 1 of the project known as "
                               "Aaradhya Parkwood constructed on the land bearing Old S. No. 260(Pt) and New S. No. 12 "
                               "and Hissa No. 2 of Village Mahajanwadi, Taluka and District Thane",
                               transaction="BT + Top Up"))],
        [loan(902901, "A", "04-SEP-26", 7800000)],
        case="Old and new survey numbers with Hissa (only the new one is used: 12/2); three owners at "
             "33.33/33.33/33.34%; the title report lists 'A, B and C' with full names."))
    return apps


# Title reports saved under a wrong or odd file name
EXTRA_PDFS = [
    ("APPL00090027_COLL0009030.pdf", "APPL00090027", "COLL00090030",
     "misnamed: the collateral ID has 7 digits (reported; the collateral is treated as having no title report)"),
    ("APPL00090001_COLL00090002.pdf", "APPL00090002", "COLL00090002",
     "misfiled: COLL00090002 belongs to APPL00090002, not APPL00090001 (reported)"),
    ("Title report APPL00090021.pdf", "APPL00090021", "COLL00090024",
     "a name that does not follow <APPL>_<COLL>.pdf (reported)"),
]


# --------------------------------------------------------------------------
# The LMS extracts
# --------------------------------------------------------------------------
def dod(first_disbursal):
    if not first_disbursal:
        return None
    day, month, year = first_disbursal.split("-")
    return f"{day}-{MONTHS[int(month) - 1]}-{year[2:]}"


def ckyc_rows(apps):
    rows = []
    for a in apps:
        active = [l for l in a["loans"] if l["status"] == "A"]
        closed = [l for l in a["loans"] if l["status"] == "C"]
        account = (active or closed or [None])[0]
        for c in a["customers"]:
            for id_type, id_number in (c["ids"] or [(None, None)]):
                for address in (c["ckyc_addresses"] or [None]):
                    rows.append([a["first_disbursal"], a["branch"], dod(a["first_disbursal"]), a["app"],
                                 account["lan"] if account else None, c["ctype"], c["applicant_type"], c["name"],
                                 c["constitution"], c["cif"], c["dob"], c["inc"], c["first"], c["middle"], c["last"],
                                 id_type, id_number, c["mobile"], c["email"], address, c["risk"],
                                 "Yes" if c["mobile"] else None, a["product"], a["received"]])
    return rows


def collateral_rows(apps):
    rows = []
    for a in apps:
        for l in a["loans"]:
            for col in a["collaterals"]:
                owners = col["owners"] or [(None, None)]
                status = None if not col["owners"] else ("Single" if len(col["owners"]) == 1 else "Joint")
                for c in a["customers"]:
                    for owner, share in owners:
                        for address_type, address in (c["addresses"] or [(None, None)]):
                            rows.append([
                                a["app"], l["lan"], l["status"], l["date"], a["branch"], c["name"], c["ctype"],
                                c["cif"], col["purpose"], col["coll"], col["cms"], col["ownership"], status, owner,
                                share, col["ptype"], col["classification"], col["nature"], col["address"], col["pin"],
                                "48", "1 Vanilla Income Program", col["built"], col["carpet"], col["age"],
                                col["market"] or int(l["sanctioned"] * 1.8), l["sanctioned"], l["disbursed"],
                                c["phone"], c["coll_email"], c["pan"], address_type, address])
    return rows


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
          "letterhead": ["Demo_______", "Associates Advocates and Legal Consultants (Notary, Govt of India.)",
                         "_______________________________www.demo-associates.example_____________________",
                         "Office 9, Sector 19A, Plot No. 88, Opp. Seawood Station (East), Nerul,",
                         "Navi Mumbai- 400706, Tel. - 022-00000000, Mob-9000000000. demo.associates@example.com.",
                         "_________________________________________________________________________"],
          "firm": "Demo Associates", "advocate": "(Adv. A. B. Sample)"},
    "B": {"columns": [54, 95, 201, 527], "continuation": True,
          "letterhead": ["SAMPLE LAW ASSOCIATES", "Legal Advisors",
                         "Office at: 315, 3rd floor, Test Tower, Plot No. 78, Sector 17, Vashi, Navi Mumbai- 400703,",
                         "Email- samplelaw@example.com Mob. No. 9000000001 / 9000000002."],
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
        self.doc.save(str(path))


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
        scan.save(str(path))
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
    owners12 = spec["owners_item12"].rstrip(".")
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
    w.paragraph("We have gone through the complete chain of documents, which were produced by the ARGFL before me. "
                "All documents are proper and in order. After perusal of all the documents as per Serial No.5, our "
                "Equitable Mortgage will be valid with regard to said property, which shall be enforceable under "
                "the SARFAESI ACT.", strike=b_style)
    w.line("Thanking you.")
    w.line(f"For {w.style['firm']}", x=380)
    w.line(w.style["advocate"], x=380)


# --------------------------------------------------------------------------
def main():
    apps = build_applications()
    for folder in (CKYC_DIR, COLL_DIR, LEGAL_DIR):
        folder.mkdir(parents=True, exist_ok=True)
        for old in folder.iterdir():
            if old.is_file():
                old.unlink()
    note = "Dummy data made by dev_tools/make_dummy_data.py - not an LMS export."
    write_extract(CKYC_DIR / "CKYC.xlsx", CKYC_HEADERS, ckyc_rows(apps), note)
    write_extract(COLL_DIR / "COLLATERAL.xlsx", COLL_HEADERS, collateral_rows(apps), note)

    specs = {}
    for a in apps:
        for col in a["collaterals"]:
            if col["pdf"]:
                specs[col["coll"]] = (a, col)
                name = col["pdf"]["file_name"] or f"{a['app']}_{col['coll']}.pdf"
                make_report(LEGAL_DIR / name, col["pdf"])
    for name, app, coll, _ in EXTRA_PDFS:
        a, col = next((a, c) for a in apps for c in a["collaterals"] if c["coll"] == coll)
        spec = col["pdf"] or report("A", a["customers"][0]["name"].title(),
                                    "; ".join(o.strip() for o, _ in col["owners"]) or "-",
                                    col["address"].rsplit(",", 3)[0])
        make_report(LEGAL_DIR / name, spec)
    (LEGAL_DIR / "notes.txt").write_text("Not a PDF: the tool ignores this file.\n", encoding="utf-8")

    lines = ["# Dummy data cases", "",
             "Made by `dev_tools/make_dummy_data.py`. All names, PANs, phone numbers, e-mails and law firms are "
             "made up; PIN codes and place names are real. Run the tool on this data with `--dummy`.", "",
             "| Application | Collateral(s) | What it tests |", "|---|---|---|"]
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
    print(f"Written to {OUT}")


if __name__ == "__main__":
    main()
