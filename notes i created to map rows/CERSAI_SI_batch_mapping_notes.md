# CERSAI 2.0 SI Registration Batch File: Automation Mapping Notes

> **Status:** implemented in `cersai_tool/` (2026-09-30). This document holds the rules the code follows; `README.md` explains how to run it.
> **Prepared:** 2026-09-27, from the files committed at `5aea88d` ("demo output attached").
> **Updated:** 2026-09-28. Only **active loans** (`Loan Status = A` in COLLATERAL) are in scope; all counts below use that scope.
> **Updated:** 2026-09-30. Added the legal title reports (`Input data/Legal/`): naming convention, extractable fields and mapping logic are in **§15**, and the field tables in §7 now show which values come from them. Input file paths updated to the new `Input data/` sub-folders. **Source priority (decided 2026-09-30):** the LMS input data (CKYC, COLLATERAL) comes first, and a title report is read only for values the input data doesn't have (§15.6). **Property owners** come from COLLATERAL too, checked against the title report for every collateral; the title report supplies them only when COLLATERAL has none, or adds one when the COLLATERAL shares show an owner is missing (§6.3).
> **Updated:** 2026-09-30 (implementation). The hand-made `Input data/Req Details/Req details.xlsx` is replaced by a **required-details file** that step 1 generates for each batch (`Output data/<batch>/2 Required details_<batch>.xlsx`, §9.3); the old file was deleted. Rules refined while building and testing on the full extract are in §7.3, §8.3, §8.5 and §15.
> **Scope:** the `introduction` (FH), `si` (SI), `immovable_asset` (IMM), `borrower` (BOR) and `charge_holder` (CHG) sheets of the format description.

---

## 1. Project description

Anand Rathi Global Finance Limited (ARGFL) has to register every security interest it creates over a borrower's property with **CERSAI**, the Central Registry of Securitisation Asset Reconstruction and Security Interest of India. In the CERSAI entity master, ARGFL is entity code **`J8084`**. For the loans in this data (mainly loans against property, plus some home loans), the security interest is a **mortgage on immovable property**.

CERSAI 2.0 accepts these registrations in bulk as a **pipe-separated (`|`) `.DAT` batch file**, laid out as defined in `CERSAI2_SIRegistration_BatchFormatDescription.xlsx`. Each file starts with one file-header line. Then, for each security interest, it has one `SI` line, one `IMM` line for the property, one `BOR` line per borrower and one `CHG` line for the lender.

**How it is done today:** someone builds the file by hand in Excel (`DAT FILE WORKING 1.xlsx`). They copy each loan's details from the LMS exports, look up CERSAI master codes (city, district, state, borrower type and so on), type them into cells, and join the cells with `&"|"&` formulas. This is slow. The demo file also shows how easy it is to shift a column, mistype a code or leave placeholder text in the output (see §10).

**What the automation does:** a Python program (`cersai_tool/`, run in three steps):

1. reads the two LMS extracts: `CKYC.xlsx` (borrowers) and `COLLATERAL.xlsx` (loans, properties and owners);
2. reads the advocate's **title report PDF for each collateral** (`Input data/Legal/`), but only for values the LMS extracts don't have or that need verifying: always the type of mortgage and the owners, and otherwise gaps such as survey numbers or missing address details (§15);
3. merges in the values that neither source holds, which operations staff type into a required-details file the program generates for each batch (§9.3);
4. converts every value into CERSAI's format and master codes, using `CERSAI2.0_Masters.xls`;
5. validates every field and every cross-field rule in the format description;
6. writes an upload-ready `.DAT` file, working and final review workbooks, a checks report (what was left out and why), and the required-details file listing whatever still needs manual input.

**Loans processed:** only **active loans** (`Loan Status = A` in `COLLATERAL.xlsx`). Cancelled (`X`) and closed (`C`) loan accounts are excluded.

**Out of scope for now:** movable and intangible assets, attachment orders, consortium rows and source-of-funds rows. Asset-owner (`ASO`) rows are also out of scope, but any third-party mortgage needs them, so see §6.3 and §11.

---

## 2. Files reviewed

| File | What it is | Key facts |
|---|---|---|
| `Input data/CKYC/CKYC.xlsx` | LMS borrower/KYC extract. Sheet `Export Worksheet` holds the data and sheet `SQL` holds the query that produced it. | 4,726 rows × 24 columns. **One row per customer per ID document**, so a customer with a PAN and a Udyam number appears twice. 1,177 applications and 4,256 (application, customer) pairs. Disbursal window in the query: 01-APR-26 to 23-SEP-26. Despite the file name, it contains **no CKYC (KIN) numbers**. |
| `Input data/COLLATERAL/COLLATERAL.xlsx` | LMS loan, property and owner extract (sheets `Export Worksheet` and `SQL`). | 7,685 rows × 33 columns. Grain: application × loan account × customer × collateral × owner (× address type). 1,177 applications, 1,317 loan accounts, 1,341 collateral IDs. |
| `Output data/<batch>/2 Required details_<batch>.xlsx` | The **required-details file**, generated by step 1 for each batch. It replaces the earlier hand-made draft `Input data/Req Details/Req details.xlsx` (8 field names, no data and no keys), which was deleted. | Lists only what the LMS and the title report don't give or what needs checking, keyed by collateral and by borrower. Structure in §9.3; §9.1 shows where the draft's 8 items went. |
| `Input data/Legal/<Application number>_<Collateral ID>.pdf` | Advocates' **title reports**, one PDF per collateral. | 4 demo files from 2 law firms, with made-up filenames. All follow ARGFL's standard 13-item title-report template and have a text layer (no OCR needed). Details in §15. |
| `Output Excel Description/CERSAI2_SIRegistration_BatchFormatDescription.xlsx` | The official field-by-field layout. **This is the primary reference.** | 11 sheets; 5 are in scope. Batch file version 1. |
| `Output excel demo file/DAT FILE WORKING 1.xlsx` | A hand-made sample for one loan (`APPL00027326` / `COLL00266688`), with a threaded comment on most cells explaining its source. | Useful for intent. It contains **35 issues**, several of which would get the file rejected (§10). |
| `Master data/CERSAI2.0_Masters.xls` | CERSAI master tables, **version 29.3, last updated 04-09-2026**. | CITY 45,899 rows · DISTRICT 764 · STATE 36 · PIN_CODE 20,832 (PIN, district) pairs · COUNTRY · GENDER · BORROWER_ASSET_OWNER_TYPE · SI type · financing type · asset category/type/sub-type · AREA_UNIT · ENTITY_CODE. |
| `Master data/crs2_addrss_mstr_pincd_analysis_20260416.xlsx` | An April 2026 analysis of PIN-code master changes: `addition_smry`, `removal_smry`, `Detailed_list` (1,972 "to be added", 20,093 "to be removed") and `effective_master`. | Every (PIN, district) pair in `effective_master` already exists in the v29.3 PIN master, which has 30 extra pairs. **Use v29.3 as the authority**; this file is background only. |

---

## 3. End-to-end workflow

```
 ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐  ┌──────────────────┐
 │ LMS extracts    │  │ Legal title     │  │ CERSAI masters  │  │ Required details │
 │ CKYC.xlsx       │  │ reports (PDF),  │  │ v29.3 (.xls),   │  │ manual values    │
 │ COLLATERAL.xlsx │  │ 1 per collateral│  │ CERSAI portal   │  │ and overrides    │
 └────────┬────────┘  └────────┬────────┘  └────────┬────────┘  └─────────┬────────┘
          │                    │                    │                     │
          ▼                    ▼                    ▼                     ▼
 ┌─────────────────────────────────────────────────────────────────────────────────┐
 │ 1. LOAD      read all sheets as text; check master version; log row counts      │
 │ 2. SCOPE     keep active loans only (Loan Status = A); drop already-filed SIs;  │
 │              1 SI candidate = 1 (Application number, Collateral ID)             │
 │ 3. LINK      match each title-report PDF to its collateral by file name;        │
 │              read only what the LMS lacks; check it is the right PDF (§15)      │
 │ 4. BUILD     borrowers per application (CKYC, de-duplicated per customer)       │
 │              owners per collateral (COLLATERAL, verified by title report) → TPM │
 │ 5. ENRICH    apply the values typed in the required-details file                │
 │ 6. TRANSFORM clean text, reformat dates, split addresses, map codes,            │
 │              resolve PIN → district → state → city                              │
 │ 7. VALIDATE  field rules (spec) + cross-field rules + file rules                │
 │              failures → left out and listed; missing data → required details    │
 │ 8. WRITE     .DAT (only fully valid SIs, ≤1000 per file)                        │
 │              working / final Excel, checks report, required-details file        │
 └─────────────────────────────────────────────────────────────────────────────────┘
           │
           ▼
 Ops reviews → uploads .DAT on the CERSAI portal → checks "Batch Report"
 → records CERSAI SI IDs in a "filed register" (so the next run skips them)
```

It runs as **three steps**, with people filling in and checking values between them:

- **Step 1 (build):** creates the batch folder, the working Excel (every value found, laid out like the demo) and the required-details file (what is missing, plus LMS / title report differences side by side).
- **Ops fills in** the missing values and answers the checks. **Step 2 (apply)** reads them and makes the final Excel; it can be run again as often as needed.
- **Step 3 (create the .DAT):** checks the final Excel and writes the `.DAT` file. SIs that are still incomplete are left out, listed, and kept pending, so they don't block the rest of the batch.

---

## 4. Output file structure

### 4.1 File-level rules (from the `introduction` sheet)

| Rule | Value |
|---|---|
| Format | Pipe-separated text file with the `.DAT` extension |
| File name | `crs_NNNNN_012039_YYYYMMDDHHMISS.DAT`. `NNNNN` is most likely the 5-character entity code (**`J8084`**); `012039` is given literally in the spec; the timestamp is the generation time in 24-hour format. **Confirm (§11 Q3).** |
| Dates | Always text in `YYYY-MM-DD` format. Never future-dated. Never before 1900. |
| Records per file | 1–1000 SI records. If there are more, split into several files. |
| Batch file version | 1 |
| Blank lines / helper text | **None.** The demo's spacer rows and `YYY-MM-DD` notes must not reach the file. |
| Empty fields | Still separated by pipes, so every line has exactly *(field count − 1)* pipes. |
| Duplicate uploads | The spec asks users to check the portal's *Batch Report* before re-uploading. The automation should keep a log of generated files and filed SIs to avoid this. |

### 4.2 Line order and field counts

The order follows the sheet order in the spec and matches the demo:

```
FH                       ← once per file
SI                       ← once per security interest (repeat this block per SI)
IMM                      ← the property
BOR  (× NUMBER_OF_BORROWERS)
ASO  (× NUMBER_OF_ASSET_OWNERS)   ← only for third-party mortgages; out of scope for now
CHG                      ← charge holder (ARGFL)
```

| Row type | Spec sheet | Fields | Excel columns in a review sheet (B = field 1) |
|---|---|---|---|
| FH | introduction | 4 | B–E |
| SI | si | 15 | B–P |
| IMM | immovable_asset | 27 | B–AB |
| BOR | borrower | 27 | B–AB |
| CHG | charge_holder | 11 | B–L |

### 4.3 What counts as one SI

**Recommendation: one SI record per (Application number, Collateral ID).** This is what the demo implies: its MIS token is `APPL…` + `COLL…`, and there is exactly one IMM line per SI. The SI row also has no "number of assets" field, which suggests one asset per SI.

- The current extract gives **1,312 SI candidates** from **1,150 active applications**. Of these applications, 1,034 have exactly 1 collateral and 116 have 2–8.
- **An application can have several collaterals, but each Collateral ID belongs to exactly one application** (checked: no Collateral ID appears under two applications, in any loan status). *(implementation)*
  - Each collateral of an application gets its own SI block. The borrowers (BOR rows) are the same in each, because they come from the application; each SI carries the full sanctioned amount of the loan (Q7; all 116 multi-collateral applications have one sanctioned amount).
  - The **owners are worked out per collateral**: across the 116 applications they are the same on every collateral in 46, completely different in 35, overlapping in 28, and missing on some collaterals in 7. So IS_ASSET_OWNER, NUMBER_OF_ASSET_OWNERS and IS_TPM can differ between the SIs of one application.
  - The SI blocks of one application are written one after the other.
  - If COLLATERAL ever listed a Collateral ID under two applications, it would be a data error: the collateral is left out of the batch and reported, rather than mixing two applications' rows.
  - Checked on the extract for all 116 applications: one active loan account, one sanctioned amount and one disbursal date each; the same customers (and the same customer addresses, phones and PANs) on every collateral; one value of every property field per collateral. The tool's output was checked against the extract for all 278 collaterals: the same BOR rows on every SI of an application apart from IS_ASSET_OWNER, owners taken only from their own collateral, and IS_ASSET_OWNER differing between the SIs of 54 applications.
  - **Two collaterals of one application with the same address** (1 application: `APPL00036867`, two agricultural-plot collaterals with the same address but different areas and Global collateral Nos) are flagged on both SIs, to check that the same property is not recorded twice. They are not held: they may be two parts of one plot.
  - The same property can still secure two applications under **different** Collateral IDs (24 properties share a `Global collateral No`; see TYPE_OF_CHARGE).
- The title reports follow the same grain: one PDF per collateral, named `<Application number>_<Collateral ID>.pdf` (§15.1).
- **Confirm (§11 Q5)** whether CERSAI accepts several IMM lines under one SI instead.

---

## 5. Source data model

### 5.1 `CKYC.xlsx` (borrowers)

| Column | Use | Notes |
|---|---|---|
| `Application Number` | Key | Joins to COLLATERAL `Application number`. The two files have identical application and customer sets. |
| `Neo CIF ID` | Customer key | Unique per customer. |
| `Loan Account No` | Reference | For in-scope applications, it is the active loan account. Null for the 23 applications that only have cancelled loan accounts (out of scope). CKYC has no loan status of its own, so filter it by the active application list from COLLATERAL. |
| `First Disbursal Date` | SI_CREATION_DATE | `DD-MM-YYYY` with a 4-digit year, so it is safe to parse. |
| `DOD` | Not needed | Same date in `DD-MON-YY` format. |
| `Customer Type` | Borrower selection | `primary` (exactly 1 per application), `Co-Applicant`, `Guarantor` (34). |
| `Applicant Type` | Individual vs non-individual | `individual` / `non_individual`. Cross-check against `CONSTITUTION`. |
| `Customer Name` | BORROWER_NAME / INDIVIDUAL_NAME | |
| `CONSTITUTION` | BORROWER_TYPE | 9 values. Mapping in §8.7. |
| `Date Of Birth` | DATE_OF_BIRTH | `DD-MON-YY`, **2-digit year** (§8.2). |
| `Incorporation Date` | BORROWER_REGISTRATION_DATE | `DD-MON-YY`, 2-digit year. |
| `First Name` / `Middle Name` / `Last Name` | Fallback for the individual name | Middle name is **not** the father's name. It contains values like `WO` ("wife of"), and `Last Name` is a single space for 138 customers. |
| `Identification Type` / `Identification Number` | PAN, CIN, CKYC | One row per document. Types: PAN 3,942 · GST_NO 369 · UDHYAM_REGISTRATION 352 · UDYAM 34 · Shop/Establishment 13 · Voter's ID 5 · PASSPORT 4 · Aadhar 3 · Unsupported 2 · **KYC Identification Number 1** · **CIN No 1**. |
| `Mobile No` | MOBILE_NUMBER | Always `+91--XXXXXXXXXX`. |
| `Email ID` | EMAIL | |
| `Full Address` | Borrower address | Always ends with `, <City>, <State>- <PIN>, India`. It has **no address type**, and 40 in-scope customers have more than one address. |
| Others (`Branch Name`, risk category, `Primary Mobile`, `Product ID`, `Application Received Date`) | Not needed | `Branch Name` only matters if the CHG row becomes branch-specific. |

### 5.2 `COLLATERAL.xlsx` (loans, properties, owners)

| Column | Use | Notes |
|---|---|---|
| `Application number` | Key | |
| `Loan Account #` | NARRATION (proposed) | 24 characters, which is too long for ENTITY_MIS_TOKEN. |
| `Loan Status` | **Scope filter: keep `A` only** | **A** = active (6,652 rows, kept), **X** = cancelled with disbursed amount always 0 (1,008 rows, dropped), **C** = closed (25 rows, dropped). After the filter, every collateral has exactly one loan account, and its sanctioned amount and disbursal date are the same on all its rows. |
| `Disbursal Date` | Alternative for SI_CREATION_DATE | This is the loan booking date (`DD-MON-YY`). On active rows it equals CKYC `First Disbursal Date` in 97.5% of applications (30 differ). |
| `Neo CIF ID`, `Customer Name`, `Customer Type` | Borrower join | |
| `Collateral ID` | SI key | Format `COLL` + 8 digits. |
| `Global collateral No` | ASSET_UNIQUE_ID (proposed) | Format `CMS` + 8 digits. **24 properties back 2 active applications**, which has implications for TYPE_OF_CHARGE and duplicate filing (§11 Q12). |
| `Ownership Status`, `Owner Name`, `Percentage Share` | Owner matching | Owner names often carry `Mr.`/`Mrs.`/`Smt.` prefixes and spelling variants. 37 collaterals have no owner rows at all. |
| `Property Type` | ASSET_TYPE_ID | Active counts: RESIDENTIAL 998 · COMMERCIAL 201 · MIXED USAGE 40 · INDUSTRIAL 37 · PLOT 35 · PG 1. |
| `Property Classification` | Sub-type hint only | About 99% are "Others", so it is not usable on its own. |
| `Full Address` | IMM address parsing | Free text of up to 501 characters. **1,194 of 1,312 are over 100 characters.** About 60% mention a survey, plot, CTS or khasra number, about 45% a floor, and about 70% a flat, unit, office or house number. |
| `Property Pincode` | IMM PIN_CODE | 6 digits. Matches the PIN in the address tail in 99.9% of rows. |
| `Build up Area`, `Carpert Area` | ASSET_AREA (fallback only) | Unit is not stated (probably sq ft). The demo value (600) matches neither column. |
| `Sanctioned Loan Amount` | TOTAL_SECURED_AMOUNT | The same on every active row of a collateral (checked: no exceptions), so any active row can be used. 263 active loan accounts are only partly disbursed. |
| `Disbursed Loan Amount` | Filter check | |
| `PHONE_NUMBER`, `EMAIL_ADDRESS` | Fallback contact | |
| `Identification Number` | Fallback PAN | Agrees with the CKYC PAN wherever both exist (3,719 in-scope customers). |
| `ADDRESS_TYPE`, `Customer Address` | **Preferred borrower-address source** | Same text as CKYC `Full Address`, but typed (Residential / Office / Permanent / AlternateBusiness), which lets us pick the right address. |
| Others (purpose, ownership, nature, FOIR, program, age, market value) | Not needed | |

### 5.3 CERSAI masters used (v29.3)

| Master sheet | Used for | Values relevant here |
|---|---|---|
| `SI_SUB_TYPE_ID` / `TRANSACTION_SUB_TYPE_ID` | SI_TYPE_ID | Immovable: **17** Equitable Mortgage · 18 Registered Mortgage · 19 Agreement to Create Mortgage (UUC) · 20 Conditional Sale · 21 Other |
| `FINANCING_TYPE_ID` | FINANCING_TYPE_ID | **1** Sole · 2 Consortium · 3 Multiple |
| `ASSET_CATEGORY_ID` | ASSET_CATEGORY_ID | **1** Immovable |
| `ASSET_TYPE_ID` | ASSET_TYPE_ID | 1 Residential · 2 Commercial · 3 Other |
| `ASSET_SUB_TYPE_ID` | ASSET_SUB_TYPE_ID | Residential: 1 Structure with Land / 2 Structure without Land / 3 Land without Structure / 20 Unidentified. Commercial: 4 / 5 / 6 / 21. Other: 7 / 8 / 19 / 22. **Each sub-type belongs to exactly one asset type.** |
| `AREA_UNIT` | ASSET_AREA_UNIT | 1 Acre · 2 Ares · 3 Cents · 4 Gunta · 5 Hectare · **6 Square Feet** · 7 Square Metre · 8 Square Yard · 9 None |
| `GENDER` | GENDER | M · F · T |
| `BORROWER_ASSET_OWNER_TYPE` | BORROWER_TYPE | AOP · BOI · COM (Indian Company) · FRC · GOV · HUF · IND · LLP · NTC · PRF (Partnership Firm) · SOC · SPR (Sole Proprietorship) · TRS |
| `PIN_CODE` | PIN → district | 19,482 distinct PINs. **1,271 PINs map to more than one district.** There is a trailing empty column to ignore. |
| `DISTRICT` | District → state | 764 districts, each with one state code. |
| `STATE` | State code | 36 states, each with country code `IND`. |
| `CITY` | City code | 45,899 cities, each tied to one district. **8,066 city names repeat across districts**, so never look up a city by name alone. |
| `COUNTRY` | COUNTRY | 3-letter codes only. India = **`IND`**. |
| `ENTITY_CODE` | File name / reference | `J8084` = Anand Rathi Global Finance Ltd |

### 5.4 Legal title reports (`Input data/Legal/`)

One advocate's title report per collateral, named `<Application number>_<Collateral ID>.pdf`. The LMS input data comes first; the title report is read only for what the LMS lacks.
- It is the only source for the **type of mortgage**.
- It is read for the **owners** of every collateral, to verify the COLLATERAL owner list (§6.3).
- It fills gaps such as survey/plot numbers (neither appears in 39% of LMS property addresses) and address details that can't be parsed.
- It has no amounts, no mortgage date and no borrower KYC data.

Structure, extractable fields and mapping are in **§15**.

---

## 6. Record assembly logic and dependencies

### 6.1 Which loans go into a batch

1. **Keep only active loans: COLLATERAL rows with `Loan Status = A`** (decided 2026-09-28).
   - This drops cancelled loan accounts (`X`, nothing disbursed) and closed loans (`C`). 27 applications drop out completely (23 with only cancelled accounts, 4 closed); 120 more lose a stale cancelled row.
   - After this filter, every collateral has exactly one loan account, and its sanctioned amount and disbursal date are identical on all its rows. Keep a check that flags any collateral where they ever differ.
   - CKYC has no loan status, so restrict the CKYC borrower list to the applications that remain after this filter.
2. **Exclude SIs already registered with CERSAI.** No register of filed SIs exists yet, so one must be created (§14.4). The date filter in the SQL is not enough on its own.
3. One SI candidate per remaining (Application number, Collateral ID). That gives **1,312** in this extract (from 1,150 applications).

### 6.2 Borrowers per SI (BOR rows)

- **Source:** CKYC rows for the application. De-duplicate to one row per `Neo CIF ID`, because CKYC repeats a customer once per ID document.
- **Include:** `primary` and `Co-Applicant`. **Exclude `Guarantor` by default** (§11 Q10), since a guarantor is not a borrower. If a guarantor owns the property, they become an asset owner, which makes the loan a third-party mortgage.
- **Order:** primary first, then co-applicants (for example by `Neo CIF ID`). SERIAL_NUMBER runs 1..n within the SI.
- **Per-customer attributes:**
  - PAN = the `Identification Number` of the row where `Identification Type = PAN`. There is at most one per customer.
  - CIN = the row with type `CIN No`.
  - CKYC = the row with type `KYC Identification Number`.
  - Everything else is identical across a customer's rows, except mobile (35 customers) and address (40 customers). The address rule is in §8.4.
- `SI.NUMBER_OF_BORROWERS` = the number of BOR rows written. It must never be typed in by hand; that is how the demo ended up with a mismatch.

### 6.3 Asset owners, IS_ASSET_OWNER and third-party mortgage (TPM)

- **Owners come from COLLATERAL first; the title report checks them** (decided 2026-09-30, revised the same day). Owner names read from a title report are less reliable than the COLLATERAL `Owner Name` column, so the title report is the main source only when COLLATERAL has no owners.
  - Read both lists:
    - the distinct `Owner Name` values for that `Collateral ID` in active COLLATERAL rows;
    - title report item 4 "Name of the owner" (§15.5). It lists all owners in one sentence (`Mr. A & Mrs. B`, `A, B and C`), so split it into names first (a `S/O ...` piece stays with the name before it).
  - **Double-check the COLLATERAL count** with `Percentage Share` (the owners' shares must add up to 100%, ±1) and `Ownership Status` (`Single` = one owner, `Joint` = two or more). In the current extract every collateral adds up, so this is a safety net: a problem is listed in Check and correct.
  - **One `Owner Name` cell can hold two people** (`Kalaiselvi and Radhakrishnan Sankaranarayanan`, 100%, Single; `KAMALA AND PUNAM` next to `PUNAM`: 3 collaterals). The cell is split into owners when each part is a different borrower and the whole name is not itself a borrower's name. Company names contain "and" too (`Jamson Machine and tools`, `OSNAR PAINTS AND CONTRACTS`, `RAMPAL PAREEK AND COMPANY`), so otherwise the cell stays whole. If it has no business word, it is listed as possibly two people (`Nikhil Bindroo and Vimi Bindroo`, a third-party mortgage).
  - If the report's header box or item 12 names different owners from item 4, flag the report.
  - Compare the two lists after the name normalisation below.

    | Situation | Owners used | Status |
    |---|---|---|
    | Both lists name the same owners | Either (they agree) | Verified |
    | The lists differ (names or number of owners) | COLLATERAL's | Both lists shown in Check and correct, where the owners can be typed (Q32) |
    | The COLLATERAL shares show an owner is missing (less than 100%, or `Joint` with one owner), and the title report names every COLLATERAL owner plus others | COLLATERAL's plus the title report's others | Shown in Check and correct |
    | COLLATERAL has no owner rows (37 collaterals) | Title report's | Shown in Check and correct |
    | The title report is held as possibly the wrong PDF (§15.1) | COLLATERAL's | Not verified until the PDF is released |
    | No usable title report | COLLATERAL's | Not verified |
    | Neither source has owners | Typed in the required-details file | Manual |
- **Name normalisation:** upper-case; remove honorifics (`MR`, `MRS`, `MS`, `SMT`, `SHRI`, `DR`, `M/S`, `KUM`); remove punctuation; collapse spaces; treat `PVT LTD` and `PRIVATE LIMITED` as the same, and `MOHD` / `MD` / `MOHAMMAD` as `MOHAMMED`. Drop a relative's name after `S/O`, `D/O`, `W/O` (CKYC writes `SO`, `DO`, `WO` in 20 names: `KAMALA DO HIMMATMAL` is compared as `KAMALA`, and `ASHOK SO BAJRANG LAL` is not taken for `BAJRANG LAL`). Two owners with the same name but different relatives stay two people.
- **Word-by-word comparison** *(implementation, replaces a plain fuzzy score)*. The words of the two names are paired:
  - the **same word**, also when written differently: sounding the same (`LAXMI` / `LAKSHMI`, `JAGJEET` / `JAGJIT`), with or without a `-BEN`, `-BHAI` or `-AMMA` ending (`LATABEN` / `LATA`), or as two words (`ANILKUMAR` / `ANIL KUMAR`, `VIRENDRASINGH` / `VIRENDRA SINGH`);
  - an **initial** for a word, on either side (`CHANDAN K SINGH` / `CHANDAN KAILASH SINGH`); title reports often write middle names as initials;
  - a **close spelling** (`PAREKH` / `PARIKH`, `NUTALA` / `NUTALAPATI`).

  Two names are **different people when each has a word the other lacks** — a different given name — or when only initials correspond. Relatives share the surname and often the father's or husband's name (`SUNANDA RAMCHANDRA SAWANT` / `RAMCHANDRA GANPAT SAWANT`, a wife `POOJABEN ASHWINBHAI CHAUHAN` / her husband `ASHWINBHAI JAGDISHBHAI CHAUHAN`). A plain fuzzy score rates such pairs as the same person, which would hide a third-party mortgage.
- **Given names must correspond.** When one name has extra words, the surname is found (the matched word at the end: last in `RAVIRAJ HIREMATH`, first in `SHAH RAGINI DHIRAJ`) and the given names must pair up. A wife's name usually contains her husband's given name and surname: `DEEPA RAVIRAJ HIREMATH` / `RAVIRAJ HIREMATH` has the given names DEEPA and RAVIRAJ, so they are relatives, not the same person. A missing middle name is fine (`ROHAN JOSHI` / `ROHAN MAHESH JOSHI`). A leading initial (`M GURU RAGHAVENDRA`, usually the father's) and a leading `MOHAMMED` are skipped when finding the given name; `SARDAR` is a title. When the order of the names cannot be told, the names are "similar" at most. *(implementation)*
- **One borrower per owner, per property.** The owners of one collateral are linked to the borrowers best match first, each borrower to one owner at most: owners listed separately are different people. So when husband and wife both own a flat and only one is a borrower, the other cannot take his or her place. *(implementation)*
- **Match each owner to the application's borrowers.** The tool decides every case itself; nothing is held for an owner answer:
  - **same** (all words alike, or at least two alike and the rest initials or extra words, with the given names corresponding) → the borrower;
  - **similar** (a close spelling, too few words alike, or the word order unclear) → taken to be the borrower;
  - **relative** (similar, but the given names differ, e.g. the owner `GOPALAKRISHNAN` and the borrower `KAVITHA GOPALAKRISHNAN`) → a different person, so a non-borrower owner, unless another owner already has that borrower;
  - anything else → a **non-borrower owner** (third-party mortgage).

  The similar and relative decisions are listed in Check and correct (field `OWNER_MATCH`, with the tool's decision shown as the value used). A `Y` or `N` typed there changes the decision *(implementation)*, and one answer covers every collateral of the application where that owner appears.
- **Current extract (active loans, LMS owner names):** 1,636 owner rows. A raw exact text match finds only 798. With the rules above: 1,450 are the same as a borrower, 149 are similar (taken to be the borrower), 3 are relatives, and 34 match no borrower; 142 collaterals have a decision listed for review, none is held. With the "two people in one cell" split, 37 collaterals are third-party mortgages; 37 collaterals have no owner data. The earlier rules linked relatives to each other: 15 owners were linked to a spouse, parent or child instead of themselves (for example `DIVYESHBHAI PATEL` to his wife `SAROJBEN DIVYESHBHAI PATEL` rather than to `DIVYESH NARANDAS PATEL`), which put IS_ASSET_OWNER on the wrong borrower, and in 11 husband-and-wife properties both owners were linked to the same borrower, which hid a third-party mortgage.
- **Outputs:**
  - `BOR.IS_ASSET_OWNER` = 1 if that borrower matched an owner of this collateral, else 0.
  - `SI.NUMBER_OF_ASSET_OWNERS` = the number of **non-borrower** owners. It is 0 when every owner is a borrower, as the spec says.
  - `SI.IS_TPM` = 1 if `NUMBER_OF_ASSET_OWNERS > 0`, else 0.
  - If `IS_TPM = 1`, ASO rows are required. **ASO is out of scope**, so for now those SIs go to exceptions or are filed manually (§11 Q13).
- **Dependency:** owners typed in the required-details file are only a manual override: when neither source has owners, when an owner matches no borrower, or when a reviewer corrects an owner difference listed in Check and correct.

### 6.4 Cross-row dependencies to enforce

| Dependency | Rule |
|---|---|
| SI ↔ BOR | `NUMBER_OF_BORROWERS` = the number of BOR lines in the block, and must be at least 1. |
| SI ↔ ASO | `NUMBER_OF_ASSET_OWNERS` = the number of ASO lines. `IS_TPM` = 1 exactly when it is above 0. |
| SI ↔ consortium | `FINANCING_TYPE_ID = 1` (Sole) means `NUMBER_OF_CONSORTIUM_MEMBERS = 0`. |
| SI ↔ IMM | `SI_TYPE_ID` must be an immovable type (17–21) when the asset category is 1. |
| IMM type ↔ sub-type | The sub-type must belong to the asset type (§5.3). |
| IMM type ↔ others | `ASSET_TYPE_OTHERS` is required when the type is 3, and blank otherwise. |
| IMM area ↔ unit | If `ASSET_AREA` is filled, `ASSET_AREA_UNIT` is required. |
| IMM sub-type ↔ house no. | `HOUSE_NUMBER` is required when the sub-type is "Structure with Land" or "Structure without Land" (1, 2, 4, 5, 7, 8). |
| IMM survey ↔ plot | At least one of `SURVEY_NUMBER` and `PLOT_NUMBER` must be filled. |
| BOR type ↔ field group | Non-individual types use fields 7–11. `IND` uses fields 12–19. The other group must be blank. |
| BOR PAN ↔ CKYC | At least one is needed (PAN or CKYC for non-individuals; individual PAN or individual CKYC for individuals). |
| BOR type ↔ registration no. | CIN or LLPIN is required when the type is COM or LLP. |
| Geography | PIN must be in the PIN master for the chosen district; the district must be in the state; the city must be in the district. |
| FH ↔ SI | `NUMBER_OF_RECORDS` = the number of SI blocks in the file, between 1 and 1000. |

---

## 7. Field-by-field mapping

Key: **M** = mandatory, **O** = optional, **C** = conditional. "Col" is the column in a review sheet where column B holds field 1, the same layout as the demo. "Restricted chars" means the spec's character rule for that field (§8.1).

### 7.1 FH: file header (`introduction` sheet)

| # | Col | Field | M/O/C | Source | Rule / default | Validation |
|---|---|---|---|---|---|---|
| 1 | B | FILE_HEADER | M | Constant | `FH` | |
| 2 | C | FILE_TYPE | M | Constant | `SI_REGN` | |
| 3 | D | NUMBER_OF_RECORDS | M | Derived | Number of SI blocks in this file (§11 Q2) | Between 1 and 1000 |
| 4 | E | FILE_DATE | M | System date | Generation date as `YYYY-MM-DD` | Not in the future |

### 7.2 SI: security interest (`si` sheet)

| # | Col | Field | M/O/C | Source | Rule / default | Validation |
|---|---|---|---|---|---|---|
| 1 | B | ROW_TYPE | M | Constant | `SI` | |
| 2 | C | SERIAL_NUMBER | M | Generated | 1..n across SI blocks in the file | Unique in the file |
| 3 | D | NUMBER_OF_BORROWERS | M | Derived (§6.2) | Count of BOR rows written | ≥ 1; equals the BOR line count |
| 4 | E | NUMBER_OF_ASSET_OWNERS | M | Derived (§6.3) from COLLATERAL `Owner Name` (count checked with `Percentage Share` and `Ownership Status`), checked against title report item 4 (the report's owners are used only when COLLATERAL has none); required details override | Number of non-borrower owners; 0 if all owners are borrowers | ≥ 0; equals the ASO line count |
| 5 | F | NUMBER_OF_CONSORTIUM_MEMBERS | M | Constant | `0` | |
| 6 | G | SI_TYPE_ID | M | **Title report item 8 "Nature of Mortgage"** (§15.5); the required-details file overrides | `Equitable Mortgage` → `17` (all 4 demo reports). Default **`17`** when there is no title report. | Must be in {17, 18, 19, 20, 21} |
| 7 | H | SI_TYPE_OTHERS | C | Title report item 8 wording; required details override | Only when SI_TYPE_ID = 21 | ≤ 100, restricted chars |
| 8 | I | FINANCING_TYPE_ID | M | Constant | `1` (Sole) | FINANCING_TYPE master |
| 9 | J | SI_CREATION_DATE | M | CKYC `First Disbursal Date` | `DD-MM-YYYY` → `YYYY-MM-DD` (demo record: `2026-07-01`) | Not in the future, ≥ 1900. See §11 Q6 on whether the mortgage-deed date should be used instead. |
| 10 | K | TOTAL_SECURED_AMOUNT | M | COLLATERAL `Sanctioned Loan Amount`, from any active row of the collateral (the value is the same on all of them) | Whole rupees, no commas; flag the collateral if its active rows disagree | > 0; at most 13 characters; at most 2 decimals. See §11 Q7. |
| 11 | L | ENTITY_MIS_TOKEN | M | Application number + Collateral ID | Compressed to fit 20 characters: `A` + 8 digits of the application + `C` + 8 digits of the collateral, e.g. `A00027326C00266688` (18 characters, reversible) | Alphanumeric only; ≤ 20. **The demo value is 24 characters.** |
| 12 | M | NARRATION | O | COLLATERAL `Loan Account #` | Proposed `LAN ARGFLLAPMUM2000005018390`, which keeps the loan account number that doesn't fit in the MIS token. Blank is also fine; avoid `NA`. | ≤ 100, restricted chars |
| 13 | N | TYPE_OF_CHARGE | M | Constant, with required details override | `FIRST CHARGE`. Override for properties already securing another loan (§11 Q12); the LMS finds these through a shared `Global collateral No`. | Letters and spaces only; ≤ 20 |
| 14 | O | IS_TPM | M | Derived (§6.3) | `1` if NUMBER_OF_ASSET_OWNERS > 0, else `0` | 0 or 1 |
| 15 | P | BATCH_REF_NUMBER | O | Generated | Blank, **or** exactly 20 alphanumeric characters, e.g. `ARGF` + file date `YYYYMMDD` + 8 digits of the collateral = `ARGF2026090100266688` | **Exactly 20 characters**, alphanumeric only |

### 7.3 IMM: immovable asset (`immovable_asset` sheet)

The LMS source listed first is used whenever it gives a usable value. "Title report item N" (the numbered row of the collateral's title report) is read only when the LMS value is blank, invalid or can't be parsed (§15.6); the extraction rules are in §15.4 and the value mapping in §15.5. the required-details file always overrides.

| # | Col | Field | M/O/C | Source | Rule / default | Validation |
|---|---|---|---|---|---|---|
| 1 | B | ROW_TYPE | M | Constant | `IMM` | |
| 2 | C | ASSET_CATEGORY_ID | M | Constant | `1` (Immovable) | |
| 3 | D | ASSET_TYPE_ID | M | COLLATERAL `Property Type` (never blank); required details override | Mapping in §8.7. For `PLOT` (35 collaterals), use title report item 2 "Nature of immovable property" to choose: `Residential` → 1, `Commercial` (also the typo `Commerical`) → 2. | 1, 2 or 3 |
| 4 | E | ASSET_TYPE_OTHERS | C | Derived from `Property Type`; required details override | Only when the type is 3, e.g. `Industrial`, `Mixed Usage` | ≤ 50, restricted chars |
| 5 | F | ASSET_SUB_TYPE_ID | M | Unit keywords in the LMS address (§8.7); else the unit word in title report item 1; required details override | Flat/shop/office/unit in a building → "Structure without Land" (2 / 5 / 8); independent house → "with Land" (1 / 4 / 7); plot → "Land without Structure" (3 / 6 / 19). e.g. an office in a building (Commercial) = `5` | Must belong to ASSET_TYPE_ID |
| 6 | G | ASSET_UNIQUE_ID | O | COLLATERAL `Global collateral No` (proposed) | e.g. `CMS00260877`, or blank. **The demo omits this column entirely.** | Rule is "**" (undefined); keep it alphanumeric |
| 7 | H | ASSET_DESCRIPTION | M | Built from the LMS parts (property type, unit, project); missing parts from title report item 1; required details override | `<type> <unit kind> <house no> <project>`, e.g. `COMMERCIAL UNIT 21 YASHODHAM BUSINESS CSL`; with no unit number, the plot number (`PLOT NO 45`); cut at a word boundary | ≤ 100, restricted chars |
| 8 | I | SURVEY_NUMBER | C | Regex on the property `Full Address`; **title report item 1 when the address has none** (39% of active addresses have neither a survey nor a plot number); required details override | Numbers after `Survey No`, `S No`, `Sy No`, `CTS No`, `C.S. No`, `Khasra/Kh No`, `Gat No`, `Khata/Khewat No`, `Mustatil/Killa`, joined with `^`. If old and new survey numbers are both given, use the new one plus its Hissa (`12/2`). | ≤ 1000; separator question in §11 Q15 |
| 9 | J | PLOT_NUMBER | C | Regex `Plot No` on the LMS address; else title report item 1 (`Plot No.`, or the `FP/PF No.` of a town-planning scheme); required details override | Leave blank rather than `NA` | ≤ 20; letters, digits, `- _ . /` and space |
| 10 | K | ASSET_AREA | O | COLLATERAL `Carpert Area` (blank for 10 active collaterals); else title report item 1 (`adm./admeasuring <number> <unit> <area type>`); required details override | LMS area assumed to be sq ft (§11 Q19). A title-report area is used as the report states it (§11 Q26). | Decimal(7,2), so at most 99,999.99. One collateral exceeds this. |
| 11 | L | ASSET_AREA_UNIT | C | Follows the area's source; required details override | `6` (Square Feet) with an LMS area. With a title-report area, the unit written next to it: `Sq. Mtrs.`/`sq.mtr` → `7`, `Sq. Ft.`/`Sq. Fts.` → `6`, `Sq. Yds.` → `8` (§15.5). | Required when ASSET_AREA is filled; AREA_UNIT master |
| 12 | M | HOUSE_NUMBER | C | Regex on the LMS address (`Flat/Office/Unit/Shop/House/H/Gala/Galla/Row House/Bungalow/Door No`, also `House bearing No.` and `Property (bearing) No.`; about 60% have one); else the unit number in title report item 1; required details override | e.g. `21`, `13B`, `C-3002` | Required for sub-types 1, 2, 4, 5, 7, 8; also asked for when the sub-type is not known yet (except PLOT), so both are filled together *(implementation)*. ≤ 20; letters, digits, `- /` and space only (**no dot**). |
| 13 | N | FLOOR_NUMBER | O | Regex `(\d+)(st\|nd\|rd\|th)? Floor` on the LMS address; else title report item 1; `Ground` → 0; `Basement`/`Lower Ground` → −1; required details override | e.g. `1` | Whole number from −3 to 100 |
| 14 | O | BUILDING_NAME | O | Regex hints on the LMS address (`Wing`, `Tower`, `Building`, `Bldg`); else title report item 1 (`Building No. N known as …`, `Tower-N (name)`); required details override | Blank if unknown (not `NA`) | ≤ 50, restricted chars |
| 15 | P | PROJECT_NAME | M | The LMS address segment after the house and floor segments; else **title report item 1** (`project known as …`, else `Society known as …`); required details override | e.g. `Yashodham Business CSL`, `Aaradhya Parkwood` | ≤ 50, restricted chars. For land or villages with no project, see §11. |
| 16 | Q | STREET_NAME | O | Regex on a segment containing `Road/Rd/Marg/Street/Lane/Path/Highway`; else title report item 1 when it names a road | e.g. `Aba Karmarkar Road`. **Never remove spaces** to make it fit (the demo did). | ≤ 50; letters, digits, `- /` and space only |
| 17 | R | POCKET | O | Not in source | Blank | |
| 18 | S | LOCALITY | M | LMS address segment(s) just before the city/state tail; else title report item 1 (a named area, or `Village <name>`); required details override | e.g. `Dindoshi Goregaon East Mumbai` | ≤ 50, restricted chars |
| 19 | T | CITY | M | CITY master code (§8.3) | e.g. `187602` (GOREGAON EAST) | Must be in DISTRICT |
| 20 | U | DISTRICT | M | PIN_CODE master, from `Property Pincode`. When the PIN maps to several districts, pick the one named in title report item 1 (`Taluka and District …`). | e.g. `10680` (MUMBAI SUBURBAN) | Must be in the PIN master |
| 21 | V | STATE | M | DISTRICT master | e.g. `21` (Maharashtra) | Cross-check against the state name in the address |
| 22 | W | PIN_CODE | M | COLLATERAL `Property Pincode`. Only if it is blank or not in the PIN master (8 active collaterals): the PIN in title report item 1, when there is one (1 of 4 demo files); else the required-details file. | 6 digits | Must be in the PIN_CODE master |
| 23 | X | COUNTRY | M | Constant | `IND`: "as per master" in the format description (§11 Q4) | COUNTRY master |
| 24–27 | Y–AB | LATITUDE_LONGITUDE_1…4 | O | Not in source | Blank, but **the 4 empty positions must still be written** | |

### 7.4 BOR: borrower (`borrower` sheet)

Fields 7–11 apply only to **non-individual** borrower types (everything except `IND`, including `SPR`). Fields 12–19 apply only to **`IND`**. The unused group stays blank.

| # | Col | Field | M/O/C | Source | Rule / default | Validation |
|---|---|---|---|---|---|---|
| 1 | B | ROW_TYPE | M | Constant | `BOR` | |
| 2 | C | SERIAL_NUMBER | M | Generated | 1..n within the SI, primary borrower first | Unique within the SI |
| 3 | D | BORROWER_TYPE | M | CKYC `CONSTITUTION`, cross-checked with `Applicant Type` | Mapping in §8.7, e.g. Partnership Firm → `PRF` | BORROWER_ASSET_OWNER_TYPE master |
| 4 | E | IS_ASSET_OWNER | M | Derived (§6.3): `1` if this borrower matches one of the collateral's owners (COLLATERAL `Owner Name`, checked against title report item 4); required details override | `1` or `0` | |
| 5 | F | BORROWER_UID_TYPE | O | — | Blank | **Must be empty** |
| 6 | G | BORROWER_UID_VALUE | O | — | Blank | **Must be empty** |
| 7 | H | BORROWER_PAN | C | CKYC PAN row, falling back to COLLATERAL `Identification Number` | Non-individual only | `[A-Z]{5}[0-9]{4}[A-Z]`. A 4th-character check against the type (§8.6) raises a warning. |
| 8 | I | BORROWER_CKYC_NUMBER | C | **Required details** (not in the extract) | Non-individual only; needed when there is no PAN (431 SPR) | Exactly 14 digits |
| 9 | J | BORROWER_NAME | C | CKYC `Customer Name`, cleaned | Non-individual only; `&` → `AND`; brackets removed | ≤ 100; letters, digits, `- _ .` and space |
| 10 | K | BORROWER_REGISTRATION_DATE | C | CKYC `Incorporation Date` | `DD-MON-YY` → `YYYY-MM-DD` using the century rule (§8.2) | Not in the future, ≥ 1900 |
| 11 | L | BORROWER_REGISTRATION_NUMBER | C | CKYC `CIN No` row, else the required-details file | **Required for COM and LLP**; blank for other types | CIN is 21 characters in a fixed pattern; LLPIN per spec (§11 Q16) |
| 12 | M | INDIVIDUAL_PAN | C | CKYC PAN row | Individual only | PAN regex; 4th character `P` |
| 13 | N | INDIVIDUAL_CKYC_NUMBER | C | **Required details** | Individual only; needed when there is no PAN (5 cases) | Exactly 14 digits |
| 14 | O | GENDER | C | CKYC `Gender` (`Male` → `M`, `Female` → `F`); required details when blank or the column is missing | `M`, `F` or `T` | GENDER master. **Don't guess from names.** |
| 15 | P | INDIVIDUAL_NAME | C | CKYC `Customer Name`, falling back to First + Middle + Last | Remove dots and digits (e.g. `SARVES .` → `SARVES`) | Letters and spaces only; ≤ 50 |
| 16 | Q | FATHER_MOTHER_NAME | C | **Required details** (not in the extract) | Don't build it from the middle name (the demo did) | Letters and spaces only; ≤ 50 |
| 17 | R | DATE_OF_BIRTH | C | CKYC `Date Of Birth` | `DD-MON-YY` → `YYYY-MM-DD` using the century rule (§8.2) | Not in the future, ≥ 1900 |
| 18 | S | MOBILE_NUMBER | O | CKYC `Mobile No` with `+91--` removed (fallback: COLLATERAL `PHONE_NUMBER`) | Individual only | Exactly 10 digits |
| 19 | T | EMAIL | O | CKYC `Email ID`, trimmed | Individual only; leave blank if invalid | Letters, digits, `_ .`, one `@`, a dot after the `@`; ≤ 50. **Hyphens are not allowed.** |
| 20 | U | ADDRESS_LINE_1 | M | Borrower address (§8.4) | Tail removed, cleaned, split | ≤ 100, restricted chars |
| 21 | V | ADDRESS_LINE_2 | O | Overflow from line 1 | | ≤ 100 |
| 22 | W | ADDRESS_LINE_3 | O | Overflow from line 2 | | ≤ 100 |
| 23 | X | CITY | M | CITY master code (§8.3) using the borrower's PIN | | Must be in DISTRICT |
| 24 | Y | DISTRICT | M | PIN_CODE master | | |
| 25 | Z | STATE | M | DISTRICT master | | |
| 26 | AA | PIN_CODE | M | 6-digit PIN before `, India` in the address | | Must be in the PIN master |
| 27 | AB | COUNTRY | M | Constant | `IN`, as the format description gives it (§11 Q4) | |

### 7.5 CHG: charge holder (`charge_holder` sheet)

This is the same line for every SI. It comes from configuration, not from the data. **Confirmed 2026-09-30:** the CHG row stays fixed and is not mapped from any source, including the title reports. The values are set in `cersai_tool/rows/chg_row.py`.

| # | Col | Field | M/O/C | Source | Value | Validation |
|---|---|---|---|---|---|---|
| 1 | B | ROW_TYPE | M | Constant | `CHG` | |
| 2 | C | OFFICE_CODE | O | — | Blank | **Must be empty** |
| 3 | D | OFFICE_NAME | M | Config | `ANAND RATHI GLOBAL FINANCE LIMITED` (§11 Q20) | Letters and spaces only; ≤ 100. Branch names like `Mumbai 2` contain a digit and are **not allowed**. |
| 4 | E | ADDRESS_LINE_1 | M | Config | `Express Zone B Wing 5th Floor Western Express Highway Opposite Oberoi Mall Malad East Mumbai 400063` (99 characters) | ≤ 100 |
| 5 | F | ADDRESS_LINE_2 | O | Config | Blank | ≤ 100 |
| 6 | G | ADDRESS_LINE_3 | O | Config | Blank | ≤ 100 |
| 7 | H | CITY | M | Config | **`81133`** (MALAD, district 10680). The demo's `62995` MALAD EAST belongs to district 355. | Must be in the district |
| 8 | I | DISTRICT | M | Config | `10680` (MUMBAI SUBURBAN) | |
| 9 | J | STATE | M | Config | `21` (Maharashtra) | |
| 10 | K | PIN_CODE | M | Config | `400063` | |
| 11 | L | COUNTRY | M | Constant | `IN`, a constant in the format description (§11 Q4) | |

---

## 8. Shared transformation rules

### 8.1 Text cleaning (address lines, description, project, street, locality, names)

The spec's recurring rule is: only certain special characters are allowed (per field), **no two special characters in a row**, and no special character at the start or end. Whitespace is listed as one of the allowed special characters.

1. Trim. Replace tabs, newlines and non-breaking spaces (`\xa0` appears in the data) with a space.
2. In names, replace `&` with ` AND `.
3. Replace every character not allowed for the target field with a space. This covers commas, brackets, `#`, `:`, `'` and `"`.
4. **Remove dots** in addresses and names (`No.` → `No`, `Pvt. Ltd.` → `Pvt Ltd`). `. ` is a dot followed by a space, which is two special characters in a row. Keep dots only inside numbers where the field allows them.
5. Collapse repeated whitespace to a single space. Tighten ` / ` → `/` and ` - ` → `-`.
6. Strip any special characters at the start or end.
7. Enforce the maximum length by cutting at a **word boundary** (never mid-word, and never by deleting spaces).

The current data needs this: 60 non-individual names contain `&`, brackets or `. `; 20 individual names contain dots; and every address contains commas.

### 8.2 Dates

| Input | Format | Handling |
|---|---|---|
| CKYC `First Disbursal Date` | `DD-MM-YYYY` | Safe: reformat to `YYYY-MM-DD`. |
| `Date Of Birth`, `Incorporation Date`, `DOD`, COLLATERAL `Disbursal Date` | `DD-MON-YY` | **The century is ambiguous.** Python's `%y` reads 00–68 as 2000–2068, which would turn **540 birth dates** (YY 50–68) into future dates. Excel uses a different cut-off (00–29). |

Rules:

- **Date of birth:** year = `20YY` if that is no later than (current year − 18); otherwise `19YY`.
- **Incorporation date:** year = `20YY` if that is not in the future; otherwise `19YY`.
- **Better fix:** change both SQL queries to output `to_char(..., 'YYYY-MM-DD')` (§9.4).
- **Output:** always the *text* `YYYY-MM-DD`, never an Excel date cell. The demo's date cells turned into serial numbers (`38169`, `18639`, `22463`) in the concatenated line.

### 8.3 Geography: PIN → district → state → city

1. **Get the PIN.** For the property, use `Property Pincode`. For a borrower, use the 6 digits in the address tail (`…<State>- NNNNNN, India`).
2. **PIN → district code(s)** from the `PIN_CODE` master.
   - No match: exception. This affects 8 property PINs (including 1 blank) and about 30 borrower addresses.
   - More than one match: choose the district whose name, or whose state's name, appears in the address. Otherwise raise an exception. This affects 89 properties and about 300 borrower addresses.
3. **District → state code** from the `DISTRICT` master. Cross-check against the state name in the address tail.
4. **City code:** the candidates are the CITY rows whose `District Code` is that district.
   - Normalise both the candidate names and the address: upper-case, remove punctuation, expand `(E)`/`(W)` to `EAST`/`WEST`, and so on.
   - Choose the **longest** candidate name that appears as a whole phrase in the address. The most specific name wins: `GOREGAON EAST` beats `MUMBAI`.
   - The general city named in the address tail (`…, Mumbai, Maharashtra- …`) counts only when no more specific city of the district is named: `MALAD` beats `MUMBAI` for `Malad East, Mumbai`. A candidate named like the **state** never counts: the CITY master has a city called `UTTAR PRADESH` in Ghaziabad district, which every UP address tail would otherwise match. *(implementation)*
   - Fallbacks, in order: the LMS city token from the address tail, if it exists in that district; then the city named the same as the district; then it is asked for in the required-details file, where the city (or district) can be typed by name or code.
   - Today a plain whole-phrase match gives one hit for 296 properties, several hits for 999 (resolved by the longest-match rule) and none for 17.
5. Write the **codes**, not the names.
6. **Never pick a city by name alone:** the city must belong to the district. The demo breaks this twice: `60744 ANDHERI WEST` and `62995 MALAD EAST` both belong to district 355 (Mumbai), while the PINs map to 10680 (Mumbai Suburban). The correct codes are `186829 ANDHERI W` and `81133 MALAD`.

### 8.4 Borrower address selection and splitting

1. **Selection.** Use COLLATERAL `Customer Address` with its `ADDRESS_TYPE`, because the CKYC address has no type.
   - Individuals: ResidentialAddress, then PermanentAddress, then OfficeAddress.
   - Non-individuals: OfficeAddress, then AlternateBusinessAddress, then ResidentialAddress.
   - If nothing is found there, fall back to CKYC `Full Address`.
2. **Remove the tail** `, <City>, <State>- <PIN>, India`, since city, state and PIN have their own fields.
3. **Clean** as in §8.1.
4. **Split.** If the result is 100 characters or less, it all goes in line 1. Otherwise, fill lines 1–3 greedily with whole comma segments, each line at most 100 characters. Split a single segment at a word boundary only if it is over 100 on its own.
   - After tail removal, **245 in-scope borrower addresses are 101–200 characters**. None exceeds 200.
   - The demo's single-line address for Vijay Bhosale is **102 characters**, which is too long.

### 8.5 Property address parsing (IMM)

**This is the first source for these fields** (decided 2026-09-30: LMS input data first). When the LMS address doesn't give a value, read it from the collateral's title report, item 1 (§15.5). For example, 39% of active addresses have neither a survey nor a plot number.

The property `Full Address` is free text, and 1,194 of 1,312 exceed 100 characters. Parse it with keyword regexes on comma segments:

- `Flat/Office/Unit/Shop/House No …` → HOUSE_NUMBER
- `Nth Floor` → FLOOR_NUMBER
- `… Wing/Tower/Building` → BUILDING_NAME
- `CTS/Survey/Sy/Khasra/Gat No …` → SURVEY_NUMBER
- `Plot No …` → PLOT_NUMBER
- `… Road/Marg/Street` → STREET_NAME
- the segment before the city tail → LOCALITY
- the first named segment → PROJECT_NAME candidate

Refinements from testing on the full extract *(implementation)*:

- `House bearing No.`, `Property No.` and `Property bearing No.` also give the unit (house) number. A number marked "Old" (`Old Property No. 121, New Property No. 1777/…`) is used only when there is no other.
- `P.No.` is a plot number (Rajasthan-style addresses).
- Numbers are read only before a description of the boundaries (`bounded on the East by Road, on the West by Plot No. 16 …`): those numbers belong to the neighbouring plots.
- An ordinal such as `11th` never continues a survey or plot list (`Sy No. 1, 11th Division` → `1`).
- An address naming several floors (`Ground and First Floor`) gives no floor number.
- Segments that are administrative units (`Taluk`, `Tehsil`, `Hobli`, `Mandal`, `Division`, `District`, `Ward`, `Circle`) are never the locality.
- A value cut to fit its length never ends with a joining word such as `of` or `near`.

Coverage on the 1,312 active addresses: unit number 60%, locality 88%, survey or plot number 64%, floor 42%, project 39%, street 28%, building 26%.

A value counts as found only when its pattern matches cleanly; otherwise treat it as missing and go to the title report. Show the parsed values in the review sheet, with the address they came from: address formats vary a lot, so spot checks are worthwhile.

### 8.6 Identifiers and contacts

| Item | Rule |
|---|---|
| PAN | `^[A-Z]{5}[0-9]{4}[A-Z]$` (all 3,942 PANs in the data pass). The 4th character should match the type: P for IND/SPR, F for PRF/LLP, C for COM, H for HUF, T for TRS. A mismatch is a **warning** (6 today: e.g. 2 "Private Limited Company" borrowers have P-type PANs, 1 has L, and 2 sole proprietorships have F-type PANs). |
| CKYC (KIN) | Exactly 14 digits. Not present in the extracts. |
| CIN | 21 characters: letter, 5 digits, 2 letters, 4 digits, 3 letters, 6 digits (e.g. `U67190MH2011PTC213188`). |
| Mobile | Remove `+91--`; keep only if exactly 10 digits. |
| Email | Trim. Keep only if it matches the CERSAI rule; otherwise leave blank and log it. 5 fail today: a leading space, and hyphens such as `sriramanjaneya-traders@…`. |

### 8.7 Code mappings

**CONSTITUTION → BORROWER_TYPE** (borrower counts for active loans):

| LMS `CONSTITUTION` | CERSAI code | Count | Notes |
|---|---|---|---|
| Individual | `IND` | 3,137 | 1 row is flagged `non_individual`: send to exceptions |
| NO PAN | `IND` | 4 | All are individuals with no PAN, so they need a CKYC number |
| Sole Proprietorship Firm | `SPR` | 774 | Non-individual type, usually with the proprietor's personal (P) PAN; 25 flagged `individual` |
| Partnership Firm | `PRF` | 97 | The demo wrongly used `COM` |
| Limited Liability Partnership | `LLP` | 17 | LLPIN required |
| Private Limited Company | `COM` | 122 | CIN required; 2 flagged `individual` |
| Public Limited Company | `COM` | 4 | CIN required |
| Hindu Undivided Company *(sic)* | `HUF` | 2 | |
| Private Trust | `TRS` | 1 | |
| *anything new* | — | — | Exception. Never default to `NTC` silently. |

**Property Type → ASSET_TYPE_ID / sub-type.** This is a *proposed default* that the business must confirm (§11 Q21); the required-details file always overrides it.

| LMS `Property Type` | ASSET_TYPE_ID | ASSET_TYPE_OTHERS | Default sub-type rule |
|---|---|---|---|
| RESIDENTIAL (998) | 1 | — | Flat/apartment → 2 (without land); independent/row house/bungalow → 1 (with land) |
| COMMERCIAL (201) | 2 | — | Office/shop/unit/gala in a building → 5; independent building → 4 |
| INDUSTRIAL (37) | 3 | `Industrial` | Shed/building → 7; unit in an estate → 8; open land → 19 |
| MIXED USAGE (40) | 3 | `Mixed Usage` | 7 or 8 |
| PLOT (35) | Depends on use: 1, 2 or 3 | If 3: `Plot` | Land without structure: 3, 6 or 19 |
| PG (1) | 1 | — | 1 or 2 |

The sub-type keywords come from the address. "Structure without Land" fits units in multi-storey buildings; "Structure with Land" fits independent structures; "Land without Structure" fits plots.

---

## 9. Manual inputs: the required-details file

### 9.1 What the earlier draft (`Req details.xlsx`, now deleted) contained and where each item went

| Draft entry | Maps to | Level | Automatic source (LMS first, then title report; §15.6) | Comment |
|---|---|---|---|---|
| Count of property owner/owner name | SI.NUMBER_OF_ASSET_OWNERS, SI.IS_TPM, BOR.IS_ASSET_OWNER (and future ASO rows) | Per collateral | COLLATERAL `Owner Name` (count checked with `Percentage Share` and `Ownership Status`), checked against title report item 4 (the report's owners are used only when COLLATERAL has none) | Manual only when neither source has owners, an owner matches no borrower, or a reviewer corrects an owner difference listed in Check and correct (§6.3). |
| security interest type | SI.SI_TYPE_ID (+ SI_TYPE_OTHERS) | Per SI | Title report item 8 (the LMS doesn't hold it) | Default 17 when there is no title report. |
| Asset type | IMM.ASSET_TYPE_ID (+ ASSET_TYPE_OTHERS) | Per collateral | COLLATERAL `Property Type`; title report item 2 only to tell residential from commercial `PLOT`s | |
| Asset Description | IMM.ASSET_DESCRIPTION | Per collateral | **Built** from LMS parts; gaps from title report item 1 | Manual only to improve the wording. |
| Survey/Plot No | IMM.SURVEY_NUMBER / IMM.PLOT_NUMBER | Per collateral | LMS address; title report item 1 when the address has neither (39% of active addresses) | Keep them as **two columns**: they are different fields with different rules. |
| Asset Area | IMM.ASSET_AREA (+ ASSET_AREA_UNIT) | Per collateral | COLLATERAL `Carpert Area` (sq ft, to confirm); title report item 1, with its unit, when blank | Add an "Asset Area Unit" column for manual overrides (the unit is required whenever an area is given). |
| Registration no | Probably BOR.BORROWER_REGISTRATION_NUMBER (the demo's BOR placeholder says "Registration Number") | **Per borrower**, not per SI | Only the **property deed's** registration number (title report item 5), not a borrower's CIN/LLPIN | If you meant the deed number, it can be read automatically; **confirm (§11 Q14).** |
| asset_sub_type | IMM.ASSET_SUB_TYPE_ID | Per collateral | **Derived** from unit keywords in the LMS address, else the title report's item 1 wording (flat/shop in a building → without land) | Confirm the rule (§11 Q27). |

### 9.2 Missing from the draft (cannot be automated from the current extracts)

| Needed value | Field | Level | Current workload |
|---|---|---|---|
| **Gender** | BOR.GENDER | Per individual borrower | **3,141** (every individual) |
| **Father's / mother's name** | BOR.FATHER_MOTHER_NAME | Per individual borrower | **3,141** |
| CKYC number | BOR.BORROWER_CKYC_NUMBER / INDIVIDUAL_CKYC_NUMBER | Per borrower without a PAN | 436 (431 SPR, 4 "NO PAN", 1 individual) |
| CIN / LLPIN | BOR.BORROWER_REGISTRATION_NUMBER | Per COM/LLP borrower | 143 (only 1 CIN exists in the data) |
| Project name | IMM.PROJECT_NAME | Per collateral | LMS address first; else title report item 1 (all 4 demo reports name one). Manual only where both fail. |
| House number | IMM.HOUSE_NUMBER | Per collateral | LMS address (about 72% have a unit number); else title report item 1. Manual only where both fail. |
| Locality | IMM.LOCALITY | Per collateral | LMS address, else the village or area named in the title report. Manual only where both fail. |
| Floor, building, street | IMM optional fields | Per collateral | LMS address, else title report item 1. Optional. |
| Area unit | IMM.ASSET_AREA_UNIT | Per collateral | `6` with an LMS area; the title report's own unit with a title-report area. Manual only with a manual area. |
| Type of charge | SI.TYPE_OF_CHARGE | Per SI | Only for exceptions (24 properties back 2 active loans) |
| City override | IMM/BOR.CITY | Per collateral / borrower | About 17 properties plus borrower addresses with no city match |

### 9.3 The required-details file (as built)

Step 1 generates `Output data/<batch>/2 Required details_<batch>.xlsx` for each batch; step 2 reads it back. It lists only what is missing or needs checking, so a day's file is short.

**Principle:** every row carries its key. The value the tool found is shown in the input cell (grey) and a missing value is an empty yellow cell. When the file is read, a value counts as typed only when it differs from what the tool finds on its own, so cells left as they are change nothing. Values typed for the same collateral or borrower in an earlier batch are carried into the next batch's file (shown in blue), so nobody types them twice.

| Sheet | One row per | Key | Input columns |
|---|---|---|---|
| **Borrowers** | borrower with a value to fill (most individuals, because father's/mother's name is never in the sources; gender only when CKYC leaves it blank) | Application number + Neo CIF ID | BORROWER_TYPE, PAN, CKYC, REG_NUMBER (CIN/LLPIN), REG_DATE, GENDER, FATHER_MOTHER_NAME, DATE_OF_BIRTH, ADDRESS, PIN_CODE, CITY, DISTRICT |
| **Collaterals** | collateral with a value to fill | Collateral ID | SI_TYPE_ID, SI_TYPE_OTHERS, TYPE_OF_CHARGE, ASSET_TYPE_ID, ASSET_TYPE_OTHERS, ASSET_SUB_TYPE_ID, ASSET_DESCRIPTION, SURVEY_NUMBER, PLOT_NUMBER, ASSET_AREA, ASSET_AREA_UNIT, HOUSE_NUMBER, FLOOR_NUMBER, BUILDING_NAME, PROJECT_NAME, STREET_NAME, LOCALITY, PIN_CODE, CITY, DISTRICT, OWNERS |
| **Check and correct** | item to look at | Collateral ID (+ Neo CIF ID for a borrower) | Your value |
| Lists (hidden) | — | — | the drop-down choices, from the masters |

- A value typed in **Borrowers** applies to every SI of that application.
- **Check and correct** shows, side by side: LMS value, title-report value, value the tool will use. It lists LMS / title report differences that matter (owners, unit number, floor, survey and plot numbers, area beyond 10%, property type, PIN, district), held title reports (answer `PDF OK` / `WRONG PDF`), the owner decisions the tool made on its own (owner lists that differ, owners taken or added from the title report, COLLATERAL shares that do not add up, an Owner Name cell taken as two people, and owner names that only resemble a borrower's, where an optional Y / N changes the decision), third-party mortgages, and needed values that have no column elsewhere. About 30 empty rows at the bottom let a reviewer correct any other field (Collateral ID, row, Neo CIF ID for BOR rows, field, value). Values here are applied last, so they win.
- Code fields offer drop-downs showing `code - name`; city and district can be typed as a name or a code; dates as `YYYY-MM-DD` or `DD-MM-YYYY`; owners separated by `;`. Typing `BLANK` leaves an optional value empty.
- Colours: yellow = needed, amber = default used, grey = found, orange = LMS and title report disagree, blue = typed earlier, dark grey = does not apply.
- Settings that were planned for a `Config` sheet (entity code, file-name segment, charge-holder row, defaults) live in `cersai_tool/settings.py` and `cersai_tool/rows/chg_row.py`.

### 9.4 Reducing manual work permanently (LMS SQL changes)

Most of the per-borrower manual work disappears if the two queries are extended. The source tables are already joined in the queries (`Customer`, `IDENTIFICATION_DETAILS_REPORT_VIEW`, `Address Details`, `Property Details`, `Property Ownership Details`). Ask the LMS team to add:

- **Gender** and **father's/mother's name** from the Customer table, if captured at onboarding.
- **CKYC number (KIN)**, and **CIN/LLPIN**, as separate columns. Pivot the identification view into one column per document type, so each customer has one row.
- **Address type** in the CKYC extract, or pick the correct address in SQL.
- **4-digit years** (`YYYY-MM-DD`) for every date.
- The **owner's Neo CIF ID** in `Property Ownership Details`, if it exists. That would replace fuzzy name matching with an exact key.
- Any **structured property fields** the LMS stores separately (survey/plot number, flat number, floor, project, locality, area unit).
- The **MODT / title-deed deposit date**, if that is the legally correct SI creation date.

---

## 10. Review of the demo output (`DAT FILE WORKING 1.xlsx`)

The demo covers one SI: application `APPL00027326`, collateral `COLL00266688`, borrowers Samcon Builders (primary, partnership firm), Vijay Dattajirao Bhosale and Rajaram R Masane (co-applicants). Rajaram R Masane is the sole owner of the property.

**What it gets right:**
- SI field positions;
- the use of master codes for city, district and state;
- `SI_TYPE_ID 17`, `FINANCING_TYPE_ID 1` and `TYPE_OF_CHARGE`;
- `TOTAL_SECURED_AMOUNT` = sanctioned amount;
- `IS_TPM = 0` and `NUMBER_OF_ASSET_OWNERS = 0` (the owner is a borrower);
- the IS_ASSET_OWNER flags;
- the use of CTS number and unit number from the address.

### 10.1 Issues found

| # | Line | Field / cell | Demo value | Problem | Correct handling |
|---|---|---|---|---|---|
| 1 | FH | FILE_DATE (E1) | `01-09-2026` | Wrong format | `2026-09-01` |
| 2 | FH | Line end | `…\|01-09-2026\|` | Trailing pipe on FH only; the other lines have none | Use one convention everywhere (§11 Q1) |
| 3 | SI | NUMBER_OF_BORROWERS (D3) | `2` | There are **3** BOR lines | `3` (derive it, don't type it) |
| 4 | SI | SI_CREATION_DATE (J3) | `07-01-2026` | Wrong format **and** day/month swapped; the source is `01-07-2026` | `2026-07-01` |
| 5 | SI | ENTITY_MIS_TOKEN (L3) | `APPL00027326COLL00266688` | 24 characters; the maximum is 20 | e.g. `A00027326C00266688` |
| 6 | SI | NARRATION (M3) | `NA` | Valid but meaningless | Blank, or `LAN <loan account>` |
| 7 | SI | BATCH_REF_NUMBER (P3) | `BATCH_2026-01-09` | Contains `_` and `-` (not allowed); 16 characters, but exactly 20 are required; the date also looks swapped | Blank, or 20 alphanumeric characters, e.g. `ARGF2026090100266688` |
| 8 | IMM | ASSET_UNIQUE_ID | *(missing column)* | **Every later IMM field is shifted one position left.** CERSAI would read the description as the unique ID, the survey number as the description, and so on. | Emit all 27 positions |
| 9 | IMM | HOUSE/FLOOR vs AREA/UNIT order | Unit and floor placed before area and unit | Order differs from the spec (area is field 10, unit 11, house 12, floor 13). Combined with #8, CERSAI would read area = `1` and area unit = `600`. | Follow the spec order |
| 10 | IMM | LATITUDE_LONGITUDE_1–4 | *(missing)* | The line ends early: 22 fields instead of 27 | Emit 4 empty positions |
| 11 | IMM | Cell comments D5/F5 | "Immovable Category" / "Asset Type" | The labels are wrong: D is ASSET_TYPE_ID and F is ASSET_SUB_TYPE_ID | — |
| 12 | IMM | ASSET_SUB_TYPE_ID (F5) | `2` | Sub-type 2 belongs to *Residential*; it is invalid with ASSET_TYPE_ID 2 (Commercial) | `5` (Commercial, structure without land) |
| 13 | IMM | STREET_NAME (P5) | `YashodhamNrRoyalChallengeRestaurantAbaKarmarkarRd` | Spaces removed to fit 50 characters | `Aba Karmarkar Road`, cut at a word boundary |
| 14 | IMM | PLOT_NUMBER / BUILDING_NAME / POCKET | `NA` | `NA` would be read as an actual plot number or building name | Blank |
| 15 | IMM | ASSET_AREA (L5) | `600` | Matches neither LMS built-up (650) nor carpet (520) area; the source is undocumented | Take it from the required-details file and record the source |
| 16 | IMM | ASSET_DESCRIPTION (G5) | `COMMERCIAL UNIT` | Valid but minimal | Optionally add the unit and project |
| 17 | BOR 1 | BORROWER_TYPE (D7) | `COM` | Samcon Builders is a **Partnership Firm**; `COM` means Indian Company | `PRF` |
| 18 | BOR 1 | BORROWER_CKYC_NUMBER (I7) | `CKYC NO` | Placeholder text in a digits-only field | Blank (PAN is present) |
| 19 | BOR 1 | BORROWER_REGISTRATION_NUMBER (L7) | `Registration Number` | Placeholder text; not required for PRF | Blank |
| 20 | BOR 1 | BORROWER_REGISTRATION_DATE (K7) | Excel date, which becomes `38169` in the line | Date cell turned into a serial number in the concatenation | `2004-07-01` as text |
| 21 | BOR 2/3 | DATE_OF_BIRTH (R9, R11) | Excel dates, which become `18639` / `22463` | Same serial-number problem | Text `YYYY-MM-DD` |
| 22 | BOR 3 | DATE_OF_BIRTH (R11) | 1961-07-01 | The source says `01-JUN-61` | `1961-06-01` |
| 23 | BOR 2/3 | EMAIL (T9, T11) | `Email` | Placeholder text | `samconbuilders@gmail.com`, or blank |
| 24 | BOR all | ADDRESS_LINE_2 / ADDRESS_LINE_3 | *(missing)* | **CITY, DISTRICT, STATE, PIN and COUNTRY are shifted two positions left.** Each line has 25 fields instead of 27. | Emit all 27 positions |
| 25 | BOR 2 | ADDRESS_LINE_1 (U9) | 102 characters | Over the 100-character maximum | Split across lines 1–2 |
| 26 | BOR 2 | CITY (V9) | `60744` ANDHERI WEST | That city is in district 355 (Mumbai); PIN 400058 maps to district 10680 | `186829` ANDHERI W |
| 27 | BOR 2/3 | GENDER (O9, O11) | `M` | Not in any source file; origin unknown | From the required-details file or the LMS |
| 28 | BOR 2/3 | FATHER_MOTHER_NAME (Q9, Q11) | `Dattajirao Bhosale`, `R MASANE` | Built from middle + last name, which is a guess (a middle name can be "WO" = wife of) | From the required-details file or the LMS |
| 29 | BOR | Cell comments U7/U9/U11 | "ADDRESS MIN 100 CHARACTER" | It is a **maximum** of 100 | — |
| 30 | CHG | CITY…COUNTRY (N13–R13) | In columns N–R | They should be in H–L (after ADDRESS_LINE_2/3) | Fix positions |
| 31 | CHG | Concatenation formula (A13) | Only `B13…E13` | **The CHG line in the file has 4 fields instead of 11**, so every mandatory location field is missing | Emit all 11 |
| 32 | CHG | CITY (N13) | `62995` MALAD EAST | In district 355; PIN 400063 maps to 10680 | `81133` MALAD |
| 33 | CHG | OFFICE_NAME (D13) | `ANANDRATHI GLOBAL FINANCE LIMITED` | Cosmetic: the master spells it "Anand Rathi" | `ANAND RATHI GLOBAL FINANCE LIMITED` |
| 34 | All | Spacer rows and helper text | Blank rows 2/4/6…, `YYY-MM-DD` in F1/J2/R8/R10 | Must not end up in the `.DAT` | Write lines directly, without blanks |
| 35 | BOR, CHG | COUNTRY | `IND` | The format description gives `"IN"` for BOR and CHG; only the IMM row uses the COUNTRY master (`IND`) | `IN` for BOR and CHG, `IND` for IMM (§11 Q4) |

### 10.2 Demo lines as they would be uploaded (from the cached formula results)

```
FH|SI_REGN|1|01-09-2026|
SI|1|2|0|0|17||1|07-01-2026|8800000|APPL00027326COLL00266688|NA|FIRST CHARGE|0|BATCH_2026-01-09
IMM|1|2||2|COMMERCIAL UNIT|CTS No 156/A/7/1pt|NA|21|1|600|6|NA|Yashodham Business CSL|YashodhamNrRoyalChallengeRestaurantAbaKarmarkarRd|NA|Dindoshi Goregaon East Mumbai|187602|10680|21|400063|IND
BOR|1|COM|0|||ABJFS1083R|CKYC NO|Samcon Builders|38169|Registration Number|||||||||Office No 20 Yashodham Shopping Centre 1st Floor Above HDFC Bank Goregaon East Mumbai 400063|187602|10680|21|400063|IND
BOR|2|IND|0||||||||ABLPB4977B||M|Vijay Dattajirao Bhosale|Dattajirao Bhosale|18639|9820302743|Email|Flat No A/602 Rukmini Vaibhav CHS Ltd Veera Desai Road Azad Nagar Road No 3 Andheri West Mumbai 400058|60744|10680|21|400058|IND
BOR|3|IND|1||||||||AFMPM1762L||M|RAJARAM R MASANE|R MASANE|22463|9820302743|Email|112 Raag Tower Megh Malhar Complex Nr Royal Challenge Hotel Goregaon East Mumbai 400063|187602|10680|21|400063|IND
CHG||ANANDRATHI GLOBAL FINANCE LIMITED|Express Zone B Wing 5th Floor Western Express Highway Opp Oberoi Mall Malad East Mumbai 400063
```

### 10.3 Corrected version of the same record (target output / golden test)

This is what the tool produces for the demo record from the current extract (no title report linked, so SI_TYPE_ID is the default 17), with gender and father's/mother's names typed in the required-details file (`<…>` marks them). It is re-checked whenever a rule changes. The field counts match the spec: FH 4, SI 15, IMM 27, BOR 27, CHG 11. Differences from the demo are deliberate: the area is the LMS carpet area (520), the description is built from the address, survey numbers are written without their label, NARRATION is blank (setting `INCLUDE_LAN_IN_NARRATION`), and borrowers after the primary one are ordered by Neo CIF ID. The line-end convention (no trailing pipe) is still to be confirmed.

```
FH|SI_REGN|1|2026-09-30
SI|1|3|0|0|17||1|2026-07-01|8800000|A00027326C00266688||FIRST CHARGE|0|ARGF2026093000266688
IMM|1|2||5|CMS00260877|COMMERCIAL OFFICE 21 YASHODHAM BUSINESS CSL|156/A/7/1PT||520|6|21|1||Yashodham Business CSL|Aba Karmarkar Road||Goregaon East|187602|10680|21|400063|IND||||
BOR|1|PRF|0|||ABJFS1083R||Samcon Builders|2004-07-01||||||||||Office no 20 Yashodham Shopping Centre 1st Floor Above HDFC Bank Goregaon East Mumbai 400063|||187602|10680|21|400063|IND
BOR|2|IND|1||||||||AFMPM1762L||<M>|RAJARAM R MASANE|<FATHER_MOTHER_NAME>|1961-06-01|9820302743|samconbuilders@gmail.com|112 Raag Tower Megh Malhar Complex Nr Royal Challenge Hotel Goregaon East Mumbai 400063|||187602|10680|21|400063|IND
BOR|3|IND|0||||||||ABLPB4977B||<M>|Vijay Dattajirao Bhosale|<FATHER_MOTHER_NAME>|1951-01-11|9820302743|samconbuilders@gmail.com|Flat no A/602 Rukmini Vaibhav CHS Ltd Veera Desai Road Azad Nagar Road no 3 Andheri West|Mumbai 400058||186829|10680|21|400058|IND
CHG||ANAND RATHI GLOBAL FINANCE LIMITED|Express Zone B Wing 5th Floor Western Express Highway Opposite Oberoi Mall Malad East Mumbai 400063|||81133|10680|21|400063|IND
```

---

## 11. Conflicts, unclear mappings and open questions

Each item has a recommended default so implementation can start; please confirm or correct.

| # | Topic | The conflict or gap | Recommended default |
|---|---|---|---|
| Q1 | Line ending | The demo puts a trailing `\|` on FH only; the spec is silent. | No trailing pipe; exactly (fields − 1) pipes. Confirm with a file CERSAI has accepted. |
| Q2 | NUMBER_OF_RECORDS | Is it "records in the batch file" meaning SI blocks, or all lines? | Number of SI blocks (the demo uses 1 for 7 lines; the range 1–1000 fits SI blocks). |
| Q3 | File name | The meaning of `NNNNN` and `012039` is not explained. | `NNNNN` = entity code `J8084`; `012039` kept literally. Confirm with the CERSAI user manual or a past upload. |
| Q4 | Country | The spec text says `"IN"` for BOR and CHG, `"IND"` for movable assets and "as per master" for IMM. The COUNTRY and STATE masters only use `IND`. | **Resolved 2026-10-09:** follow the format description field by field: `IND` for IMM (as per the COUNTRY master), `IN` for BOR and CHG (`settings.IMM_COUNTRY`, `settings.BOR_CHG_COUNTRY`). |
| Q5 | SI grain | One SI per collateral, or one SI per loan with several IMM lines? | One SI per (application, collateral), as in the demo. |
| Q6 | SI creation date | The demo uses the disbursal date. The legal SI date is usually the mortgage deed / title-deed deposit date. The title reports don't give it either: they are written before disbursal, and none of their dates (initiation, issuance, deed dates) is the mortgage date. | CKYC `First Disbursal Date`, unless the LMS can supply the MODT date. |
| Q7 | Secured amount | Sanctioned or disbursed (263 are partly disbursed)? If one loan has several properties, does each SI carry the full amount? | Sanctioned amount, the full amount on each SI. |
| Q8 | ENTITY_MIS_TOKEN | The demo value is too long. | `A<8 digits>C<8 digits>`. Alternative: the Collateral ID alone (12 characters, already unique). |
| Q9 | BATCH_REF_NUMBER | One per SI or one per file? Generate or leave blank? Must be exactly 20 alphanumeric characters. | One per SI, `ARGF` + date + 8 collateral digits; or blank. |
| Q10 | Guarantors | Should the 34 guarantors be BOR rows? | Exclude them. A guarantor who owns the property becomes an asset owner (TPM). |
| Q11 | Closed loans (C) | **Resolved 2026-09-28:** only active loans (`Loan Status = A`) are processed, so closed and cancelled loans are out of scope. | — |
| Q12 | Shared properties | 24 properties back 2 active applications (e.g. a home loan plus a top-up). Second charge, pari passu, or a modification of the existing SI? | Send to exceptions for manual decision; TYPE_OF_CHARGE comes from the required-details file. |
| Q13 | Asset-owner (ASO) rows | Out of scope, but required whenever IS_TPM = 1 (at least 35 collaterals, and possibly more among the 120 unmatched). | Phase 1: exclude TPM SIs and report them. Phase 2: implement ASO (same pattern as BOR). |
| Q14 | "Registration no" in the old Req details draft | Borrower CIN/LLPIN, or property deed registration number? The title reports give the **deed's** registration details (item 5, e.g. `Sub-Registrar Thane-1, Sr. No. 01359/2005`), but never a borrower's CIN. | Borrower CIN/LLPIN, as the demo's BOR placeholder suggests. If the deed number is also wanted, read it from the title report into ASSET_UNIQUE_ID. |
| Q15 | Survey number separator | The field description says comma; the validation says `^`. | `^` (validation rules are what the upload checks). |
| Q16 | LLPIN format | The spec says 7 alphanumeric characters; real LLPINs look like `AAA-1234`. | Remove the hyphen, then validate against 7 characters; test with CERSAI. |
| Q17 | PAN vs CKYC | "Either PAN or CKYC must be entered" and "PAN of the Borrower is non mandatory" appear in the same field. | Require at least one of the two. |
| Q18 | "Consecutive special characters" | Does whitespace count? | Assume yes, and clean so no two special characters sit together. |
| Q19 | Asset area | Required details or LMS? Carpet or built-up area? Which unit? The LMS area columns don't state a unit. | LMS `Carpert Area` first, assumed to be sq ft (unit 6). **Confirm the unit with the LMS team**: sq m vs sq ft changes the value about 10.8 times. When it is blank, use the title report's area with its own unit (§15.5). required details override. See also Q26. |
| Q20 | OFFICE_NAME / CHG address | **Resolved 2026-09-30:** the CHG row is a fixed constant, the same for every SI, and is not mapped. | — |
| Q21 | Property type mapping | INDUSTRIAL, MIXED USAGE, PLOT and PG have no direct CERSAI type. | The mapping in §8.7. |
| Q22 | File encoding and line breaks | Not specified. | ASCII/UTF-8 without BOM; CRLF line endings (Windows portal). Confirm. |
| Q23 | Mobile/email for non-individuals | The spec says "when BORROWER_TYPE is Individual". | Blank for non-individuals. |
| Q24 | PROJECT_NAME for land/villages | Mandatory, but plots in villages have no project. | Take it from the LMS address, else the project or society named in the title report (all 4 demo files have one); otherwise the colony or village name. Confirm with CERSAI. |
| Q25 | Collateral ID format in PDF file names | **Resolved 2026-09-30:** the PDFs will be named with the exact IDs from COLLATERAL.xlsx (`COLL` + 8 digits); the linker requires an exact match (§15.1). | — |
| Q26 | Which area to file from a title report | When the LMS area is blank, the title report may state carpet, built-up or RERA carpet area, in sq m or sq ft. | File it exactly as the title report states it, with its unit; show the area type in the review sheet. |
| Q27 | Sub-type for units in a building | Is a flat, shop or office in a co-operative society "Structure without Land"? | Yes: 2 (residential) / 5 (commercial) / 8 (other). Confirm with compliance. |
| Q28 | Old vs new survey numbers; town-planning plots | Reports can give `Old S. No. … New S. No. … Hissa No.` or `PF No. 4 of TP`. | New survey number with its Hissa (`12/2`). A town-planning scheme's final plot number goes to PLOT_NUMBER. |
| Q29 | Title report not fully positive | 2 of 4 demo reports say "Positive subject to search", and their Final Certificate is struck through. | File anyway (the mortgage exists), but show the status in the review sheet. |
| Q30 | Missing, unreadable or shared title reports | A collateral may have no PDF, a scanned PDF without text, or share one report with another collateral. | No PDF or no text: the LMS values are used as normal; the type of mortgage defaults to 17, and any remaining gaps go to the required-details file. List the collateral. Shared report: save a copy under each collateral's name. |
| Q31 | Source priority | **Resolved 2026-09-30:** LMS input data first; a title report is read only for values the input data doesn't have (§15.6). Owners follow it too: COLLATERAL first, checked against the title report (§6.3). | — |
| Q32 | Owners differ between COLLATERAL and the title report | Which list to file? It decides IS_ASSET_OWNER, NUMBER_OF_ASSET_OWNERS and IS_TPM. | **Decided 2026-09-30:** file the COLLATERAL owners and show both lists in Check and correct, where the owners can be typed; the title report's owners are used only when COLLATERAL has none, and its extra owners are added only when the COLLATERAL shares show an owner is missing (§6.3). |

---

## 12. Assumptions made in these notes

1. `CERSAI2.0_Masters.xls` v29.3 is the current master. The program should read and log the version and warn if a newer file appears.
2. City, district and state are submitted as **codes**, as in the demo and the masters.
3. Each run covers only SIs not already filed. A register of filed SIs will exist before go-live.
4. Only active loans (`Loan Status = A`) are processed (confirmed 2026-09-28). `X` = cancelled (every X row has disbursed amount 0) and `C` = closed are excluded.
5. The COLLATERAL `Customer Address` text is the same as CKYC `Full Address` (confirmed in the samples) and can be used for its address type.
6. LMS areas are in square feet (to be confirmed).
7. The CHG (charge holder) row is a fixed constant, identical for every SI and never mapped (confirmed 2026-09-30).
8. Record order within an SI block follows the spec's sheet order: SI, IMM, BOR…, (ASO…), CHG.
9. Before each run, every title-report PDF is checked and renamed by hand to `<Application number>_<Collateral ID>.pdf`, using the IDs exactly as they appear in COLLATERAL.xlsx (confirmed 2026-09-30). The four demo PDFs have made-up names, so their content was analysed without comparing it to the file name.

---

## 13. Data-quality findings from the current extract

| Finding | Count | Effect |
|---|---|---|
| Applications with no active loan (23 cancelled-only, 4 closed) | 27 | Excluded (active-only scope) |
| Active applications / SI candidates after filtering | 1,150 / 1,312 | — |
| Collaterals whose active rows disagree on sanctioned amount or disbursal date | 0 | Keep the check |
| Borrowers with no PAN | 436 (431 SPR) | Need a CKYC number |
| Borrowers needing gender and father's/mother's name | 3,141 | Manual work or SQL change |
| COM/LLP borrowers needing CIN/LLPIN | 143 | Manual work or SQL change |
| `Applicant Type` vs `CONSTITUTION` mismatches | ~29 | Exceptions |
| PAN type character inconsistent with constitution | 6 | Warnings |
| Birth dates that would parse into the future with Python's default | 540 | Century rule / SQL change |
| Individual names with non-letters | 20 | Cleaning |
| Non-individual names with disallowed characters | 60 | Cleaning |
| Invalid emails | 5 | Leave blank |
| Borrower addresses over 100 characters after tail removal | 245 | Split across lines |
| Property addresses over 100 characters | 1,194 of 1,312 | Parse, don't copy |
| Property PIN not in master / blank | 8 | Exceptions |
| Property PIN mapping to more than one district | 89 | Disambiguate |
| Properties with no city match in their district | 17 | Override |
| Owner names not matching any borrower | 165 | Review / TPM |
| Collaterals with no owner data | 37 | Required details |
| Properties securing 2 active applications | 24 | Q12 |
| Collateral with carpet area above 99,999.99 | 1 | Area field limit |

---

## 14. Technology recommendation

### 14.1 Decision

**Python + Pandas** as the core, with:

| Library | Purpose |
|---|---|
| `pandas` | Load, join, group and reshape the extracts and masters |
| `openpyxl` | Read `.xlsx`; write the working and final workbooks, the checks report and the required-details file, with formatting, dropdowns and colour-coded cells |
| `xlrd` | Read the legacy `.xls` masters. Alternatively, convert the master to `.xlsx`/CSV once per CERSAI release. |
| `rapidfuzz` | Fuzzy owner↔borrower and city-name matching |
| `pymupdf` (PyMuPDF) | Read the title-report PDFs: text, word positions, and the lines drawn through struck-out words (§15.4) |
| `pdfplumber` | Read the title reports' 13-item table row by row (PyMuPDF's own table finder is an alternative) |
| Standard library (`re`, `datetime`, `dataclasses`) | Cleaning, dates, and one typed record per row type |
| `dev_tools/make_dummy_data.py` | Dummy inputs (LMS extracts and title-report PDFs) covering the kinds of data in the real files, for end-to-end tests with `--dummy`; the golden record of §10.3 is re-checked on the real extract |

**No database for the transformation itself.** A small **SQLite** file (or even an Excel/CSV log at first) is recommended only for the **register of filed SIs** (§14.4).

**No OCR for now:** all four demo title reports are Word exports with a text layer. Add OCR (for example Tesseract) only if scanned reports turn up. Keep the PDFs on the local machine; don't send them to online PDF tools.

### 14.2 Why this choice

| Criterion | Pandas | DuckDB | SQLite | Polars |
|---|---|---|---|---|
| Read `.xlsx` inputs | ✓ (openpyxl) | ✓ (excel extension) | ✗ needs an import step | ✓ (fastexcel) |
| Read the legacy `.xls` master | ✓ (xlrd) | ✗ | ✗ | ✓ (calamine) |
| Joins and group-bys at this size (~12k input rows, ~67k master rows) | Instant | Instant | Fine | Instant |
| Regex cleaning, address parsing, fuzzy matching (the bulk of the work) | Native Python | Awkward (Python UDFs) | Poor | OK |
| Per-row conditional rules, and writing a nested FH/SI/IMM/BOR/CHG text file | Python loop over frames | Still needs Python | Still needs Python | Still needs Python |
| Writing a formatted Excel review file | ✓ | Via pandas | Via pandas | Limited |
| Familiarity / maintainability for an ops-facing team | Highest | Medium | High | Lower |
| Keeping state between runs | Files | File DB | ✓ | Files |

**Reasoning:**

- The hard part of this project is not SQL-shaped. It is cleaning free text, resolving codes (PIN → district → city), fuzzy name matching, many conditional field rules, and writing a hierarchical fixed-position file. All of that is Python code whichever engine does the joins.
- Adding DuckDB would mean two paradigms (SQL plus Python) for joins that pandas does in milliseconds on these volumes, and DuckDB cannot read the `.xls` master.
- SQLite adds a load step and no transformation benefit.
- Polars is excellent, but the data is far too small to need its speed, and its Excel writing is weaker.

**When DuckDB would make sense:** if volumes grew to millions of rows, or if analysts wanted to write the mapping rules in SQL. Even then, pushing joins into the LMS extraction SQL (§9.4) would help more.

### 14.3 Design as built

- **Field order by name.** Each row type has one file in `cersai_tool/rows/` with one function per column, in CERSAI's order, and a plain-English comment above each function saying where the value comes from. Every output line is built from the field list, never by hand-typed position, which rules out the demo's column shifts.
- **Read everything as text** (`dtype=str`), so PINs, PANs, codes and leading zeros never change. The project has its own virtual environment and `requirements.txt` (`Setup.bat` on Windows).
- **Module layout:**
  - `settings.py` — entity code, file-name segment, defaults, thresholds, paths;
  - `sources/` — LMS extracts (active loans only, one bundle per collateral), CERSAI masters, title reports (link by file name, right-PDF check, struck-out text skipped), and the required-details file (write and read);
  - `rules/` — text, dates, identifiers, address parsing, PIN → district → state → city, owner-name matching;
  - `context.py` — per collateral: the usable title report, the owners decision, owner ↔ borrower links, the property location;
  - `rows/` — FH, SI, IMM (+ location), BOR (+ address), CHG;
  - `output/` — working/final Excel layout, step 3 checks, `.DAT` writer;
  - `steps/` — step 1 (build), step 2 (apply), step 3 (create the `.DAT`); `main.py` is the command line (`run1`, `run2`, `run3`, `--batch`, `--yes`, `--dummy`).
- **Batch folders.** Each step 1 creates `Output data/<date>` (or `<date>_2`…). Everything read in step 1 is saved in the folder (`batch data (do not edit).json`), so steps 2 and 3 use exactly the same data even if the input files change. Nothing is overwritten: a replaced file is kept as a `(backup …)` copy, and a final Excel edited by hand is detected before step 2 replaces it.
- **Validation layers (step 3):**
  1. field level (mandatory, length, character set, date, master membership, ASCII only, no `|`);
  2. cross-field (§6.4), including PIN ↔ district ↔ state ↔ city consistency;
  3. file level (duplicate collaterals or MIS tokens, already filed, held, third-party mortgage).

  A failed SI is left out of the `.DAT` and listed with the reason; the rest of the batch still goes.
- **Outputs per batch folder:**
  1. `1 Working_<batch>.xlsx` and `3 Final_<batch>.xlsx`: the demo layout (values from column B, blank row between SI blocks), each cell coloured by its source and status, with a comment naming the field and where the value came from; plus Columns, Issues and Read me sheets; the final file also previews each `.DAT` line in column A;
  2. `2 Required details_<batch>.xlsx` (§9.3);
  3. `4 Checks_<batch>.xlsx`: included SIs, SIs left out and why, warnings (including defaults and differences still open);
  4. `crs_J8084_012039_<timestamp>.DAT` (≤ 1000 SIs per file; serial numbers, FH count and date, and BATCH_REF_NUMBER set when the file is written).

### 14.4 Filed-SI register

`Output data/Filed register.xlsx` keeps one row per Collateral ID: application number, batch, status, `.DAT` file, date and reason.

- **Filed**: the collateral is in a `.DAT` file; step 1 leaves it out of later batches.
- **Pending**: it was left out, with the reason; it comes back in the next batch, with the values typed for it earlier.
- If CERSAI rejects a filed SI, change its status to Pending (or delete the row) so it is processed again.

Still to add: the upload status and the **CERSAI SI ID** returned by the portal, once the portal's batch report format is known (needed for later modification or satisfaction filings).

---

## 15. Legal title reports (`Input data/Legal/`)

Added 2026-09-30 from four demo reports. Their file names were made up only to show the naming format, so this section describes the documents as they are. A mismatch between a file name and its content is never treated as an error here.

**How the title reports are used (decided 2026-09-30):** the LMS input data (CKYC, COLLATERAL) comes first. A title report is read only for values the input data doesn't have or that can't be used: always the type of mortgage, and otherwise gaps such as survey numbers or missing address details (§15.6).

### 15.1 File naming and linking to the data

**Convention:** one PDF per collateral, named `<Application number>_<Collateral ID>.pdf`, for example `APPL00027326_COLL00266688.pdf`. An application with several collaterals has one PDF per collateral, all starting with the same application number.

**The file name is the only link.** The reports contain no application number, loan account, collateral ID, customer ID or PAN, and their "Ref no." is blank or a placeholder. The renaming therefore has to be exact.

**ID format (confirmed 2026-09-30):** the PDFs are named with the IDs exactly as they appear in COLLATERAL.xlsx: `APPL` + 8 digits and `COLL` + 8 digits (e.g. `COLL00266688`). The four demo files use made-up IDs with 9 digits after `COLL`, so they don't link to the current data.

**Linking steps:**

1. List `Input data/Legal/*.pdf` and split each name with the pattern `^(APPL\d{8})_(COLL\d{8})\.pdf$` (extension in any case). A name that doesn't fit → exception "bad file name".
2. Look up the Collateral ID in COLLATERAL (any loan status). It must match exactly, and because a collateral belongs to one application only, COLLATERAL must list it under the application in the file name. Listed under another application → exception "misfiled PDF: rename". Not in COLLATERAL at all → exception "no such collateral" (for example a typo in the name). Don't guess which half of the name is wrong.
3. Two PDFs for the same collateral → exception.
4. Active collaterals with no PDF are listed as "no title report". Their values come from the LMS as normal; the type of mortgage defaults to 17, and any remaining gaps go to the required-details file.

The Legal folder keeps PDFs from earlier days, so only PDFs that concern the current batch are reported. A report of a collateral that is not in the batch (filed earlier, or not active) is ignored silently, even when another collateral of the same application is in the batch. *(implementation)*

**"Right PDF" check.** Before any value is taken from a PDF, confirm that it describes this collateral. A PDF saved under the wrong name would otherwise fill gaps with another property's details. The check compares three things that both sources hold:

| Check | From the title report | Compared with |
|---|---|---|
| Borrower | "Loan applicant" in the header box (ignore a bracketed business name) | CKYC primary borrower's name (fuzzy) |
| Owners | Item 4 names | COLLATERAL `Owner Name` rows for the collateral |
| Unit | Unit number in item 1 (e.g. `1102`, `C-3002`) | Appears in COLLATERAL `Full Address` |

The unit check counts as failed only when the LMS address names a unit of its own; an address without one cannot be compared. *(implementation)*

**Reports swapped between the collaterals of one application.** They share the loan applicant and often the owners, so for two of them only the unit number tells the reports apart. When a report's unit is not in this collateral's LMS address but is the unit of another collateral of the same application (also one filed in an earlier batch), the SI is held whatever the other checks say, with a message naming that collateral (rename the two files and run step 1 again, or answer `WRONG PDF`). *(implementation)* Simulated on the 116 real multi-collateral applications: none of the 278 correct reports is held, and every swap that the unit numbers can tell apart is caught; swaps between collaterals whose addresses name no unit, or the same unit number, cannot be told apart.

If two or more checks fail, hold the SI as "possible wrong PDF" until someone answers `PDF OK` or `WRONG PDF` in the required-details file. **While the SI is held, no value is taken from that PDF** *(implementation)*. A single failure is shown as a warning. The owner comparison doubles as the owner verification (§6.3). Apart from the owners, these checks only verify the file: the values used still follow the LMS-first rule.

### 15.2 Common structure and differences

All four files are **title reports** written by an advocate for ARGFL before disbursal: two by Walke Associates (`APPL00001254_…`, 5 pages each) and two by Lawist Associates (`APPL00001278_…`, 9 pages each). Both firms fill in the **same template**, which looks like ARGFL's standard format:

| Part | What it contains |
|---|---|
| Header | "TITLE REPORT", `Status:`, `Date of Initiation –`, `Ref no. –`, `Date of Issuance –`, addressed to ARGFL |
| Box | `Loan applicant – …` and `Property Owner – …` |
| Table (3 columns: item no. · label · value) | 1 Full Description of the property · 2 Nature of immovable property · 3 Type of Transaction · 4 Name of the owner · 5 List of documents/deed provided to the advocate · 6 (varies, see below) · 7 Title Deeds/documents to be obtained · 8 Nature of Mortgage · 9 Search/Certificate of Non-Encumbrance · 10 Views on Encumbrance · 11 Whether the documents are duly stamped/registered · 12 Certificate of title · 13 Any other suggestion/advice |
| Closing | `REQUIREMENTS:`, `FINAL CERTIFICATE:`, digital signature |

Every page has a text layer (the files are Word exports), so no OCR is needed.

**Where the two firms differ:**

| Aspect | Walke Associates | Lawist Associates |
|---|---|---|
| Item 6 | No row 6: a "LAND HISTORY" block sits inside item 5's rows | Row 6 "Flow of Title" |
| Status | `POSITIVE` | `Positive subject to search` |
| Date of Issuance | One date | Two dates (`11/07/2026; 16/07/2026`) |
| Ref no. | Blank | Placeholder `ARGFL/LA/AN/2025-26/____` |
| Struck-through text | None | Item 9 and the whole Final Certificate are struck through (search still pending), yet the words are still in the text layer |
| Item 11 | Names the registered deed | Just `YES` |
| Item 1 wording | `adm. 60.90 Sq. Mtrs., Carpet area`; `Survey No. …` | `admeasuring 41.99 sq.mtr RERA carpet area`; `Old S. No. … New S. No. … Hissa No. …` |
| PIN in item 1 | 1 of 2 files | Neither file |
| Page layout | Page markers `…1…` | Page markers plus a "Continuation Sheet…" footer; in one file the item 4 row breaks across two pages |

Parsing must also tolerate:
- typos: `Commerical`; the same village spelt `Panchpakhadi` and `Panchapakhadi`; `Sq. Fts.`;
- curly quotes around names (one closing quote is missing), double spaces (`Equitable  Mortgage`), and words split by a hyphen at a line end (`Co-` / `op.`);
- dates as `dd/mm/yyyy`, and occasionally in words in the history text (`3 August 2012`).

### 15.3 Fields we can extract

"4/4" means found in all four demo files.

**Header and table items**

| Field | Where | Found in | Example / format | Use |
|---|---|---|---|---|
| Report status | Header `Status:` | 4/4 | `POSITIVE`, `Positive subject to search` | Review flag (Q29) |
| Date of initiation / issuance | Header | 4/4 | `20/08/2026`; issuance can hold two dates | Reference only; **not** the SI creation date |
| Loan applicant | Header box | 4/4 | A short personal or business name, sometimes with a business name in brackets | "Right PDF" check only |
| Property owner(s) | Header box, item 4 and item 12 | 4/4; the three places agree in every file | `Mr. A & Mrs. B`, `Dr. A`, `A, B and C` (1–3 owners) | Verifies the COLLATERAL owners; used when COLLATERAL has none or the two disagree (§6.3) |
| Nature of property | Item 2 | 4/4 | `Residential`, `Commerical` (typo) | ASSET_TYPE_ID, for `PLOT` properties only |
| Type of transaction | Item 3 | 4/4 | `LAP`, `BT + Top Up (PNB)` | Not used (the LMS already shows which properties secure two loans) |
| Title-deed registration | Item 5 | 4/4 (wording varies) | `registered with the Sub-Registrar at Thane - 1 under Sr. No. 01359/2005 dated 11/03/2005` | Optional ASSET_UNIQUE_ID (Q14) |
| Nature of mortgage | Item 8 | 4/4 | `Equitable Mortgage` | SI_TYPE_ID (always: the LMS doesn't hold it) |
| Encumbrance search | Items 9–10 | 4/4 | "no any charges/mortgage registered", `Search awaited` | Review flag |

**Details inside item 1 (Full Description of the property).** Each one is used only when the LMS doesn't give a usable value (§15.6).

| Field | Found in | Examples from the demo files | Use |
|---|---|---|---|
| Unit type and number | 4/4 | `Flat No. 1102`, `Shop No. 13B`, `Flat No. C-3002` | HOUSE_NUMBER, sub-type; also the "right PDF" check |
| Floor | 4/4 | `11th Floor`, `Ground Floor`, `30th floor` | FLOOR_NUMBER |
| Area, unit and area type | 4/4 | `60.90 Sq. Mtrs. Carpet area`, `342 Sq. Fts. Built up area`, `41.99 sq.mtr RERA carpet area` | ASSET_AREA and ASSET_AREA_UNIT |
| Society / project name | 4/4 | `Society known as 'Devpriya Kores Towers Co-op. Hsg. Soc. Ltd.'`, `project known as Aaradhya Parkwood` | PROJECT_NAME |
| Building / tower | 3/4 | `Building No. 3 known as 'Dev Daya Park'`, `Tower-3(Clove) … building known as Aaradhya Parkwood 1` | BUILDING_NAME |
| Survey, Hissa and town-planning plot numbers | 4/4, in three different phrasings | `Survey No. 149/1/3, 150/6, 152/5`; `Survey No. 142A, 179 & 180, PF No. 4 of TP`; `Old S. No. 260(Pt) and New S. No. 12 and Hissa No. 2` | SURVEY_NUMBER, PLOT_NUMBER |
| Village | 4/4 | `Village Panchpakhadi`, `Village Mahajanwadi` | LOCALITY fallback |
| Taluka / district name | 4/4 | `Taluka and District Thane` | Picks the district when the LMS PIN maps to several |
| Road and area name | 1/4 | `Pokharan Road No. 1, Vartak Nagar` | STREET_NAME, LOCALITY |
| PIN | 1/4 | `Thane (West)400 606` (written with a space) | PIN_CODE, only when the LMS PIN is blank or invalid |
| Extras | 1/4 | `along with one car parking area` | Can go in ASSET_DESCRIPTION |

**Not in the title reports at all:**
- the loan or sanctioned amount (the only "sanctioned" is a "sanctioned layout plan");
- the mortgage (SI creation) date;
- any borrower KYC data (PAN, date of birth, gender, father's name, CKYC, CIN);
- CERSAI codes (city, district, state).

The PIN is missing in 3 of the 4 files.

### 15.4 Extraction rules

1. **Read the table by label, not by number.** Walke's reports have no row 6, so numbering can't be trusted. Match the label text (for example "Nature of Mortgage"), allowing for line breaks inside it. A row whose number cell is empty continues the row above, even across a page break. Use pdfplumber (or PyMuPDF's table finder) to find the rows. Take the text itself from PyMuPDF, because pdfplumber sometimes splits a word (`W e`).
2. **Header fields** (`Status`, the dates, `Loan applicant`, `Property Owner`) come from patterns on the page-1 text.
3. **Search item 1 only** for the PIN, area, survey number and so on. Elsewhere, the report contains other PINs: ARGFL's (`400063`, in the address block), the law firm's (`400706` / `400703`) and a postal code in the signature block. A "first 6-digit number" search would return a wrong PIN in 3 of 4 demo files.
4. **Ignore struck-through text.** Strikethrough is a line drawn through the words, so plain text extraction can't see it. Detect horizontal lines through the middle of words and drop those words.
5. **Take the mortgage type from item 8**, never from the Final Certificate paragraph (it is struck through in 2 of 4 files).
6. **Normalise before matching:** collapse spaces, rejoin `Co- op` → `Co-op`, straighten curly quotes, ignore case, and accept known typos (`Commerical`).
7. **Split the owner list** on `,`, `&` and `and` (any case). If a piece starts with a company word (Company, Co, Associates, Enterprises, Industries, Traders, Sons, Brothers, Pvt, Private, Ltd, Limited, LLP, Trust), join it back to the previous piece: `X AND COMPANY` is one owner (40 LMS owner names contain "and" like this). Then remove honorifics (`Mr.`, `Mrs.`, `Ms.`, `Smt.`, `Shri`, `Dr.`, `M/s.`).
8. **Keep the evidence:** store each value with its page and source sentence for the review sheet.
9. **No text layer** (a scanned report): list it in the exceptions. The LMS values are used as normal, the type of mortgage defaults to 17, and any remaining gaps go to the required-details file.

### 15.5 Mapping to CERSAI fields

These rules turn a title-report value into a CERSAI value. Under the LMS-first rule (§15.6) they run only when the LMS lacks the value. The exceptions are SI_TYPE_ID, which always comes from the title report, and the owners, which are always read to verify the LMS owners. They were run on the four demo files, and every result passes the field's length and character rules.

| CERSAI field | Title-report source | Rule | Demo results (file 1 · 2 · 3 · 4) |
|---|---|---|---|
| SI.SI_TYPE_ID | Item 8 (always) | `Equitable` (also "mortgage by deposit of title deeds") → 17; `Registered` (also "simple" or "English" mortgage) → 18; `Agreement to create mortgage` → 19; `Conditional sale` → 20; anything else → 21, with the text in SI_TYPE_OTHERS | 17 · 17 · 17 · 17 |
| SI.NUMBER_OF_ASSET_OWNERS, SI.IS_TPM, BOR.IS_ASSET_OWNER | Item 4 (fallback: header box), for every collateral with a title report | Split the owners (§15.4 rule 7) and compare them with COLLATERAL `Owner Name`; §6.3 decides which list is used. Match the owners used to this SI's borrowers: an owner who matches no borrower counts in NUMBER_OF_ASSET_OWNERS and makes IS_TPM = 1. | 2 · 1 · 3 · 3 owners |
| IMM.ASSET_TYPE_ID | Item 2, only for `PLOT` properties | `Residential` → 1; `Commercial`/`Commerical` → 2 | 1 · 2 · 1 · 1 |
| IMM.ASSET_SUB_TYPE_ID | Unit word in item 1 | Flat/shop/office/unit/gala → without land (2 / 5 / 8); house/bungalow/row house → with land (1 / 4 / 7); plot/land → land only (3 / 6 / 19) | 2 · 5 · 2 · 2 |
| IMM.ASSET_DESCRIPTION | Item 1 parts | `<unit type> <no> <floor> Floor <project>`, cleaned and cut at a word boundary (≤ 100) | e.g. `Shop 13B Ground Floor Devdaya Park Co-op Hsg Soc Ltd` (52 characters) |
| IMM.SURVEY_NUMBER | Item 1 | The numbers after `Survey No./S. No./Sy. No./CTS No./C.S. No./Gat No./Khasra No.`, split on `,`, `&` and `and`, joined with `^` (Q15). If old and new survey numbers are both given, use the new number plus `/<Hissa>`. | `149/1/3^150/6^152/5` · `142A^179^180` · `12/2` · `12/2` |
| IMM.PLOT_NUMBER | Item 1 | `Plot No.`, or the `FP/PF No.` of a town-planning scheme | — · `4` · — · — |
| IMM.ASSET_AREA | Item 1 | The number after `adm.`/`admeasuring` | `60.90` · `342` · `41.99` · `41.99` |
| IMM.ASSET_AREA_UNIT | Item 1 | sq. mtr/mtrs/metre → 7; sq. ft/fts/feet → 6; sq. yd/yds → 8; acre → 1; guntha → 4; hectare → 5; ares → 2; cents → 3 | 7 · 6 · 7 · 7 |
| IMM.HOUSE_NUMBER | Item 1 | The unit number; several units → `1301/1302` | `1102` · `13B` · `C-3002` · `C-3001` |
| IMM.FLOOR_NUMBER | Item 1 | `11th Floor` → 11; Ground → 0; Basement → −1 | 11 · 0 · 30 · 30 |
| IMM.BUILDING_NAME | Item 1 | `Building No. N known as '…'`, or the tower plus `building known as …` | — · `Dev Daya Park Building No 3` · `Aaradhya Parkwood 1 Tower-3 Clove` (×2) |
| IMM.PROJECT_NAME | Item 1 | `project known as …`, else `Society known as …`, cleaned (≤ 50) | `Devpriya Kores Towers Co-op Hsg Soc Ltd` · `Devdaya Park Co-op Hsg Soc Ltd` · `Aaradhya Parkwood` (×2) |
| IMM.STREET_NAME | Item 1 | A `… Road/Marg/Street` phrase, if any | `Pokharan Road No 1` · — · — · — |
| IMM.LOCALITY (fallback) | Item 1 | A named area after the road, else `Village <name>` | `Vartak Nagar` · `Village Panchapakhadi` · `Village Mahajanwadi` (×2) |
| IMM.DISTRICT (tie-break) | Item 1 | When the LMS PIN maps to several districts, pick the one named here | Thane (4/4) |
| IMM.PIN_CODE (gap only) | Item 1 | Only when the LMS PIN is blank or not in the master: remove the space (`400 606` → `400606`) and check it against the PIN master | `400606` · — · — · — |
| IMM.ASSET_UNIQUE_ID (optional) | Item 5 | Deed registration office and number, if preferred over the CMS number (Q14) | e.g. `Thane-1 01359/2005` |

**Owners are per collateral, not per application.** In the demo files, two reports with the same loan applicant list different owners (a flat owned by two people, and a shop owned by one of them). So owner matching, IS_ASSET_OWNER and IS_TPM must be worked out separately for each collateral.

**These never come from the title report:** SI_CREATION_DATE, TOTAL_SECURED_AMOUNT, all borrower (BOR) identity fields, and the CITY / STATE codes. The CHG row is fixed and is not mapped.

### 15.6 Which source wins, and the effect on manual work

**Order for each field (decided 2026-09-30):** the required-details file (manual override) → LMS input data (CKYC, COLLATERAL) → title report → default.

- A value counts as available in the LMS only if it is **present and usable**:
  - not blank;
  - valid against the masters (e.g. a PIN that exists in the PIN master);
  - for details inside the free-text address, found cleanly by the parsing rules (§8.5).

  Otherwise it counts as missing, and the title report is read.
- The title report is always the source of the **type of mortgage** (SI_TYPE_ID), because the LMS doesn't hold it.
- The title report is always read for the **owners**, to check the COLLATERAL owners. Its owners are used only when COLLATERAL has none; when the COLLATERAL shares show an owner is missing, its extra owners are added (§6.3).
- Apart from the owners, the only comparison between the two sources is the "right PDF" check (§15.1). It verifies the file, but never replaces an LMS value.

**Where the title report is still needed** (active collaterals):

| Gap in the LMS | Collaterals | Title report item |
|---|---|---|
| Type of mortgage (always) | 1,312 | Item 8 |
| No survey or plot number in the address (one of the two is mandatory) | 516 | Item 1 |
| No unit or house number in the address | about 370 | Item 1 |
| `PLOT` property: residential or commercial? | 35 | Item 2 |
| Owners: verified for every collateral (37 have none in the LMS) | 1,312 | Item 4 |
| No LMS area | 10 | Item 1 |
| PIN blank or not in the master | 8 | Item 1, when it states a PIN |
| PIN maps to several districts | 89 | Item 1 district name |
| Project name, locality, floor, building or street not parsed from the address | depends on the parsing | Item 1 |

**Effect on manual work:** the required-details file asks only for what both the LMS and the title report fail to give, and for borrower data (gender, father's/mother's name, CKYC number, CIN/LLPIN), which neither source holds.

### 15.7 Handling the files

- Process the PDFs only on the local machine; don't send them to online PDF tools.
- The project's GitHub repository, which holds them, is private (confirmed 2026-09-30).

---

## 16. Next steps

1. **Answer the open questions in §11**, especially Q1–Q5, Q13 and Q14, which change the output format or scope.
2. **Ask the LMS team for the SQL additions in §9.4.** Gender, father's name, CKYC, CIN and 4-digit years remove most of the manual work.
3. **Get one file that CERSAI has already accepted** (from any past upload) to confirm the line ending, file name, country code and NUMBER_OF_RECORDS.
4. **Set up the Windows laptop** (`Setup.bat`) and do a first live batch with a few collaterals; upload it and check the portal's batch report.
5. **Rename the title-report PDFs** to `<Application number>_<Collateral ID>.pdf` using the exact IDs from COLLATERAL.xlsx, and decide Q26–Q31.
6. **Collect a few more title reports**, ideally from other law firms and including a plot or independent house, to confirm the template and the §15.5 rules.
7. **Asset-owner (ASO) rows** for third-party mortgages (about 36 collaterals in the current extract), once the format is agreed (Q13).
