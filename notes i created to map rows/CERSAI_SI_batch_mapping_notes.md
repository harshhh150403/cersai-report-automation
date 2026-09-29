# CERSAI 2.0 SI Registration Batch File: Automation Mapping Notes

> **Status:** analysis and documentation only. No implementation code has been written yet.
> **Prepared:** 2026-09-27, from the files committed at `5aea88d` ("demo output attached").
> **Updated:** 2026-09-28. Only **active loans** (`Loan Status = A` in COLLATERAL) are in scope; all counts below use that scope.
> **Scope:** the `introduction` (FH), `si` (SI), `immovable_asset` (IMM), `borrower` (BOR) and `charge_holder` (CHG) sheets of the format description.

---

## 1. Project description

Anand Rathi Global Finance Limited (ARGFL) has to register every security interest it creates over a borrower's property with **CERSAI**, the Central Registry of Securitisation Asset Reconstruction and Security Interest of India. In the CERSAI entity master, ARGFL is entity code **`J8084`**. For the loans in this data (mainly loans against property, plus some home loans), the security interest is a **mortgage on immovable property**.

CERSAI 2.0 accepts these registrations in bulk as a **pipe-separated (`|`) `.DAT` batch file**, laid out as defined in `CERSAI2_SIRegistration_BatchFormatDescription.xlsx`. Each file starts with one file-header line. Then, for each security interest, it has one `SI` line, one `IMM` line for the property, one `BOR` line per borrower and one `CHG` line for the lender.

**How it is done today:** someone builds the file by hand in Excel (`DAT FILE WORKING 1.xlsx`). They copy each loan's details from the LMS exports, look up CERSAI master codes (city, district, state, borrower type and so on), type them into cells, and join the cells with `&"|"&` formulas. This is slow. The demo file also shows how easy it is to shift a column, mistype a code or leave placeholder text in the output (see §10).

**What the automation will do:** a Python program will:

1. read the two LMS extracts: `CKYC.xlsx` (borrowers) and `COLLATERAL.xlsx` (loans, properties and owners);
2. merge in the values the LMS does not hold, which operations staff maintain in `Req details.xlsx`;
3. convert every value into CERSAI's format and master codes, using `CERSAI2.0_Masters.xls`;
4. validate every field and every cross-field rule in the format description;
5. write an upload-ready `.DAT` file, a human-readable review workbook, an exceptions report, and a pre-filled `Req details` sheet listing whatever still needs manual input.

**Loans processed:** only **active loans** (`Loan Status = A` in `COLLATERAL.xlsx`). Cancelled (`X`) and closed (`C`) loan accounts are excluded.

**Out of scope for now:** movable and intangible assets, attachment orders, consortium rows and source-of-funds rows. Asset-owner (`ASO`) rows are also out of scope, but any third-party mortgage needs them, so see §6.3 and §11.

---

## 2. Files reviewed

| File | What it is | Key facts |
|---|---|---|
| `Input data/CKYC.xlsx` | LMS borrower/KYC extract. Sheet `Export Worksheet` holds the data and sheet `SQL` holds the query that produced it. | 4,726 rows × 24 columns. **One row per customer per ID document**, so a customer with a PAN and a Udyam number appears twice. 1,177 applications and 4,256 (application, customer) pairs. Disbursal window in the query: 01-APR-26 to 23-SEP-26. Despite the file name, it contains **no CKYC (KIN) numbers**. |
| `Input data/COLLATERAL.xlsx` | LMS loan, property and owner extract (sheets `Export Worksheet` and `SQL`). | 7,685 rows × 33 columns. Grain: application × loan account × customer × collateral × owner (× address type). 1,177 applications, 1,317 loan accounts, 1,341 collateral IDs. |
| `Input data/Req details.xlsx` | Your draft list of values that cannot be automated yet. | Only 8 field names in column A, with no data and no keys. See §9 for a proposed structure. |
| `Output Excel Description/CERSAI2_SIRegistration_BatchFormatDescription.xlsx` | The official field-by-field layout. **This is the primary reference.** | 11 sheets; 5 are in scope. Batch file version 1. |
| `Output excel demo file/DAT FILE WORKING 1.xlsx` | A hand-made sample for one loan (`APPL00027326` / `COLL00266688`), with a threaded comment on most cells explaining its source. | Useful for intent. It contains **35 issues**, several of which would get the file rejected (§10). |
| `Master data/CERSAI2.0_Masters.xls` | CERSAI master tables, **version 29.3, last updated 04-09-2026**. | CITY 45,899 rows · DISTRICT 764 · STATE 36 · PIN_CODE 20,832 (PIN, district) pairs · COUNTRY · GENDER · BORROWER_ASSET_OWNER_TYPE · SI type · financing type · asset category/type/sub-type · AREA_UNIT · ENTITY_CODE. |
| `Master data/crs2_addrss_mstr_pincd_analysis_20260416.xlsx` | An April 2026 analysis of PIN-code master changes: `addition_smry`, `removal_smry`, `Detailed_list` (1,972 "to be added", 20,093 "to be removed") and `effective_master`. | Every (PIN, district) pair in `effective_master` already exists in the v29.3 PIN master, which has 30 extra pairs. **Use v29.3 as the authority**; this file is background only. |

---

## 3. End-to-end workflow

```
 ┌────────────────────┐   ┌────────────────────────┐   ┌───────────────────────────┐
 │ LMS (NEO CAS/LMS)  │   │ CERSAI masters v29.3   │   │ Req details.xlsx          │
 │  SQL → CKYC.xlsx   │   │ (.xls, from the CERSAI │   │ (manual values + overrides│
 │  SQL → COLLATERAL  │   │  portal "Downloads")   │   │  keyed by Collateral/CIF) │
 └─────────┬──────────┘   └───────────┬────────────┘   └─────────────┬─────────────┘
           │                          │                              │
           ▼                          ▼                              ▼
 ┌─────────────────────────────────────────────────────────────────────────────────┐
 │ 1. LOAD      read all sheets as text; check master version; log row counts      │
 │ 2. SCOPE     keep active loans only (Loan Status = A); drop already-filed SIs;  │
 │              1 SI candidate = 1 (Application number, Collateral ID)             │
 │ 3. BUILD     borrowers per application (CKYC, de-duplicated per customer)       │
 │              owners per collateral (COLLATERAL) → match to borrowers → TPM      │
 │ 4. ENRICH    apply Req details overrides / manual values                        │
 │ 5. TRANSFORM clean text, reformat dates, split addresses, map codes,            │
 │              resolve PIN → district → state → city                              │
 │ 6. VALIDATE  field rules (spec) + cross-field rules + file rules                │
 │              failures → exceptions; missing manual data → Req details to fill   │
 │ 7. WRITE     .DAT (only fully valid SIs, ≤1000 per file)                        │
 │              review workbook, exceptions report, Req details template           │
 └─────────────────────────────────────────────────────────────────────────────────┘
           │
           ▼
 Ops reviews → uploads .DAT on the CERSAI portal → checks "Batch Report"
 → records CERSAI SI IDs in a "filed register" (so the next run skips them)
```

It is normally a **two-pass** cycle:

- **Pass 1 (prepare):** produces the exceptions list and a `Req details` sheet pre-filled with keys, LMS reference data and auto-suggestions for every missing value.
- **Ops fills in** the missing values.
- **Pass 2 (build):** produces the `.DAT` file. SIs that are still incomplete are left out and listed, so they don't block the rest of the batch.

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
- Each Collateral ID belongs to exactly one application.
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

- **Owners of a collateral:** the distinct `Owner Name` values for that `Collateral ID` in active rows.
- **Name normalisation:** upper-case; remove honorifics (`MR`, `MRS`, `MS`, `SMT`, `SHRI`, `DR`, `M/S`, `KUM`); remove punctuation; collapse spaces; treat `PVT LTD` and `PRIVATE LIMITED` as the same.
- **Match each owner to the application's borrowers:**
  - an exact normalised match is **matched**;
  - a high fuzzy score (for example rapidfuzz token-set ratio of about 90 or more) is **probable** and must be confirmed in Req details;
  - anything else is a **non-borrower owner**.
- **Current extract (active loans):** 1,636 owner rows. A raw exact text match finds only 798. After normalisation, 1,188 match exactly, 284 more match on tokens (for example `MR. PREM MATA SEVAK MISHRA` vs `PREM MATASEVAK MISHRA`), and 165 don't match. By collateral: 1,120 have all owners among the borrowers, 35 have some non-borrower owners, 120 have no owner matched, and 37 have no owner data.
- **Outputs:**
  - `BOR.IS_ASSET_OWNER` = 1 if that borrower matched an owner of this collateral, else 0.
  - `SI.NUMBER_OF_ASSET_OWNERS` = the number of **non-borrower** owners. It is 0 when every owner is a borrower, as the spec says.
  - `SI.IS_TPM` = 1 if `NUMBER_OF_ASSET_OWNERS > 0`, else 0.
  - If `IS_TPM = 1`, ASO rows are required. **ASO is out of scope**, so for now those SIs go to exceptions or are filed manually (§11 Q13).
- **Dependency:** Req details "Count of property owner / owner name" is the manual override for this step.

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
| 4 | E | NUMBER_OF_ASSET_OWNERS | M | Derived (§6.3) / Req details | Number of non-borrower owners; 0 if all owners are borrowers | ≥ 0; equals the ASO line count |
| 5 | F | NUMBER_OF_CONSORTIUM_MEMBERS | M | Constant | `0` | |
| 6 | G | SI_TYPE_ID | M | Req details "security interest type", otherwise default | Default **`17`** (Equitable Mortgage) | Must be in {17, 18, 19, 20, 21} |
| 7 | H | SI_TYPE_OTHERS | C | Req details | Only when SI_TYPE_ID = 21 | ≤ 100, restricted chars |
| 8 | I | FINANCING_TYPE_ID | M | Constant | `1` (Sole) | FINANCING_TYPE master |
| 9 | J | SI_CREATION_DATE | M | CKYC `First Disbursal Date` | `DD-MM-YYYY` → `YYYY-MM-DD` (demo record: `2026-07-01`) | Not in the future, ≥ 1900. See §11 Q6 on whether the mortgage-deed date should be used instead. |
| 10 | K | TOTAL_SECURED_AMOUNT | M | COLLATERAL `Sanctioned Loan Amount`, from any active row of the collateral (the value is the same on all of them) | Whole rupees, no commas; flag the collateral if its active rows disagree | > 0; at most 13 characters; at most 2 decimals. See §11 Q7. |
| 11 | L | ENTITY_MIS_TOKEN | M | Application number + Collateral ID | Compressed to fit 20 characters: `A` + 8 digits of the application + `C` + 8 digits of the collateral, e.g. `A00027326C00266688` (18 characters, reversible) | Alphanumeric only; ≤ 20. **The demo value is 24 characters.** |
| 12 | M | NARRATION | O | COLLATERAL `Loan Account #` | Proposed `LAN ARGFLLAPMUM2000005018390`, which keeps the loan account number that doesn't fit in the MIS token. Blank is also fine; avoid `NA`. | ≤ 100, restricted chars |
| 13 | N | TYPE_OF_CHARGE | M | Constant, with Req details override | `FIRST CHARGE`. Override for properties already securing another loan (§11 Q12). | Letters and spaces only; ≤ 20 |
| 14 | O | IS_TPM | M | Derived (§6.3) | `1` if NUMBER_OF_ASSET_OWNERS > 0, else `0` | 0 or 1 |
| 15 | P | BATCH_REF_NUMBER | O | Generated | Blank, **or** exactly 20 alphanumeric characters, e.g. `ARGF` + file date `YYYYMMDD` + 8 digits of the collateral = `ARGF2026090100266688` | **Exactly 20 characters**, alphanumeric only |

### 7.3 IMM: immovable asset (`immovable_asset` sheet)

| # | Col | Field | M/O/C | Source | Rule / default | Validation |
|---|---|---|---|---|---|---|
| 1 | B | ROW_TYPE | M | Constant | `IMM` | |
| 2 | C | ASSET_CATEGORY_ID | M | Constant | `1` (Immovable) | |
| 3 | D | ASSET_TYPE_ID | M | COLLATERAL `Property Type`; Req details "Asset type" overrides | Mapping in §8.7 | 1, 2 or 3 |
| 4 | E | ASSET_TYPE_OTHERS | C | Req details, or derived from `Property Type` | Only when the type is 3, e.g. `Industrial`, `Mixed Usage` | ≤ 50, restricted chars |
| 5 | F | ASSET_SUB_TYPE_ID | M | Req details "asset_sub_type"; auto-suggested from address keywords (§8.7) | e.g. an office in a building (Commercial) = `5` | Must belong to ASSET_TYPE_ID |
| 6 | G | ASSET_UNIQUE_ID | O | COLLATERAL `Global collateral No` (proposed) | e.g. `CMS00260877`, or blank. **The demo omits this column entirely.** | Rule is "**" (undefined); keep it alphanumeric |
| 7 | H | ASSET_DESCRIPTION | M | Req details "Asset Description"; otherwise auto-built | Auto: `<type> <unit kind> <house no> <project>`, e.g. `COMMERCIAL UNIT 21 YASHODHAM BUSINESS CSL`, cut at a word boundary | ≤ 100, restricted chars |
| 8 | I | SURVEY_NUMBER | C | Req details "Survey/Plot No"; auto-suggested by regex on the property `Full Address` | Patterns: `CTS No`, `Survey No`, `S No`, `Sy No`, `Khasra/Kh No`, `Gat No`, `Khata/Khewat No`, `Mustatil/Killa` | ≤ 1000; separator question in §11 Q15 |
| 9 | J | PLOT_NUMBER | C | Req details; regex `Plot No` | Leave blank rather than `NA` | ≤ 20; letters, digits, `- _ . /` and space |
| 10 | K | ASSET_AREA | O | Req details "Asset Area"; fallback option: `Carpert Area` | See §11 Q19 | Decimal(7,2), so at most 99,999.99. One collateral exceeds this. |
| 11 | L | ASSET_AREA_UNIT | C | Req details; default `6` (Square Feet) when the area comes from the LMS | Required when ASSET_AREA is filled | AREA_UNIT master |
| 12 | M | HOUSE_NUMBER | C | Regex on the address (`Flat/Office/Unit/Shop/House/H/Gala/Galla/Row House/Bungalow/Door No`); Req details override | e.g. `21` | Required for sub-types 1, 2, 4, 5, 7, 8. ≤ 20; letters, digits, `- /` and space only (**no dot**). |
| 13 | N | FLOOR_NUMBER | O | Regex `(\d+)(st\|nd\|rd\|th)? Floor`; `Ground` → 0; `Basement`/`Lower Ground` → −1; Req details override | e.g. `1` | Whole number from −3 to 100 |
| 14 | O | BUILDING_NAME | O | Req details; regex hints (`Wing`, `Tower`, `Building`, `Bldg`) | Blank if unknown (not `NA`) | ≤ 50, restricted chars |
| 15 | P | PROJECT_NAME | M | **Req details** (to be added); auto-suggested from the address segment after the house and floor segments | e.g. `Yashodham Business CSL` | ≤ 50, restricted chars. For land or villages with no project, see §11. |
| 16 | Q | STREET_NAME | O | Regex on a segment containing `Road/Rd/Marg/Street/Lane/Path/Highway` | e.g. `Aba Karmarkar Road`. **Never remove spaces** to make it fit (the demo did). | ≤ 50; letters, digits, `- /` and space only |
| 17 | R | POCKET | O | Not in source | Blank | |
| 18 | S | LOCALITY | M | Address segment(s) just before the city/state tail; Req details override | e.g. `Dindoshi Goregaon East Mumbai` | ≤ 50, restricted chars |
| 19 | T | CITY | M | CITY master code (§8.3) | e.g. `187602` (GOREGAON EAST) | Must be in DISTRICT |
| 20 | U | DISTRICT | M | PIN_CODE master, from `Property Pincode` | e.g. `10680` (MUMBAI SUBURBAN) | Must be in the PIN master |
| 21 | V | STATE | M | DISTRICT master | e.g. `21` (Maharashtra) | Cross-check against the state name in the address |
| 22 | W | PIN_CODE | M | COLLATERAL `Property Pincode` | 6 digits | Must be in the PIN_CODE master |
| 23 | X | COUNTRY | M | Constant | `IND` (§11 Q4) | COUNTRY master |
| 24–27 | Y–AB | LATITUDE_LONGITUDE_1…4 | O | Not in source | Blank, but **the 4 empty positions must still be written** | |

### 7.4 BOR: borrower (`borrower` sheet)

Fields 7–11 apply only to **non-individual** borrower types (everything except `IND`, including `SPR`). Fields 12–19 apply only to **`IND`**. The unused group stays blank.

| # | Col | Field | M/O/C | Source | Rule / default | Validation |
|---|---|---|---|---|---|---|
| 1 | B | ROW_TYPE | M | Constant | `BOR` | |
| 2 | C | SERIAL_NUMBER | M | Generated | 1..n within the SI, primary borrower first | Unique within the SI |
| 3 | D | BORROWER_TYPE | M | CKYC `CONSTITUTION`, cross-checked with `Applicant Type` | Mapping in §8.7, e.g. Partnership Firm → `PRF` | BORROWER_ASSET_OWNER_TYPE master |
| 4 | E | IS_ASSET_OWNER | M | Derived (§6.3) / Req details | `1` or `0` | |
| 5 | F | BORROWER_UID_TYPE | O | — | Blank | **Must be empty** |
| 6 | G | BORROWER_UID_VALUE | O | — | Blank | **Must be empty** |
| 7 | H | BORROWER_PAN | C | CKYC PAN row, falling back to COLLATERAL `Identification Number` | Non-individual only | `[A-Z]{5}[0-9]{4}[A-Z]`. A 4th-character check against the type (§8.6) raises a warning. |
| 8 | I | BORROWER_CKYC_NUMBER | C | **Req details** (not in the extract) | Non-individual only; needed when there is no PAN (431 SPR) | Exactly 14 digits |
| 9 | J | BORROWER_NAME | C | CKYC `Customer Name`, cleaned | Non-individual only; `&` → `AND`; brackets removed | ≤ 100; letters, digits, `- _ .` and space |
| 10 | K | BORROWER_REGISTRATION_DATE | C | CKYC `Incorporation Date` | `DD-MON-YY` → `YYYY-MM-DD` using the century rule (§8.2) | Not in the future, ≥ 1900 |
| 11 | L | BORROWER_REGISTRATION_NUMBER | C | CKYC `CIN No` row, else Req details "Registration no" | **Required for COM and LLP**; blank for other types | CIN is 21 characters in a fixed pattern; LLPIN per spec (§11 Q16) |
| 12 | M | INDIVIDUAL_PAN | C | CKYC PAN row | Individual only | PAN regex; 4th character `P` |
| 13 | N | INDIVIDUAL_CKYC_NUMBER | C | **Req details** | Individual only; needed when there is no PAN (5 cases) | Exactly 14 digits |
| 14 | O | GENDER | C | **Req details** (not in the extract) | `M`, `F` or `T` | GENDER master. **Don't guess from names.** |
| 15 | P | INDIVIDUAL_NAME | C | CKYC `Customer Name`, falling back to First + Middle + Last | Remove dots and digits (e.g. `SARVES .` → `SARVES`) | Letters and spaces only; ≤ 50 |
| 16 | Q | FATHER_MOTHER_NAME | C | **Req details** (not in the extract) | Don't build it from the middle name (the demo did) | Letters and spaces only; ≤ 50 |
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
| 27 | AB | COUNTRY | M | Constant | `IND` (the spec text says `"IN"`; §11 Q4) | |

### 7.5 CHG: charge holder (`charge_holder` sheet)

This is the same line for every SI. It comes from configuration, not from the data.

| # | Col | Field | M/O/C | Source | Value | Validation |
|---|---|---|---|---|---|---|
| 1 | B | ROW_TYPE | M | Constant | `CHG` | |
| 2 | C | OFFICE_CODE | O | — | Blank | **Must be empty** |
| 3 | D | OFFICE_NAME | M | Config | `ANAND RATHI GLOBAL FINANCE LIMITED` (§11 Q20) | Letters and spaces only; ≤ 100. Branch names like `Mumbai 2` contain a digit and are **not allowed**. |
| 4 | E | ADDRESS_LINE_1 | M | Config | `Express Zone B Wing 5th Floor` | ≤ 100 |
| 5 | F | ADDRESS_LINE_2 | O | Config | `Western Express Highway Opp Oberoi Mall` | ≤ 100 |
| 6 | G | ADDRESS_LINE_3 | O | Config | `Malad East` | ≤ 100 |
| 7 | H | CITY | M | Config | **`81133`** (MALAD, district 10680). The demo's `62995` MALAD EAST belongs to district 355. | Must be in the district |
| 8 | I | DISTRICT | M | Config | `10680` (MUMBAI SUBURBAN) | |
| 9 | J | STATE | M | Config | `21` (Maharashtra) | |
| 10 | K | PIN_CODE | M | Config | `400063` | |
| 11 | L | COUNTRY | M | Config | `IND` | |

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
- **Better fix:** change both SQL queries to output `to_char(..., 'YYYY-MM-DD')` (§9.5).
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
   - Fallbacks, in order: the LMS city token from the address tail, if it exists in that district; then the city named the same as the district; then an exception, with a manual override in Req details.
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

The property `Full Address` is free text, and 1,194 of 1,312 exceed 100 characters. Parse it **into suggestions** with keyword regexes on comma segments:

- `Flat/Office/Unit/Shop/House No …` → HOUSE_NUMBER
- `Nth Floor` → FLOOR_NUMBER
- `… Wing/Tower/Building` → BUILDING_NAME
- `CTS/Survey/Sy/Khasra/Gat No …` → SURVEY_NUMBER
- `Plot No …` → PLOT_NUMBER
- `… Road/Marg/Street` → STREET_NAME
- the segment before the city tail → LOCALITY
- the first named segment → PROJECT_NAME candidate

Show every suggestion in Req details for ops to confirm. Treat the parsing as a helper, not the final authority: the address formats vary too much to trust fully automatic output.

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

**Property Type → ASSET_TYPE_ID / sub-type.** This is a *proposed default* that the business must confirm (§11 Q21); Req details always overrides it.

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

## 9. Manual inputs: `Req details.xlsx`

### 9.1 What the current draft contains and where each item goes

| Draft entry | Maps to | Level | Comment |
|---|---|---|---|
| Count of property owner/owner name | SI.NUMBER_OF_ASSET_OWNERS, SI.IS_TPM, BOR.IS_ASSET_OWNER (and future ASO rows) | Per collateral | Mostly automatable (§6.3). Needed only for the 165 unmatched owners, the 37 collaterals with no owner data, and to confirm fuzzy matches. |
| security interest type | SI.SI_TYPE_ID (+ SI_TYPE_OTHERS) | Per SI | Default 17; override when different. |
| Asset type | IMM.ASSET_TYPE_ID (+ ASSET_TYPE_OTHERS) | Per collateral | Auto-suggested from `Property Type`. |
| Asset Description | IMM.ASSET_DESCRIPTION | Per collateral | Can be auto-built; manual gives better quality. |
| Survey/Plot No | IMM.SURVEY_NUMBER / IMM.PLOT_NUMBER | Per collateral | **Split into two columns**: they are different fields with different rules. |
| Asset Area | IMM.ASSET_AREA | Per collateral | **Add an "Asset Area Unit" column** (required when area is filled). |
| Registration no | Probably BOR.BORROWER_REGISTRATION_NUMBER (the demo's BOR placeholder says "Registration Number") | **Per borrower**, not per SI | Could also mean the property's deed registration number; **confirm (§11 Q14).** |
| asset_sub_type | IMM.ASSET_SUB_TYPE_ID | Per collateral | Auto-suggested by keyword. |

### 9.2 Missing from the draft (cannot be automated from the current extracts)

| Needed value | Field | Level | Current workload |
|---|---|---|---|
| **Gender** | BOR.GENDER | Per individual borrower | **3,141** (every individual) |
| **Father's / mother's name** | BOR.FATHER_MOTHER_NAME | Per individual borrower | **3,141** |
| CKYC number | BOR.BORROWER_CKYC_NUMBER / INDIVIDUAL_CKYC_NUMBER | Per borrower without a PAN | 436 (431 SPR, 4 "NO PAN", 1 individual) |
| CIN / LLPIN | BOR.BORROWER_REGISTRATION_NUMBER | Per COM/LLP borrower | 143 (only 1 CIN exists in the data) |
| Project name | IMM.PROJECT_NAME | Per collateral | 1,312 (mandatory; parsing gives suggestions only) |
| House number | IMM.HOUSE_NUMBER | Per collateral | Wherever parsing fails (≈30%) |
| Locality | IMM.LOCALITY | Per collateral | Wherever parsing fails |
| Floor, building, street | IMM optional fields | Per collateral | Optional; nice to have |
| Area unit | IMM.ASSET_AREA_UNIT | Per collateral | Whenever area is filled |
| Type of charge | SI.TYPE_OF_CHARGE | Per SI | Only for exceptions (24 properties back 2 active loans) |
| City override | IMM/BOR.CITY | Per collateral / borrower | About 17 properties plus borrower addresses with no city match |

### 9.3 Proposed structure for `Req details.xlsx`

I have **not changed the workbook yet**. The design below should be agreed first. The program would generate it pre-filled with keys, LMS reference data and suggestions, so ops only fill in the blanks and correct the suggestions.

**Principle:** a blank cell means "use the automated value or default", and a filled cell overrides it. Every row carries its key, so the file can be merged back without ambiguity.

**Sheet `SI_Asset`** (one row per SI = Application number + Collateral ID)

| Column | Type | Notes |
|---|---|---|
| Application number | key | pre-filled |
| Collateral ID | key | pre-filled |
| Primary borrower name | read-only reference | pre-filled |
| Property address (LMS) | read-only reference | pre-filled |
| LMS property type / area | read-only reference | pre-filled |
| SI_TYPE_ID | dropdown 17–21 | default 17 |
| SI_TYPE_OTHERS | text | when 21 |
| TYPE_OF_CHARGE | text | default FIRST CHARGE |
| ASSET_TYPE_ID / ASSET_TYPE_OTHERS | dropdown / text | suggestion pre-filled |
| ASSET_SUB_TYPE_ID | dropdown filtered by type | suggestion pre-filled |
| ASSET_DESCRIPTION | text ≤ 100 | suggestion pre-filled |
| SURVEY_NUMBER | text | suggestion pre-filled |
| PLOT_NUMBER | text ≤ 20 | suggestion pre-filled |
| ASSET_AREA / ASSET_AREA_UNIT | number / dropdown | |
| HOUSE_NUMBER / FLOOR_NUMBER / BUILDING_NAME / PROJECT_NAME / STREET_NAME / LOCALITY | text | suggestions pre-filled |
| CITY_CODE override | code | only when auto-match fails |
| OWNER_NAMES (confirmed) / OWNER_COUNT | text / number | only when owner matching is unresolved |

**Sheet `Borrower`** (one row per Application number + Neo CIF ID)

| Column | Notes |
|---|---|
| Application number, Neo CIF ID, Customer Name, Constitution, PAN (reference) | pre-filled keys and reference |
| GENDER | dropdown M/F/T (individuals) |
| FATHER_MOTHER_NAME | individuals |
| CKYC_NUMBER | 14 digits, when there is no PAN |
| REGISTRATION_NUMBER | CIN/LLPIN for COM/LLP |
| BORROWER_TYPE override | only when the constitution is wrong |
| IS_ASSET_OWNER override | only when owner matching is unsure |
| CITY_CODE override | only when auto-match fails |

**Sheet `Config`:**
- entity code (`J8084`) and the file-name middle segment (`012039`);
- charge-holder office name, address, city, district, state and PIN;
- defaults: SI type, financing type, type of charge, country;
- the last file sequence and timestamp.

**Sheet `Lists`:** allowed values copied from the masters, to drive dropdown validation.

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
| 15 | IMM | ASSET_AREA (L5) | `600` | Matches neither LMS built-up (650) nor carpet (520) area; the source is undocumented | Take it from Req details and record the source |
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
| 27 | BOR 2/3 | GENDER (O9, O11) | `M` | Not in any source file; origin unknown | From Req details or the LMS |
| 28 | BOR 2/3 | FATHER_MOTHER_NAME (Q9, Q11) | `Dattajirao Bhosale`, `R MASANE` | Built from middle + last name, which is a guess (a middle name can be "WO" = wife of) | From Req details or the LMS |
| 29 | BOR | Cell comments U7/U9/U11 | "ADDRESS MIN 100 CHARACTER" | It is a **maximum** of 100 | — |
| 30 | CHG | CITY…COUNTRY (N13–R13) | In columns N–R | They should be in H–L (after ADDRESS_LINE_2/3) | Fix positions |
| 31 | CHG | Concatenation formula (A13) | Only `B13…E13` | **The CHG line in the file has 4 fields instead of 11**, so every mandatory location field is missing | Emit all 11 |
| 32 | CHG | CITY (N13) | `62995` MALAD EAST | In district 355; PIN 400063 maps to 10680 | `81133` MALAD |
| 33 | CHG | OFFICE_NAME (D13) | `ANANDRATHI GLOBAL FINANCE LIMITED` | Cosmetic: the master spells it "Anand Rathi" | `ANAND RATHI GLOBAL FINANCE LIMITED` |
| 34 | All | Spacer rows and helper text | Blank rows 2/4/6…, `YYY-MM-DD` in F1/J2/R8/R10 | Must not end up in the `.DAT` | Write lines directly, without blanks |
| 35 | All | COUNTRY | `IND` | Matches the COUNTRY master, but the spec text says `"IN"` for BOR and CHG | Keep `IND`; confirm (§11 Q4) |

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

The field counts below were checked against the spec: FH 4, SI 15, IMM 27, BOR 27, CHG 11. `<FATHER_MOTHER_NAME>` marks a value that must come from Req details or the LMS. GENDER `M`, ASSET_AREA `600` and ASSET_DESCRIPTION are kept from the demo as if they came from Req details. The line-end convention (no trailing pipe) is still to be confirmed.

```
FH|SI_REGN|1|2026-09-01
SI|1|3|0|0|17||1|2026-07-01|8800000|A00027326C00266688|LAN ARGFLLAPMUM2000005018390|FIRST CHARGE|0|ARGF2026090100266688
IMM|1|2||5|CMS00260877|COMMERCIAL UNIT|CTS No 156/A/7/1pt||600|6|21|1||Yashodham Business CSL|Aba Karmarkar Road||Dindoshi Goregaon East Mumbai|187602|10680|21|400063|IND||||
BOR|1|PRF|0|||ABJFS1083R||Samcon Builders|2004-07-01||||||||||Office No 20 Yashodham Shopping Centre 1st Floor Above HDFC Bank Goregaon East Mumbai 400063|||187602|10680|21|400063|IND
BOR|2|IND|0||||||||ABLPB4977B||M|Vijay Dattajirao Bhosale|<FATHER_MOTHER_NAME>|1951-01-11|9820302743|samconbuilders@gmail.com|Flat No A/602 Rukmini Vaibhav CHS Ltd Veera Desai Road Azad Nagar Road No 3 Andheri West|Mumbai 400058||186829|10680|21|400058|IND
BOR|3|IND|1||||||||AFMPM1762L||M|RAJARAM R MASANE|<FATHER_MOTHER_NAME>|1961-06-01|9820302743|samconbuilders@gmail.com|112 Raag Tower Megh Malhar Complex Nr Royal Challenge Hotel Goregaon East Mumbai 400063|||187602|10680|21|400063|IND
CHG||ANAND RATHI GLOBAL FINANCE LIMITED|Express Zone B Wing 5th Floor|Western Express Highway Opp Oberoi Mall|Malad East|81133|10680|21|400063|IND
```

---

## 11. Conflicts, unclear mappings and open questions

Each item has a recommended default so implementation can start; please confirm or correct.

| # | Topic | The conflict or gap | Recommended default |
|---|---|---|---|
| Q1 | Line ending | The demo puts a trailing `\|` on FH only; the spec is silent. | No trailing pipe; exactly (fields − 1) pipes. Confirm with a file CERSAI has accepted. |
| Q2 | NUMBER_OF_RECORDS | Is it "records in the batch file" meaning SI blocks, or all lines? | Number of SI blocks (the demo uses 1 for 7 lines; the range 1–1000 fits SI blocks). |
| Q3 | File name | The meaning of `NNNNN` and `012039` is not explained. | `NNNNN` = entity code `J8084`; `012039` kept literally. Confirm with the CERSAI user manual or a past upload. |
| Q4 | Country | The spec text says `"IN"` for BOR and CHG, `"IND"` for movable assets and "as per master" for IMM. The COUNTRY and STATE masters only use `IND`. | `IND` everywhere. |
| Q5 | SI grain | One SI per collateral, or one SI per loan with several IMM lines? | One SI per (application, collateral), as in the demo. |
| Q6 | SI creation date | The demo uses the disbursal date. The legal SI date is usually the mortgage deed / title-deed deposit date. | CKYC `First Disbursal Date`, unless the LMS can supply the MODT date. |
| Q7 | Secured amount | Sanctioned or disbursed (263 are partly disbursed)? If one loan has several properties, does each SI carry the full amount? | Sanctioned amount, the full amount on each SI. |
| Q8 | ENTITY_MIS_TOKEN | The demo value is too long. | `A<8 digits>C<8 digits>`. Alternative: the Collateral ID alone (12 characters, already unique). |
| Q9 | BATCH_REF_NUMBER | One per SI or one per file? Generate or leave blank? Must be exactly 20 alphanumeric characters. | One per SI, `ARGF` + date + 8 collateral digits; or blank. |
| Q10 | Guarantors | Should the 34 guarantors be BOR rows? | Exclude them. A guarantor who owns the property becomes an asset owner (TPM). |
| Q11 | Closed loans (C) | **Resolved 2026-09-28:** only active loans (`Loan Status = A`) are processed, so closed and cancelled loans are out of scope. | — |
| Q12 | Shared properties | 24 properties back 2 active applications (e.g. a home loan plus a top-up). Second charge, pari passu, or a modification of the existing SI? | Send to exceptions for manual decision; TYPE_OF_CHARGE comes from Req details. |
| Q13 | Asset-owner (ASO) rows | Out of scope, but required whenever IS_TPM = 1 (at least 35 collaterals, and possibly more among the 120 unmatched). | Phase 1: exclude TPM SIs and report them. Phase 2: implement ASO (same pattern as BOR). |
| Q14 | "Registration no" in Req details | Borrower CIN/LLPIN, or property deed registration number? | Borrower CIN/LLPIN, as the demo's BOR placeholder suggests. If the deed number is also wanted, put it in ASSET_UNIQUE_ID. |
| Q15 | Survey number separator | The field description says comma; the validation says `^`. | `^` (validation rules are what the upload checks). |
| Q16 | LLPIN format | The spec says 7 alphanumeric characters; real LLPINs look like `AAA-1234`. | Remove the hyphen, then validate against 7 characters; test with CERSAI. |
| Q17 | PAN vs CKYC | "Either PAN or CKYC must be entered" and "PAN of the Borrower is non mandatory" appear in the same field. | Require at least one of the two. |
| Q18 | "Consecutive special characters" | Does whitespace count? | Assume yes, and clean so no two special characters sit together. |
| Q19 | Asset area | Req details or LMS? Carpet or built-up area? Which unit? | Req details; fallback carpet area in sq ft (unit 6). |
| Q20 | OFFICE_NAME / CHG address | Head office (as in the demo) or the lending branch? | Head office for every SI. |
| Q21 | Property type mapping | INDUSTRIAL, MIXED USAGE, PLOT and PG have no direct CERSAI type. | The mapping in §8.7. |
| Q22 | File encoding and line breaks | Not specified. | ASCII/UTF-8 without BOM; CRLF line endings (Windows portal). Confirm. |
| Q23 | Mobile/email for non-individuals | The spec says "when BORROWER_TYPE is Individual". | Blank for non-individuals. |
| Q24 | PROJECT_NAME for land/villages | Mandatory, but plots in villages have no project. | Use the colony or village name from the address; confirm with CERSAI. |

---

## 12. Assumptions made in these notes

1. `CERSAI2.0_Masters.xls` v29.3 is the current master. The program should read and log the version and warn if a newer file appears.
2. City, district and state are submitted as **codes**, as in the demo and the masters.
3. Each run covers only SIs not already filed. A register of filed SIs will exist before go-live.
4. Only active loans (`Loan Status = A`) are processed (confirmed 2026-09-28). `X` = cancelled (every X row has disbursed amount 0) and `C` = closed are excluded.
5. The COLLATERAL `Customer Address` text is the same as CKYC `Full Address` (confirmed in the samples) and can be used for its address type.
6. LMS areas are in square feet (to be confirmed).
7. The charge holder is always ARGFL's head office in Malad East.
8. Record order within an SI block follows the spec's sheet order: SI, IMM, BOR…, (ASO…), CHG.

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
| Collaterals with no owner data | 37 | Req details |
| Properties securing 2 active applications | 24 | Q12 |
| Collateral with carpet area above 99,999.99 | 1 | Area field limit |

---

## 14. Technology recommendation

### 14.1 Decision

**Python + Pandas** as the core, with:

| Library | Purpose |
|---|---|
| `pandas` | Load, join, group and reshape the extracts and masters |
| `openpyxl` | Read `.xlsx`; write the review workbook, exceptions report and Req details template, with formatting, dropdowns and highlighted errors |
| `xlrd` | Read the legacy `.xls` masters. Alternatively, convert the master to `.xlsx`/CSV once per CERSAI release. |
| `rapidfuzz` | Fuzzy owner↔borrower and city-name matching |
| Standard library (`re`, `datetime`, `dataclasses`) | Cleaning, dates, and one typed record per row type |
| `pytest` | Unit tests plus the golden test built from §10.3 |

**No database for the transformation itself.** A small **SQLite** file (or even an Excel/CSV log at first) is recommended only for the **register of filed SIs** (§14.4).

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

### 14.3 Proposed design (for the implementation session)

- **Spec-driven layout.** Read field order, M/O/C, maximum lengths and character rules from `CERSAI2_SIRegistration_BatchFormatDescription.xlsx` at runtime, or from a checked-in copy. Build every output line by field *name*, never by hand-typed position. This rules out the column-shift errors found in the demo.
- **Read everything as text** (`dtype=str`) so PINs, PANs, codes and leading zeros are never changed. pandas 3.x is fine; note that this machine's Python 3.14 has no pandas installed, so the project needs its own virtual environment and `requirements.txt`.
- **Suggested module layout:**
  - `config` — entity code, charge-holder office, defaults, paths
  - `spec` — field definitions loaded from the description workbook
  - `load` — inputs, masters (with version check), Req details
  - `clean` — text, date, PAN, mobile and email normalisers (§8.1, §8.2, §8.6)
  - `geo` — PIN → district → state → city resolver (§8.3)
  - `build` — SI candidates, borrowers, owner matching, IMM/BOR/CHG records (§6, §7)
  - `validate` — field rules, cross-field rules (§6.4) and file rules; outputs exceptions
  - `write` — `.DAT` writer (≤1000 SIs per file, naming, encoding), review workbook, exceptions, Req details template
  - a CLI entry point with `prepare` and `build` modes and a date-window argument
- **Validation layers:**
  1. field level (mandatory, type, length, character set, master membership);
  2. cross-field (§6.4);
  3. file level (count, serial uniqueness, no blank lines).

  A failed SI is left out of the `.DAT` and listed with the reason; the rest of the batch still goes.
- **Outputs per run:**
  1. `crs_J8084_012039_<timestamp>.DAT`;
  2. a review workbook, with one sheet per row type, named column headers, and one line per record, like the demo but correct;
  3. an exceptions report;
  4. a pre-filled `Req details` for the next pass;
  5. a run log (input row counts, master version, filtered counts).

### 14.4 Filed-SI register

To avoid registering the same SI twice across runs (and to support later modification or satisfaction filings), keep one row per Collateral ID with:

- the batch file name and generation time;
- the upload status;
- the **CERSAI SI ID** returned by the portal.

SQLite fits this well (one file, no server, easy to query). An Excel/CSV sheet is fine for a first version.

---

## 15. Next steps

1. **Answer the open questions in §11**, especially Q1–Q5, Q13 and Q14, which change the output format or scope.
2. **Ask the LMS team for the SQL additions in §9.4.** Gender, father's name, CKYC, CIN and 4-digit years remove most of the manual work.
3. **Agree the `Req details.xlsx` structure in §9.3**, and I will regenerate it with keys and suggestions.
4. **Get one file that CERSAI has already accepted** (from any past upload) to confirm the line ending, file name, country code and NUMBER_OF_RECORDS.
5. **Implementation session:** build per §14.3. Start with the golden test in §10.3, then the cleaning and geo rules, then the full run on the current extract. Compare the exceptions with the counts in §13.
