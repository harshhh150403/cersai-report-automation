# CERSAI SI batch: row and column mapping (corrected)

This is a corrected version of `mapping notes.md`, in the same Row → Column layout. Tags show what changed:

- **[Corrected]**: the original logic was wrong or incomplete.
- **[Filled in]**: the original had `??`.
- **[Added]**: a column or rule that was missing.
- No tag: the original logic was right (occasionally with a short clarification).

Every row's field list follows `CERSAI2_SIRegistration_BatchFormatDescription.xlsx`. The data counts quoted below come from the current `CKYC.xlsx` / `COLLATERAL.xlsx` extract.

---

## General rules

- **Only active loans are processed.** **[Corrected]**
  - Before anything else, keep only COLLATERAL rows with **`Loan Status = A`**.
  - This drops cancelled loan accounts (`X`, ₹0 disbursed, carrying older dates and sometimes a different sanctioned amount) and closed loans (`C`).
  - After this filter, **every collateral has exactly one active loan account**, and its sanctioned amount and disbursal date are the same on every row.
  - Current extract: 1,150 applications and **1,312 collaterals**.
  - All other steps, including borrower lists from CKYC, use only applications that remain after this filter.
- **Batch size** = number of unique active Collateral IDs in the file. **[Corrected]** **A file can hold at most 1000 collaterals.** The current extract has 1,312, so it must be split into 2 files, each with its own Row 1.
- **One Collateral ID = one SI block** (Row 2 to the last row below). Each Collateral ID belongs to exactly one application.
- **Date format everywhere is `yyyy-mm-dd`**, written as **text**. If Excel stores a date as a date, the concatenated output shows a serial number such as `38169` (this happened in the demo).
- **Output:**
  - The Excel file is the review copy. The upload file is a pipe-separated `.DAT` text file built from the same values.
  - Write every value as text, so codes, PINs and PANs keep their exact form.
  - No blank spacer rows.
- **Every row must have exactly its spec number of columns**, with empty columns still separated by `|`: Row 1 = 4, SI = 15, IMM = 27, BOR = 27, CHG = 11.
- **Sheets used:** `introduction`, `si` (the original said "is"), `immovable_asset`, `borrower`, `charge_holder`.
- **Column 1 of each row is the ROW_TYPE** (FH / SI / IMM / BOR / CHG). Only Row 1's column 1 is the "File header".
- **Req details should be keyed**, so values can be matched back automatically:
  - property-level values by **Collateral ID** (plus Application number);
  - borrower-level values by **Application number + Neo CIF ID**.
- **Text cleaning (applies to every text column):**
  - Keep only the characters the field allows. Most text fields allow letters, digits, `-`, `_`, `.`, `/` and space; each column below notes its own rule.
  - Remove commas, brackets, `#`, `:` and other disallowed characters.
  - Remove dots after abbreviations (`No.` → `No`).
  - Collapse double spaces. Never leave two special characters side by side (e.g. `. ` or ` /`).
  - When text is too long, cut it at a word boundary, never mid-word and never by deleting spaces.

---

## Row 1 (written once per .DAT file) — `introduction` sheet

- **Column 1 - FILE_HEADER** - value `FH`
- **Column 2 - FILE_TYPE** - value `SI_REGN`
- **Column 3 - NUMBER_OF_RECORDS** - batch size = number of collaterals (SI blocks) in **this file**, between 1 and 1000. **[Corrected]** It is counted per file, not across all batches.
- **Column 4 - FILE_DATE** - today's date, `yyyy-mm-dd`

(From here on, each collateral in the batch gets all of the rows below.)

---

## Row 2 — SI (`si` sheet), 15 columns

- **Column 1 - ROW_TYPE** - value `SI`
- **Column 2 - SERIAL_NUMBER** - 1 to N across the collaterals in this file. It must be unique in the file.
- **Column 3 - NUMBER_OF_BORROWERS** - count of unique `Neo CIF ID` for this collateral.
  - Counting in the COLLATERAL excel works: every collateral row lists all customers of its application, and the set is identical to CKYC.
  - **[Corrected]** Exclude `Customer Type = Guarantor` (a guarantor is not a borrower; 34 collaterals have one). Confirm this with compliance.
  - **[Added]** This number **must equal the number of BOR rows written** for this collateral. Derive both from the same borrower list; never type it by hand (the demo said 2 but had 3 BOR rows).
- **Column 4 - NUMBER_OF_ASSET_OWNERS** - count of the collateral's unique `Owner Name` values that are **not** borrowers.
  - The idea in the original was right. **[Corrected]** The matching must not be an exact text match: only 798 of 1,636 owner names match a `Customer Name` exactly. Owner names carry `Mr.`/`Mrs.`/`Smt.`, leading spaces, different case and spelling variants (`MR. PREM MATA SEVAK MISHRA` vs `PREM MATASEVAK MISHRA`).
  - Before comparing, on both sides: upper-case; remove `MR`, `MRS`, `MS`, `SMT`, `SHRI`, `DR`, `M/S` and punctuation; collapse spaces; treat `PVT LTD` and `PRIVATE LIMITED` as the same.
  - Then compare. An exact normalised match counts as a borrower (1,188 owners). A close fuzzy match (about 284 more) is flagged for someone to confirm in Req details. Anything else counts as a non-borrower owner (about 165).
  - Compare only against the borrowers kept in Column 3. A guarantor who owns the property therefore counts here.
  - 37 collaterals have **no owner name**; take the owner from Req details (legal document).
- **Column 5 - NUMBER_OF_CONSORTIUM_MEMBERS** - `0` by default
- **Column 6 - SI_TYPE_ID** - take the type of mortgage (e.g. `Equitable Mortgage`) from Req details for this collateral. In the `SI_SUB_TYPE_ID` sheet of CERSAI2.0_Masters, find it in "SI Type Description" and take its "SI Sub Type ID / Transaction Sub Type ID".
  - **[Corrected]** Only look at rows where "Asset Category Description" = `Immovable` (IDs 17–21). `Conditional Sale` and `Other` appear more than once for movable and intangible assets (e.g. `Other` = 21, 27 and 29).
  - Valid values: 17 Equitable Mortgage, 18 Registered Mortgage, 19 Agreement to Create Mortgage (UUC), 20 Conditional Sale, 21 Other.
  - If Req details is blank, default to `17`.
- **Column 7 - SI_TYPE_OTHERS** - blank by default. **[Corrected]** It is **mandatory when Column 6 = 21 (Other)**; then take the description from Req details (max 100 characters).
- **Column 8 - FINANCING_TYPE_ID** - financing type is "Sole" by default, so value `1`
- **Column 9 - SI_CREATION_DATE** - disbursal date. **[Corrected]**
  - Recommended: **CKYC `First Disbursal Date`**. It is one value per application, in `DD-MM-YYYY` with a 4-digit year, so convert it to `yyyy-mm-dd` (e.g. `01-07-2026` → `2026-07-01`).
  - The COLLATERAL `Disbursal Date` (from any active row; it is the same on all of them) also works. It is actually the loan booking date and has a 2-digit year (`01-JUL-26` → `2026-07-01`).
  - The two dates differ by 1–45 days in 30 applications. Confirm which one matches the mortgage date.
- **Column 10 - TOTAL_SECURED_AMOUNT** - sanctioned loan amount from any active row of this collateral. After the active-only filter it is the same on every row (checked: no collateral has two different values).
  - **[Added]** Keep a check that flags the collateral if its active rows ever show two different sanctioned amounts.
  - Digits only (max 13), greater than 0.
- **Column 11 - ENTITY_MIS_TOKEN** - `Application number` + `Collateral ID`. **[Corrected]**
  - The plain concatenation is **24 characters, but the maximum is 20** (letters and digits only).
  - Use `A` + last 8 digits of the application number + `C` + last 8 digits of the Collateral ID. For example, `APPL00027326` + `COLL00266688` becomes `A00027326C00266688` (18 characters).
  - Alternative: the Collateral ID alone (12 characters, already unique).
- **Column 12 - NARRATION** - optional. **[Corrected]** Blank by default. `NA` is technically allowed but carries no information. If something useful is wanted, put `LAN <Loan Account #>` (max 100).
- **Column 13 - TYPE_OF_CHARGE** - `FIRST CHARGE` by default (letters and spaces only, max 20). Allow an override from Req details: 24 properties (same `Global collateral No`) secure two active loans and may need a second or pari passu charge.
- **Column 14 - IS_TPM** - **[Corrected]** Not a fixed 0.
  - It is `1` if Column 4 > 0 (some owner is not a borrower = third-party mortgage), otherwise `0`.
  - When it is 1, CERSAI also needs **ASO (asset owner) rows**, which are not in our scope yet. Send these collaterals to an exceptions list for now.
- **Column 15 - BATCH_REF_NUMBER** - **[Corrected]** `BATCH_{date}` is **invalid**: underscores and hyphens are not allowed, and the length must be **exactly 20** letters and digits.
  - The field is optional. Either leave it blank, or use `ARGF` + today `yyyymmdd` + last 8 digits of the Collateral ID, e.g. `ARGF2026090100266688`.

---

## Row 3 — IMM (`immovable_asset` sheet), 27 columns

- **Column 1 - ROW_TYPE** - value `IMM`
- **Column 2 - ASSET_CATEGORY_ID** - `1` by default (Immovable)
- **Column 3 - ASSET_TYPE_ID** - asset = `Property Type` of this collateral in the COLLATERAL excel. Find its ID in the `ASSET_TYPE_ID` sheet of CERSAI2.0_Masters. **[Corrected]**
  - Only look at rows with "Asset Category ID" = `1`. `Other` also exists for movable (7) and intangible (14) assets.
  - A text lookup only works for `RESIDENTIAL` → `1` and `COMMERCIAL` → `2` (compare case-insensitively).
  - The master has no `INDUSTRIAL` (37 collaterals), `MIXED USAGE` (40), `PLOT` (35) or `PG` (1). Map them as follows:
    - INDUSTRIAL → `3`
    - MIXED USAGE → `3`
    - PG → `1`
    - PLOT → by the plot's use (1, 2 or 3; from Req details), otherwise `3`.
  - Req details "Asset type" overrides the lookup.
- **Column 4 - ASSET_TYPE_OTHERS** - **[Corrected]** Not `Others` by default.
  - **Blank unless Column 3 = 3.** When it is 3, give a short description, e.g. `Industrial`, `Mixed Usage`, `Plot` (max 50).
- **Column 5 - ASSET_SUB_TYPE_ID** - take the sub-type from Req details (e.g. `Structure without Land`), then find the ID in the `ASSET_SUB_TYPE_ID` sheet. **[Corrected]**
  - The same description exists under every asset type, so filter by **Asset Type ID = Column 3**.
    - Residential (1): 1 with land, 2 without land, 3 land only, 20 unidentified
    - Commercial (2): 4, 5, 6, 21
    - Other (3): 7, 8, 19, 22
  - Example: an office unit in a commercial building = `5`, not `2`. The demo used 2, which belongs to Residential.
- **Column 6 - ASSET_UNIQUE_ID** - optional. **[Corrected]** Leave it **blank**, not `NA`, because `NA` would be stored as the asset's ID. Optionally use COLLATERAL `Global collateral No` (e.g. `CMS00260877`).
- **Column 7 - ASSET_DESCRIPTION** - from Req details. **Mandatory**, max 100, cleaned. If it is blank, build one, e.g. `COMMERCIAL UNIT 21 YASHODHAM BUSINESS CSL`.
- **Column 8 - SURVEY_NUMBER** - from Req details (CTS / survey / khasra / gat number). **At least one of Column 8 and Column 9 must be filled.** Separate multiple numbers with `^`.
- **Column 9 - PLOT_NUMBER** - from Req details. Max 20 characters. Leave it blank if there is no plot number (not `NA`).
- **Column 10 - ASSET_AREA** - from Req details. Optional. Numbers only, max `99999.99`.
- **Column 11 - ASSET_AREA_UNIT** - from Req details. **[Corrected]**
  - Write the **ID** from the `AREA_UNIT` sheet, not the unit name: `Square Feet` → `6`, `Square Metre` → `7`, `Square Yard` → `8`, `Acre` → `1`, and so on.
  - It is **mandatory when Column 10 has a value**, and blank when Column 10 is blank.
- **Column 12 - HOUSE_NUMBER** - flat/unit/house/shop number, from Req details. **[Added]** Required when the sub-type is "Structure with Land" or "Structure without Land" (1, 2, 4, 5, 7, 8).
  - Letters, digits, `-`, `/` and space only; max 20.
  - Write only the number: `Office No. 21` → `21`.
  - For several units see "Queries" below.
- **Column 13 - FLOOR_NUMBER** - blank by default is fine (optional). If known, from Req details: a whole number from −3 to 100 (Ground = `0`, Basement = `-1`).
- **Column 14 - BUILDING_NAME** - optional, if present in the address or Req details. Max 50.
- **Column 15 - PROJECT_NAME** - **[Corrected] This column is MANDATORY.** It cannot be left blank.
  - Take the project / scheme / society name from Req details, otherwise from the address (e.g. `Yashodham Business CSL`). Max 50.
  - If no project exists (e.g. a village plot), use the colony or village name and flag it for review.
- **Column 16 - STREET_NAME** - from Req details or the address. Optional; max 50; letters, digits, `-`, `/` and space only. Cut at a word boundary: `Aba Karmarkar Road`, never `YashodhamNrRoyalChallenge…`.
- **Column 17 - POCKET** - optional, if present in the address in Req details. Max 20.

**For columns 18 to 22: [Corrected]**

1. **PIN code** = COLLATERAL `Property Pincode` of this collateral. It is already a separate column; cross-check it with the 6-digit PIN at the end of `Full Address`.
2. **Use the CERSAI2.0_Masters `PIN_CODE` sheet**, not `crs2_addrss_mstr_pincd_analysis_20260416` → `effective_master`. The masters file is the newer official list (v29.3, updated 04-09-2026). It already contains every `effective_master` pair plus 30 more.
3. Search the PIN and take its **District Code**.
   - If the PIN is not found (8 collaterals), flag it.
   - If the PIN appears with **more than one District Code** (89 collaterals), pick the district whose name appears in the address; otherwise flag it.
4. **State Code:** take it directly from the `DISTRICT` sheet (District Code → State Code). No state-name matching is needed. (Name matching would fail anyway: `effective_master` has `MAHARASHTRA` while the STATE sheet has `Maharashtra`, so 0 of 36 names match exactly.)
5. **City Code:** in the `CITY` sheet, keep only rows with **our District Code**, as the original said.
   - Then find which of those city names appears in the property address, and take the **longest match** (`GOREGAON EAST` beats `MUMBAI`).
   - Before matching, expand `(E)`/`(W)` to `EAST`/`WEST`.
   - Don't just take the city word from the end of the address (`…, Mumbai, Maharashtra- 400063, India`). That is the LMS city and is often less specific. Use it only as a fallback, if that name exists in the district.
   - Flag anything still unmatched (17 collaterals) for Req details.
   - City names repeat across districts (8,066 of them), so **never pick a city without the district filter**.
6. **Locality:** the area name just before the city in the address (e.g. `Dindoshi Goregaon East`). Take it from Req details if it cannot be found.

- **Column 18 - LOCALITY** - value = locality. **Mandatory**, max 50.
- **Column 19 - CITY** - value = City Code, the number only (e.g. `187602`, never `Goregaon (E)`)
- **Column 20 - DISTRICT** - value = District Code (e.g. `10680`)
- **Column 21 - STATE** - value = State Code (e.g. `21`)
- **Column 22 - PIN_CODE** - value = PIN code (6 digits)
- **Column 23 - COUNTRY** - `IND` by default
- **Column 24 - LATITUDE_LONGITUDE_1** - **[Added]** blank
- **Column 25 - LATITUDE_LONGITUDE_2** - **[Added]** blank
- **Column 26 - LATITUDE_LONGITUDE_3** - **[Added]** blank
- **Column 27 - LATITUDE_LONGITUDE_4** - **[Added]** blank. These 4 columns must still be written as empty values, so the IMM line ends with `IND||||`.

---

## Row 4 to Row (4 + no. of borrowers − 1) — BOR (`borrower` sheet), 27 columns

In these rows, most values come from the CKYC excel. Join CKYC to the working collateral on application number. **[Corrected]**

- CKYC has **one row per customer per ID document** (PAN, GST, Udyam…), so first **group by `Neo CIF ID`**: one BOR row per customer.
- Take the PAN from that customer's row where `Identification Type = PAN`.
- Leave out Guarantors, as in SI Column 3.

**Which group of columns to fill** depends on the **Borrower Type (Column 3)**, not on `Applicant Type`:

- If Column 3 is **not `IND`**, fill columns 7–11 and leave 12–19 blank.
- If Column 3 is **`IND`**, fill columns 12–19 and leave 7–11 blank.

`Applicant Type` is unreliable here: 27 sole proprietorships are tagged `individual`, and 2 individuals are tagged `non_individual`.

- **Column 1 - ROW_TYPE** - value `BOR`
- **Column 2 - SERIAL_NUMBER** - the i-th borrower, 1 to no. of borrowers. Primary borrower first.
- **Column 3 - BORROWER_TYPE** - CKYC `CONSTITUTION` → code from the `BORROWER_ASSET_OWNER_TYPE` sheet. **[Corrected]** The LMS text does not match the master text, so use this fixed mapping instead of a text lookup:

  | CKYC `CONSTITUTION` | Code |
  |---|---|
  | Individual | `IND` |
  | NO PAN | `IND` |
  | Sole Proprietorship Firm | `SPR` |
  | Partnership Firm | `PRF` (the demo wrongly used `COM`) |
  | Limited Liability Partnership | `LLP` |
  | Private Limited Company | `COM` |
  | Public Limited Company | `COM` |
  | Hindu Undivided Company | `HUF` |
  | Private Trust | `TRS` |
  | anything else | flag for review |

- **Column 4 - IS_ASSET_OWNER** - `1` if this borrower is an owner, else `0`. **[Corrected]**
  - CKYC has **no owner information**. Use COLLATERAL `Owner Name` of this collateral and match it to this borrower's `Customer Name`, using the same normalised matching as SI Column 4.
  - Check: the number of borrowers with `1` + SI Column 4 = total owners of the collateral.
- **Column 5 - BORROWER_UID_TYPE** - blank (must be empty)
- **Column 6 - BORROWER_UID_VALUE** - blank (must be empty)
- **Column 7 - BORROWER_PAN** - PAN from CKYC, when Column 3 is not `IND`. Format: 5 letters, 4 digits, 1 letter.
- **Column 8 - BORROWER_CKYC_NUMBER** - **[Corrected]** Not always blank.
  - **Required when the non-individual has no PAN.** That is 431 sole proprietorships today.
  - It is a 14-digit number. It is **not in the CKYC excel** despite the file's name, so take it from Req details.
- **Column 9 - BORROWER_NAME** - **[Corrected]** = CKYC `Customer Name` (the original said "Pan number"), when Column 3 is not `IND`.
  - Clean it: `&` → `AND`; remove brackets and dots (`JOGARAM PUROHIT (HUF)` → `JOGARAM PUROHIT HUF`). Max 100.
- **Column 10 - BORROWER_REGISTRATION_DATE** - CKYC `Incorporation Date`, only when Column 3 is not `IND`. **[Added]**
  - It is `DD-MON-YY` with a **2-digit year**, so convert it carefully: `01-JUL-04` → `2004-07-01`.
  - If `20YY` would be in the future, use `19YY`.
- **Column 11 - BORROWER_REGISTRATION_NUMBER** - **[Filled in]**
  - CIN (21 characters) when Column 3 = `COM`; LLPIN when Column 3 = `LLP`. It is **mandatory for those two** and blank for everyone else.
  - Source: the CKYC row with `Identification Type = CIN No` (only 1 exists today). Otherwise Req details "Registration no" (143 borrowers).
- **Column 12 - INDIVIDUAL_PAN** - PAN from CKYC, when Column 3 = `IND`
- **Column 13 - INDIVIDUAL_CKYC_NUMBER** - blank by default. **[Corrected]** **Required when the individual has no PAN** (5 today: 4 "NO PAN" plus 1 other). 14 digits, from Req details.
- **Column 14 - GENDER** - **[Filled in]** `M`, `F` or `T` (`GENDER` sheet), when Column 3 = `IND`. It is **required for individuals**, but it is **not in the CKYC excel**, so take it from Req details (or ask the LMS team to add it to the query).
- **Column 15 - INDIVIDUAL_NAME** - CKYC `Customer Name`, when Column 3 = `IND`. Letters and spaces only (e.g. `SARVES .` → `SARVES`); max 50.
- **Column 16 - FATHER_MOTHER_NAME** - **[Filled in]** Father's or mother's name, when Column 3 = `IND`.
  - It is **required for individuals** and **not in the CKYC excel**, so take it from Req details.
  - Don't build it from the middle name. The middle name is sometimes `WO` ("wife of").
- **Column 17 - DATE_OF_BIRTH** - CKYC `Date Of Birth`, when Column 3 = `IND`. **[Corrected]**
  - It is `DD-MON-YY` with a 2-digit year. Use `20YY` only if that makes the person at least 18 today; otherwise use `19YY` (e.g. `11-JAN-51` → `1951-01-11`).
  - Python's default parsing would give `2051` for 553 borrowers.
  - (The original line said "Date of Birth Individual Pan"; the "Individual Pan" part was a copy-paste leftover.)
- **Column 18 - MOBILE_NUMBER** - CKYC `Mobile No`, when Column 3 = `IND`. **[Corrected]** Remove the `+91--` prefix so exactly 10 digits remain.
- **Column 19 - EMAIL** - **[Added]** (this column was missing) CKYC `Email ID`, when Column 3 = `IND`. Optional.
  - Trim spaces. Leave it blank if it breaks the rule: one `@`, a dot after the `@`, only letters, digits, `_` and `.` (**no hyphen**), max 50.

As in Row 3 for the property, each borrower's address goes in their own row:

- **[Corrected] Which address:** 41 customers have more than one address in CKYC, and CKYC doesn't say which type each one is. Take the address from COLLATERAL `Customer Address` using `ADDRESS_TYPE`:
  - individuals: `ResidentialAddress` first;
  - non-individuals: `OfficeAddress` first.
- **[Corrected]** Remove the ending `, <City>, <State>- <PIN>, India`, because city, state and PIN have their own columns. Then clean the text (no commas, dots, brackets…).
- **Column 20 - ADDRESS_LINE_1** - the address, up to 100 characters. **[Corrected]** Cut at a **word boundary**, not at exactly 100 characters (and it is not "alphanumeric only": spaces and `/` are part of it).
- **Column 21 - ADDRESS_LINE_2** - if the address is longer than 100 characters, the rest goes here (max 100, cut at a word boundary); otherwise blank.
- **Column 22 - ADDRESS_LINE_3** - if it still doesn't fit, the rest goes here (max 100); otherwise blank. (After removing the city/state/PIN ending, no borrower address in the current data needs more than 2 lines.)
- **Column 23 - CITY** - City Code. **[Corrected]** This column is **mandatory**, so it cannot stay blank in the final file. Use the same steps as Row 3, with the borrower's PIN and address.
- **Column 24 - DISTRICT** - District Code (same steps as Row 3). The borrower's PIN is the 6 digits before `, India` at the end of the address.
- **Column 25 - STATE** - State Code
- **Column 26 - PIN_CODE** - PIN code
- **Column 27 - COUNTRY** - **[Corrected]** `IND`, not `IN`. The masters only have `IND`, and Row 3 and the last row use `IND`, so keep it the same everywhere. The spec text says `"IN"` here; confirm with the first upload.

---

## Last row — CHG (`charge_holder` sheet), 11 columns

**[Corrected]** The charge-holder row has **11 columns, not 13**. The original listed 4 blank columns (5–8), but only 2 of them exist (ADDRESS_LINE_2 and ADDRESS_LINE_3). As a result, city, district, state, PIN and country were 2 positions too far right.

- **Column 1 - ROW_TYPE** - value `CHG`
- **Column 2 - OFFICE_CODE** - blank (must be empty)
- **Column 3 - OFFICE_NAME** - `ANAND RATHI GLOBAL FINANCE LIMITED` (letters and spaces only)
- **Column 4 - ADDRESS_LINE_1** - `Express Zone B Wing 5th Floor Western Express Highway Opposite Oberoi Mall Malad East Mumbai 400063` (99 characters, which is within the 100 limit; don't make it longer)
- **Column 5 - ADDRESS_LINE_2** - blank
- **Column 6 - ADDRESS_LINE_3** - blank
- **Column 7 - CITY** - **[Corrected]** `81133` (MALAD), not `62995`. `62995` MALAD EAST belongs to district 355 (Mumbai), but PIN 400063 belongs to district `10680` (Mumbai Suburban), so the city must come from district 10680.
- **Column 8 - DISTRICT** - `10680`
- **Column 9 - STATE** - `21`
- **Column 10 - PIN_CODE** - `400063`
- **Column 11 - COUNTRY** - `IND`

---

## Values from the legal document (Req details)

**[Corrected]** The original list (si_type, cts no, property owner, area) is incomplete. The full list:

**Per collateral** (key: Application number + Collateral ID):

| Value | Goes to |
|---|---|
| Type of mortgage | SI Column 6 (+ Column 7 if "Other") |
| Type of charge, only if not first charge | SI Column 13 |
| Property owner name(s), when the LMS has none or they don't match a borrower | SI Column 4, BOR Column 4 |
| Asset type (only to override the LMS property type) | IMM Column 3 / 4 |
| Asset sub-type (with land / without land / land only) | IMM Column 5 |
| Asset description | IMM Column 7 |
| CTS / survey no. | IMM Column 8 |
| Plot no. | IMM Column 9 |
| Area and **area unit** | IMM Columns 10–11 |
| House / flat / unit no. | IMM Column 12 |
| Floor, building, street (optional) | IMM Columns 13, 14, 16 |
| **Project name** (mandatory) | IMM Column 15 |
| Locality, if not found in the address | IMM Column 18 |
| City code, only if auto-match fails | IMM Column 19 |

**Per borrower** (key: Application number + Neo CIF ID):

| Value | Goes to |
|---|---|
| Gender (individuals) | BOR Column 14 |
| Father's / mother's name (individuals) | BOR Column 16 |
| CKYC number (only when there is no PAN) | BOR Column 8 or 13 |
| CIN / LLPIN (companies and LLPs) | BOR Column 11 |

---

## Queries (answered)

**1. Is PAN compulsory? What about gender and father's name?**

- **PAN alone is not compulsory**, but every borrower needs **either a PAN or a 14-digit CKYC number**. 436 current borrowers have no PAN (431 sole proprietorships, 4 "NO PAN", 1 individual), so they need a CKYC number from Req details.
- **Gender, father's/mother's name and date of birth are required for every individual borrower** ("Conditional" in the spec means required when the borrower type is Individual).
- None of them is needed for non-individuals.
- Gender and father's/mother's name are **not in the CKYC excel**. That is about 3,141 individuals today, so either fill them in Req details or ask the LMS team to add both to the CKYC query (much less manual work).

**2. What if the full address lists two flats under a single Collateral ID?**

This happens in **37 collaterals** today (e.g. `Flat No. 1301 & 1302`, `Shop No. 18 To 25`).

**Recommended:**

- Keep **one SI block with one IMM row**. One Collateral ID is one mortgaged security.
- **Column 12 HOUSE_NUMBER:** join the numbers with `/`, e.g. `1301/1302` or `D-11/D-12` (only `-`, `/` and space are allowed; max 20). If that is longer than 20 characters, write the first unit here and list every unit in **Column 7 ASSET_DESCRIPTION** (e.g. `RESIDENTIAL FLATS 1301 AND 1302 ACROPOLIS`).
- **Column 13 FLOOR_NUMBER:** the floor if all units are on one floor; otherwise leave it blank (it is optional).
- **Column 10 ASSET_AREA:** the combined area of all the units, from the legal document.

If the legal team treats each flat as a separate mortgage, it should be split into separate collaterals in the LMS; the automation should not split it. Also confirm whether CERSAI accepts more than one IMM row under one SI.

---

## Summary of corrections vs. the original notes

| # | Where | Original | Corrected |
|---|---|---|---|
| 1 | General | Batch = all unique collaterals | Active loans only (`Loan Status = A`); max 1000 collaterals per file (1,312 today → 2 files) |
| 2 | Row 1 Col 3 | One row for all batches | One Row 1 per file; the count is per file |
| 3 | SI Col 3 | Count all Neo CIF IDs | Exclude guarantors; must equal the number of BOR rows |
| 4 | SI Col 4 | Exact name match | Normalised + fuzzy match (exact matches only 798 of 1,636 owners) |
| 5 | SI Col 6 | Lookup by description | Filter to Immovable (IDs 17–21); default 17 |
| 6 | SI Col 7 | Always blank | Mandatory when SI type = 21 |
| 7 | SI Col 9 | Disbursal date of the application | CKYC First Disbursal Date recommended (4-digit year); the COLLATERAL date needs 2-digit-year handling; they differ in 30 applications |
| 8 | SI Col 10 | Sanction from any row | Any active row is correct (the value is the same on all of them); added a check that flags a mismatch |
| 9 | SI Col 11 | App + Collateral (24 characters) | Max 20 → `A00027326C00266688` |
| 10 | SI Col 12 | `NA` | Blank (or `LAN <loan account>`) |
| 11 | SI Col 14 | Always 0 | 1 when SI Col 4 > 0 (TPM → exceptions for now) |
| 12 | SI Col 15 | `BATCH_yyyy-mm-dd` | Blank, or exactly 20 letters/digits |
| 13 | IMM Col 3 | Text lookup | Only RESIDENTIAL/COMMERCIAL match; mapping for the other types |
| 14 | IMM Col 4 | `Others` | Blank unless Col 3 = 3 |
| 15 | IMM Col 5 | Req details value | Convert to the ID for the matching asset type |
| 16 | IMM Col 6 | `NA` | Blank (or Global collateral No) |
| 17 | IMM Col 11 | Unit from Req details | Unit **ID**, required when area is given |
| 18 | IMM Col 12 | Req details | Also required for structure sub-types; digits only, `/` for several units |
| 19 | IMM Col 15 | Optional | **Mandatory** |
| 20 | IMM Cols 18–22 | `effective_master` + state-name lookup | CERSAI masters `PIN_CODE` → `DISTRICT` (gives the State Code); city = longest match within the district |
| 21 | IMM Cols 24–27 | Missing | 4 blank lat/long columns |
| 22 | BOR join | One BOR row per CKYC row | Group by Neo CIF ID first (CKYC has one row per ID document) |
| 23 | BOR Col 3 | Text lookup | Fixed CONSTITUTION → code mapping |
| 24 | BOR Col 4 | From CKYC | From COLLATERAL owner names |
| 25 | BOR Cols 7–19 | Based on Applicant Type | Based on Borrower Type (Col 3) |
| 26 | BOR Col 8 / 13 | Blank | CKYC number required when there is no PAN |
| 27 | BOR Col 9 | "Pan number" | Customer Name |
| 28 | BOR Cols 11, 14, 16 | `??` | CIN/LLPIN, Gender, Father/Mother name, from Req details |
| 29 | BOR Col 17 | DOB | 2-digit-year century rule |
| 30 | BOR Col 18 | Mobile | Remove `+91--` |
| 31 | BOR Col 19 | Missing | EMAIL |
| 32 | BOR Cols 20–22 | First 100 characters | Pick the address by type, remove the city/state/PIN ending, clean, split at word boundaries |
| 33 | BOR Col 23 | Blank for now | Mandatory; same logic as IMM |
| 34 | BOR Col 27 | `IN` | `IND` (same as the other rows; confirm) |
| 35 | CHG | 13 columns | 11 columns; city moves to Col 7 |
| 36 | CHG Col 9 → Col 7 | City `62995` | `81133` (district 10680) |
| 37 | Legal-document list | 4 items | Full list per collateral and per borrower (above) |
