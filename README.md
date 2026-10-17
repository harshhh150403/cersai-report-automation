# CERSAI SI batch tool

Builds the CERSAI 2.0 security-interest registration batch file (`.DAT`) for Anand Rathi Global Finance
from the LMS extracts (CKYC and COLLATERAL) and the advocates' title reports (legal PDFs). It runs in
three steps, with a person checking and filling in values between them.

The mapping rules behind every column are in
[notes i created to map rows1/mapping notes corrected.md](<notes%20i%20created%20to%20map%20rows1/mapping%20notes%20corrected.md>)
(short) and
[notes i created to map rows/CERSAI_SI_batch_mapping_notes.md](<notes%20i%20created%20to%20map%20rows/CERSAI_SI_batch_mapping_notes.md>)
(full).

## One-time setup (Windows)

1. Install Python 3.11 or newer from [https://www.python.org/downloads/](https://www.python.org/downloads/). On the first installer screen,
   tick **Add python.exe to PATH**.
2. Put this project folder on the laptop.
3. Double-click **Setup.bat**. It creates a `.venv` folder and installs the packages listed in
   `requirements.txt`. You need to do this again only if `requirements.txt` changes.

## Folders

| Folder / file                         | What goes there                                                                                                                           |
| ------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `Input data/CKYC/`                  | The day's CKYC extract: exactly one`.xlsx` file (any name). Its `Gender` column (`Male` / `Female`) is used when present.               |
| `Input data/COLLATERAL/`            | The day's COLLATERAL extract: exactly one`.xlsx` file. It may have hand-filled columns (below).                                                              |
| `Input data/Legal/`                 | Title reports named`<Application number>_<Collateral ID>.pdf`, e.g. `APPL00027326_COLL00266688.pdf`. PDFs from earlier days can stay. |
| `Master data/CERSAI2.0_Masters.xls` | The CERSAI masters. Replace it when CERSAI publishes a new version.                                                                       |
| `Output data/<date>/`               | One folder per batch (a second batch on the same day gets`_2`).                                                                         |
| `Output data/Filed register.xlsx`   | Every collateral that went into a`.DAT` file (Filed) or was left out (Pending).                                                         |

## Daily use

1. **Build.** Put the day's files in `Input data`, then double-click **1 Build working Excel.bat**.
   It creates `Output data/<today>/` with:

   - `1 Working_<today>.xlsx`: every value the tool found, laid out like the demo file.
   - `2 Required details_<today>.xlsx`: the values it could not find, and what someone should check.

   Collaterals already filed in an earlier `.DAT` file are left out automatically. Only active loans
   (COLLATERAL Loan Status = A) are used.
2. **Fill.** Open `2 Required details`, fill the yellow cells in **Borrowers** and **Collaterals**, and
   look at **Check and correct**. Save and close the file, then double-click
   **2 Apply required details.bat**. It creates `3 Final_<today>.xlsx` and says which SIs are still
   not ready. You can fill more values and run step 2 as often as you like.
3. **Create the .DAT file.** Review `3 Final` (you may correct values directly in it), then
   double-click **3 Create DAT file.bat**. It writes `crs_J8084_012039_<date and time>.DAT` in the batch
   folder, plus `4 Checks_<today>.xlsx` listing what was included, what was left out and why, and any
   warnings. The filed register is updated.
4. Upload the `.DAT` file to CERSAI.

Steps 2 and 3 work on the latest batch by default and ask you to confirm it (press Enter), or you can
type another batch name such as `2026-09-30`.

## Filling the required-details file

- **Yellow, empty**: needed. The tool never assumes a value it cannot find (type of mortgage, asset
  type, type of charge, secured amount...): it leaves it blank and asks for it here.
  **Grey**: found by the tool (type over it only to correct it). **Orange**: the LMS and the title
  report disagree. **Blue**: typed for the same collateral or borrower in an earlier batch.
  **Dark grey**: does not apply.
- Hover over a cell to see why it is needed or where its value came from.
- Dates: `YYYY-MM-DD` (`DD-MM-YYYY` also works). City and district: type the name or the CERSAI code.
  Fill the PIN first, because city and district are worked out from it.
- Owners: names separated by `;`.
- To leave an optional value empty in the batch file, type `BLANK`.
- **Check and correct** shows LMS and title-report values side by side. Type in **Your value** only if
  the value the tool will use is wrong. The empty rows at the bottom let you correct any other value:
  give the Collateral ID, the row (SI, IMM or BOR), the Neo CIF ID for a BOR row, the field and the
  value.

## How values are chosen

1. A value typed in the required-details file (or the final Excel) always wins.
2. Otherwise CKYC / COLLATERAL, including the hand-filled columns below when the file has them.
3. Otherwise the title report, only for what the LMS lacks (survey or plot number, area, project,
   PIN...). Without the hand-filled columns, the type of mortgage (SI type) is only in the title report.
4. Otherwise the value is asked for in the required-details file. Nothing is assumed or guessed: no
   type of mortgage when there is no usable title report, and the asset sub-type (structure with or
   without land) is never worked out from the address. Only a PLOT (COLLATERAL Property Type) is
   taken as Land without Structure.

### Hand-filled COLLATERAL columns

Someone may add these columns to the COLLATERAL extract and fill them by hand. Each is read only
when the file has it (step 1 lists the ones it found); a missing column or a blank cell leaves the
value to the usual sources above. The names are matched ignoring case, spaces and punctuation.

| Column | Fills | Write it as |
|---|---|---|
| `Father Name` | FATHER_MOTHER_NAME of an individual borrower | The name, on that customer's rows (any collateral of the application). |
| `Type of Mortgage` | SI_TYPE_ID | `Equitable Mortgage`, `Registered Mortgage`, `Agreement to Create Mortgage`, `Conditional Sale`, or `Other: <description>` (a code 17-21 also works). |
| `Structure Type` | ASSET_SUB_TYPE_ID | `Structure with Land`, `Structure without Land` or `Land without Structure`. The code follows from the asset type (Property Type). When blank, a PLOT is Land without Structure and any other property is asked for. |
| `Survey No.` | SURVEY_NUMBER | Several numbers separated by commas; a Hissa as `12/2`. |
| `Project Name` | PROJECT_NAME | The society, project or scheme name. |

- A value may be on one row of the collateral (or of the customer, for `Father Name`) or on all of
  them: blank rows are ignored. If its rows give different values, none is used and the difference
  is reported.
- `NA`, `-` and similar are treated as blank.
- When a column and the title report disagree (type of mortgage, survey numbers), the column's value
  is used and the difference is shown in Check and correct.
- Wording the tool does not recognise (for example `EM`, or `Flat` as a structure type) is not
  guessed at: it is reported in the Issues sheet, and the value comes from the usual sources or is
  asked for.

Property owners come from COLLATERAL (`Owner Name`) too. Their `Percentage Share` should add up to 100%
and `Ownership Status` should match the number of owners (Single = one, Joint = more). The title report
checks them. When it names other owners, the COLLATERAL owners are still used and both lists are shown in
Check and correct. The title report's owners are used only when COLLATERAL has no owners. The one gap it
fills: when the COLLATERAL shares show an owner is missing (for example 50% in total) and the title report
names the same owners plus others, those others are added.

Each owner is then matched to a borrower by name. When the names are only similar, the tool decides on its
own and lists the decision in Check and correct, where a `Y` or `N` changes it:

- a spelling difference or a missing surname: taken to be the same person;
- different given names (usually a relative, such as a wife whose name contains her husband's): taken to
  be different people.

## Things the tool will tell you about

| Situation                                                                                                               | What happens                                                                                                                                                                                        |
| ----------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| The title report may belong to another property (applicant, owners and flat number do not match)                        | The SI is**held**. Answer `PDF OK` or `WRONG PDF` in Check and correct, then run step 2 again. A held report's values are not used.                                                       |
| An owner is not a borrower (third-party mortgage)                                                                       | The SI is left out of the`.DAT` file: CERSAI needs asset-owner rows, which the tool does not write yet. File it by hand, or correct the owners if they are in fact borrowers.                     |
| An owner's name only resembles a borrower's (a spelling difference, or different given names)                           | The tool decides and lists the decision in Check and correct; nothing is held. Type`N` or `Y` there only if it is wrong. One answer covers every collateral of the application with that owner. |
| COLLATERAL and the title report name different owners                                                                   | The COLLATERAL owners are used and both lists are shown in Check and correct. Type the owners there if the title report is right.                                                                   |
| The COLLATERAL owners' Percentage Shares do not add up to 100%, or Ownership Status does not match the number of owners | Shown in Check and correct. If the title report names the same owners plus others, those others are added.                                                                                          |
| One COLLATERAL Owner Name holds two people ("A and B")                                                                  | Taken as two owners when each is a borrower (company names with "and" stay whole); otherwise shown in Check and correct.                                                                            |
| Two collaterals of one application seem to have each other's title reports                                              | Both SIs are held; rename the two files and run step 1 again, or answer`WRONG PDF`.                                                                                                               |
| Two collaterals of one application have the same address                                                                | A warning on both SIs: check that the same property is not recorded twice.                                                                                                                          |
| The same property secures another active loan                                                                           | TYPE_OF_CHARGE is left blank and asked for.                                                                                                                                                        |
| A title report file is misnamed, misfiled or scanned (no text)                                                          | Listed in the Issues sheet; the file is not used.                                                                                                                                                   |
| CERSAI rejects an SI you already filed                                                                                  | In`Output data/Filed register.xlsx`, change its Status from `Filed` to `Pending` (or delete its row). It will come back in the next batch.                                                    |
| Step 2 is run again after the final Excel was edited by hand                                                            | The edited file is kept as a`(backup ...)` copy.                                                                                                                                                  |
| "Could not open or save ... close it in Excel"                                                                          | Close the file in Excel and run the step again.                                                                                                                                                     |
| "Something unexpected went wrong"                                                                                       | Details are in`Output data/error log.txt`.                                                                                                                                                        |

Nothing is ever overwritten: a replaced file is kept with a `(backup ...)` tag, and every run of step 1
creates a new batch folder. Do not edit `batch data (do not edit).json` in the batch folder: steps 2
and 3 rely on it.

## Settings

`cersai_tool/settings.py` holds the fixed choices, for example the entity code in the file name
(`J8084`), the type of charge for a property that secures no other loan, whether to put the loan account in NARRATION, whether to use
the Global collateral No as ASSET_UNIQUE_ID, the name-matching thresholds, and the 10% area tolerance.
The charge-holder (CHG) row is fixed in `cersai_tool/rows/chg_row.py`.

## Testing with the dummy data

`Dummy input data/` has the same layout as `Input data/`, with made-up people and companies covering
the kinds of data found in the real extracts and title reports (`Dummy data cases.md` lists them).
The dummy COLLATERAL has the hand-filled columns, filled for every collateral and every individual
customer, as the team will fill them.
After the hand-made cases come ordinary generated records, so that there are 304 collaterals in all:
the `.DAT` file gets 300 SIs, because the 4 third-party mortgages among the hand-made cases are left out.
Each one has its CKYC rows and a title report. All loans are active (Loan Status A). To change the size,
edit `COLLATERALS` in `dev_tools/make_dummy_data.py` and recreate the data.
Add `--dummy` to any step; results go to `Dummy output data/`.

To make the complete `.DAT` file for the CERSAI approval submission, run these from the project folder:

```
.venv\Scripts\python.exe -m cersai_tool run1 --dummy
.venv\Scripts\python.exe dev_tools\fill_dummy_required_details.py
.venv\Scripts\python.exe -m cersai_tool run2 --dummy
.venv\Scripts\python.exe -m cersai_tool run3 --dummy
```

The second command types into the required-details file what a person would type: the values the
hand-made test cases leave out of the extracts on purpose (a gender CKYC leaves blank, a PAN, a
registration number...) and the answers to the questions in Check and correct. It takes them from `Dummy input data/Dummy required details answers.xlsx`. Step 3 then writes a
`.DAT` file with 300 SIs in `Dummy output data/<date>/`.

To recreate the dummy data: `.venv\Scripts\python.exe dev_tools\make_dummy_data.py`.

## Code layout

| Path                             | Contents                                                                                                                                                                                                              |
| -------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cersai_tool/rows/`            | One file per row type (`fh_row`, `si_row`, `imm_row` + `imm_location`, `bor_row` + `bor_address`, `chg_row`), one function per column, each with a plain-English comment on where the value comes from. |
| `cersai_tool/rules/`           | Shared rules: text cleaning, dates, identifiers, address reading, PIN / district / city lookup, owner-name matching.                                                                                                  |
| `cersai_tool/sources/`         | Reading the LMS extracts, the CERSAI masters, the title reports, and the required-details file.                                                                                                                       |
| `cersai_tool/output/`          | The Excel layout, the step 3 checks and the`.DAT` writer.                                                                                                                                                           |
| `cersai_tool/steps/`           | The three steps.                                                                                                                                                                                                      |
| `dev_tools/make_dummy_data.py` | Creates the dummy input data.                                                                                                                                                                                         |
