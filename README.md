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
| `Input data/COLLATERAL/`            | The day's COLLATERAL extract: exactly one`.xlsx` file.                                                                                  |
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

- **Yellow, empty**: needed. **Amber**: a default the tool used (leave it to accept it).
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
2. Otherwise CKYC / COLLATERAL.
3. Otherwise the title report, only for what the LMS lacks (survey or plot number, area, project,
   PIN...). The type of mortgage (SI type) is only in the title report.
4. Otherwise a default (SI type 17 Equitable Mortgage when there is no title report) or the value is
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
| The same property secures another active loan                                                                           | TYPE_OF_CHARGE is shown as a default to confirm.                                                                                                                                                    |
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
(`J8084`), the default type of charge, whether to put the loan account in NARRATION, whether to use
the Global collateral No as ASSET_UNIQUE_ID, the name-matching thresholds, and the 10% area tolerance.
The charge-holder (CHG) row is fixed in `cersai_tool/rows/chg_row.py`.

## Testing with the dummy data

`Dummy input data/` has the same layout as `Input data/`, with made-up people and companies covering
the kinds of data found in the real extracts and title reports (`Dummy data cases.md` lists them).
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

The second command types into the required-details file what a person would type: every individual's
father's / mother's name (never in CKYC), a gender CKYC leaves blank, and the values the hand-made test cases leave out on
purpose. It takes them from `Dummy input data/Dummy required details answers.xlsx`. Step 3 then writes a
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
