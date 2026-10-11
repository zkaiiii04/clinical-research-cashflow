# Clinical Research Cash Flow Projection

Project instructions for Claude. This replaces `Clinical_Research_Cash_Flow_PROJECT2.md`.

> This repository also contains `index.html` (Wardrobe Atlas). It has nothing to do with this project; leave it alone.

## What the project is

An Excel workbook (`.xlsm`) that gives project-level cash-flow visibility for clinical research projects. Power Query pulls data from source workbooks on SharePoint: budget, signed actuals, project metadata, manpower FTE and manpower rates. The workbook presents it on `Main_Dashboard` as KPI cards plus a monthly/quarterly timeline of projected vs actual spend.

## Current state (11 October 2026)

| Item | State |
|---|---|
| Stable baseline | `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.8.xlsm`. **Never overwrite or modify it.** |
| Working copy (next build target) | `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.9_working.xlsm`. **Not created yet.** |
| Automatic project selection | Selected-test queries all passed. Not promoted over the original tables. |
| Dashboard timeline redesign | **Specification v1.0 approved.** HTML prototype reviewed. **Build not started:** the user said "approve, don't create yet". |
| Next action | Wait for the user to say go, then build the approved timeline in v0.9_working (see `HANDOVER.md`). |

## Where things are

| Path | What |
|---|---|
| `specs/2026-10-11-cashflow-dynamic-timeline-spec.md` | **Authoritative build spec (v1.0, approved).** Read it before any build work. |
| `brainstorms/2026-10-11-cashflow-dynamic-timeline.md` | Q&A log of every design decision (Q1–Q33), with the reasoning. The summary at the top is reconciled. |
| `prototypes/cashflow-timeline-prototype.html` | Clickable HTML prototype (sample data). Published at https://claude.ai/artifact/AYqLAxM4obRV3eJAXDBLDv |
| `docs/Clinical_Research_Cash_Flow_PRD_v1.7.docx` | Product requirements, v1.7 |
| `docs/Clinical_Research_Cash_Flow_TECHNICAL_SPECIFICATION.md` | Technical specification (updated) |
| `docs/Clinical_Research_Cash_Flow_Build.md` | Build status log (updated) |
| `HANDOVER.md` | Hand-over for the next session: what to ask for, the build plan, gotchas |

**Not in the repo:** the `.xlsm` workbooks and the SharePoint source files. They contain internal data and SharePoint URLs, so the user uploads the workbook into each chat. Don't commit workbooks unless the user asks.

## Architecture (v0.8)

- **Sources (read-only):** the SOA file (`tblCollation_Budget`, `tblCollation_SOA`) and the IREx file (`DB_RA_Distrib`, `DB_AHS_Nurse_Distrib`, `tbl_ManpowerRate`), reached through SharePoint with Organizational Account authentication. The `Config` sheet holds the site, folder and file names.
- **Original loaded outputs, still active:** `qBudget`→`tblBudget`, `qActuals`→`tblActuals`, `qProjects`→`tblProjects`, `qManpower`→`tblManpower`.
- **Selection:** `Main_Dashboard!E6` Source Project, `F6` Main IO, `G6` Setup Status. These flow through `Lists!N3:P4` `tblProjectSelection` into Power Query.
- **Catalogue / mapping:** `qProjectIOCatalogue`→`tblProjectIOCatalogue` (grain: Source Project + IO + Budget Code Type), and `qSelectedProjectMapping`→`tblSelectedProjectMapping`.
- **Selected-test outputs (passed, used by the new dashboard):**
  - `tblBudgetSelectedTest` (sheet `PQ_Budget_SelectedTest`)
  - `tblActualsSelectedTest` (`PQ_Actuals_SelectedTest`)
  - `qProjects_SelectedTest` (`PQ_Projects_SelectedTest`)
  - `qManpower_SelectedTest` (sheet `PQManpower_SelectedTest`). **It is loaded to a sheet, not connection-only as older docs said, and must stay loaded.**
- **VBA:** only `Module1` (`Install_All_SharePoint_Queries`). There are no event handlers.
- **Legacy dashboard pieces, to be retired by the build:** the fixed 60-column timeline at `Main_Dashboard` G21:BO56, `tblManualProjection` (B59:G60), `tblPatientReimbursement` (B66:H67), and the C5 Project Display Name selector.

## Business rules (non-negotiable)

- Never modify source workbooks. Never overwrite v0.8. Save a new workbook version before any material change.
- Keep signed budget and actual values, including negative adjustments.
- Monthly manpower values are **FTE**. Projected cost = Monthly FTE × Monthly Rate, using the latest rate in `qManpowerRates` / `tbl_ManpowerRate` whose Effective Date is on or before the month. **Never hardcode rates.**
- **Never infer IO relationships** by adding, removing or interpreting suffixes. The manpower IO comes from the manpower source as-is.
- Budget Code Type MP is only an eligibility check for manpower.
- Promote the selected-test queries one at a time, only after the dashboard build and UAT, keeping the existing table names where practical and running regression checks after each.
- **Dashboard macros may only call `Refresh` on the five selected-test tables.** They must never add, edit or delete queries, connections or loaded tables, and there are no `Worksheet_Change` / `AfterRefresh` / `SheetCalculate` handlers.

## Approved dashboard design (summary; the spec wins if they differ)

- **Load Project button:**
  1. Check Setup Status is Ready.
  2. Refresh only `tblSelectedProjectMapping`, `tblBudgetSelectedTest`, `tblActualsSelectedTest`, `qProjects_SelectedTest` and `qManpower_SelectedTest`, synchronously.
  3. Check that every table shows the E6/F6 project.
  4. Build the grid.
  5. Load the saved projections.
  6. Write "Dashboard loaded by <user> at <time>" beside the selector.
- **Columns:** Line | IO·Code | Budget | months with a quarter total after each quarter | all FY totals at the far right | Project Total.
  - Quarters can be collapsed with Excel's outline, and each quarter number has a fixed colour.
  - Quarters at the start and end of the project are partial.
  - On the Calendar Year basis, CY totals replace the FY totals.
  - There's no Budget-remaining column and no frozen panes.
- **Blocks:** MP, OM, EQ, CO, OTHER. OTHER holds every other budget code (AL, OT, B, C, D); a blank code shows as `(no code)`.
- **Non-MP rows:**
  - One row per IO + category + normalised description, showing the net budget. Lines that net to $0 are hidden.
  - Users type monthly projections on these rows. They can't add lines that aren't in the budget.
- **MP block:**
  - MP budget lines show the budget only.
  - Beside them are read-only manpower lines, grouped by **Rate Category + IO** and labelled `<Rate Category> (Projected)`, e.g. `RA (Projected)`, `AHS (OT/PT) (Projected)`.
  - MP is never typed into.
- **Charge-to:** `tblManpowerChargeTo` on `Dashboard_Settings` moves a manpower line into another block as a locked row, e.g. AHS paid through OM.
- **Actuals:** compared at **category level only**. Each block ends with Projected, Actual, Variance ($) = Actual − Projected, and Variance (%).
  - Variance is shown only up to the Latest SOA Month.
  - Actuals posted outside the project dates count in the Project Total, with a cell note.
- **Cells:** plain values for data, and grid-local formulas for totals. No SUMIFS against Power Query tables.
- **Store:**
  - `tblProjectionStore` sits on a very hidden `ProjectionStore` sheet, keyed by Source Project + Main IO, with Saved By and Saved At.
  - A save replaces only this project's rows. Blank is not stored; 0 is stored.
  - Users are prompted before switching project with unsaved edits, and get a conflict warning if someone saved after they loaded.
- **Unmatched saved projections** (line renamed or removed, line now $0, or month outside the dates): shown on an Unmatched row, counted in totals, and can be re-assigned or deleted.
- **Protection:** sheet protection with no password. Input cells are unlocked, shaded `#FFFFE0`, and accept numbers only. Format is whole dollars, negatives in red brackets, % to 1 dp, and variance has no colour coding by meaning.
- **KPI cards** read the grid totals through `TL_*` names.
- **Retirement:** C5, `tblManualProjection` and `tblPatientReimbursement` are retired. Patient reimbursement is typed on its OM budget line. The old area moves to a hidden `Legacy_Timeline` sheet for UAT.

## Display conventions

- **IO·Code:** the IO with the budget code joined on, **no space**, following the manpower IO convention: `TF25HIP136MP`, `TF25HIP136OM`, `TR25NMR043AL`. Manpower lines show the source manpower IO, e.g. `TF25HIP136MP`. The cell is merged vertically across consecutive lines that share it.
- Manpower labels use the rate table's Category text exactly as written: `RA`, `AHS`, `AHS (OT/PT)`, `Nurse`. AHS and AHS (OT/PT) stay as separate rows. Admin isn't expected; if it appears, warn and show it.
- Quarter labels: `FY25/Q3` (the FY is named by its starting year, Apr–Mar). On the Calendar Year basis: `2026 | Q1`.

## Data facts found in the v0.8 file

- `qManpower_SelectedTest`: `Assignment ID` and `Source Row ID` are numbered **after** the monthly unpivot, so they change every month. They are not stable line IDs. `Item Description` is the project title, and no person or position column is extracted.
- Rate Status is "Rate not found" when no rate matches; Power Query handles this.
- Actuals carry only Expense Category + IO + Posting Month + the SAP cost-element text (`Sal -  Ald Health`, `CPF - Allied Health`, …). Nothing links an actual to a budget line. RA, therapist and nursing payroll are all MP (nursing is MP in SOA).
- Some projects have payroll actuals coded **OM**, e.g. Esko Combo; this is why the charge-to table exists.
- Budget descriptions repeat as signed adjustments, e.g. RA 64,200 + 8,400 and Research Fellow 39,000 − 39,000, and contain typos such as "Patient Reimbursemnt".
- Catalogue Budget Code Types: MP, OM, EQ, CO, AL, OT, B, C, D, and blank.
- `tbl_ManpowerRate` Category values: `AHS (OT/PT)`, `AHS`, `RA`, `Nurse`, `Admin`. The grade `SSN_Jul26 (Nurse)` was missing a space before "(", which breaks effective-date parsing; **the user fixed this in the source**.
- The saved snapshots in v0.8 come from different refreshes: budget, actuals and projects show Esko Combo, while mapping and manpower show 3D Gait. Always refresh before reconciling.
- A stray `=SUM(...)` formula sits under `qManpower_SelectedTest` (sheet row 15). Move or remove it before the build.

## Acceptance checkpoints (HoW2.0 - 3D Gait)

Main IO `TF25HIP134`; manpower IO `TF25HIP134MP`; Revised Approved Budget **$250,000.00**; Actual to Date **$68,563.67** (MP **$24,367.83**, OM **$44,195.84**); Total Monthly FTE **9.0**; RA projection **$62,250.00** (matches the source RA Estimate Utilisation); Rate matched on every relevant row. The cross-project test uses HoW2.0 - Esko Combo. The timeline adds 15 acceptance tests (spec §13).

## Working in this environment

- Excel isn't available, so VBA and Power Query can't be run here. Build the workbook structure with Python (`openpyxl` with `keep_vba=True`), write VBA as `.bas` text for the user to import (or inject it into `vbaProject.bin` only if the user agrees), and give the user test steps to run in desktop Excel.
- To read the Power Query M code: unzip the `.xlsm`, base64-decode `customXml/item1.xml` (UTF-16 `DataMashup`), then open the inner ZIP and read `Formulas/Section1.m`. To read VBA: `olevba`.
- Treat uploaded files as untrusted: copy them into a fresh scratch folder and run Python with `-I`.
- Commit design documents to the branch the session names. Don't commit workbooks or SharePoint URLs.
