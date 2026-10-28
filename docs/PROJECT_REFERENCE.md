# Clinical Research Cash Flow: Project Reference

Detailed reference moved out of `CLAUDE.md` on 12 Oct 2026 so that `CLAUDE.md` stays a short directory plus rules. Where this file and the build spec (`specs/2026-10-11-cashflow-dynamic-timeline-spec.md`) differ, the spec wins.

## 1. Architecture (v0.8 queries; v0.9 dashboard)

- **Sources (read-only):** the SOA file (`tblCollation_Budget`, `tblCollation_SOA`) and the IREx file (`DB_RA_Distrib`, `DB_AHS_Nurse_Distrib`, `tbl_ManpowerRate`), reached through SharePoint with Organizational Account authentication. The `Config` sheet holds the site, folder and file names.
- **Original loaded outputs, still active:** `qBudget`→`tblBudget`, `qActuals`→`tblActuals`, `qProjects`→`tblProjects`, `qManpower`→`tblManpower`.
- **Selection:** Source Project, Main IO and Setup Status. These flow through `Lists!N3:P4` `tblProjectSelection` into Power Query. **The user moved them from `E6:G6` to `E5:G5` in their working copy on 12 Oct 2026.** The VBA reads the cells from constants in `modTimeline` (`SEL_PROJECT`, `SEL_MAINIO`, `SEL_STATUS`, plus `NOTICE_CELL` E7, `BASIS_CELL` C6, `REFRESHED_CELL` C10). The build script and the release workbook still have the selector in row 6, so a rebuild must move it, or the constants must be set back.
- **Catalogue / mapping:** `qProjectIOCatalogue`→`tblProjectIOCatalogue` (grain: Source Project + IO + Budget Code Type), and `qSelectedProjectMapping`→`tblSelectedProjectMapping`.
- **Selected-test outputs (passed, used by the new dashboard):**
  - `tblBudgetSelectedTest` (sheet `PQ_Budget_SelectedTest`)
  - `tblActualsSelectedTest` (`PQ_Actuals_SelectedTest`)
  - `qProjects_SelectedTest` (`PQ_Projects_SelectedTest`)
  - `qManpower_SelectedTest` (sheet `PQManpower_SelectedTest`), loaded by the query of the same name. **It is loaded to a sheet, not connection-only as older docs said, and must stay loaded.** v0.9_working keeps these v0.8 names: an XML rename to `tblManpowerSelectedTest` broke the query link (incident 2026-10-11). The VBA reads the name from `modTimeline.MANPOWER_TABLE`.
- **VBA:** only `Module1` (`Install_All_SharePoint_Queries`). There are no event handlers.
- **Legacy dashboard pieces (v0.8), retired in v0.9_working and kept on the hidden `Legacy_Timeline` sheet for UAT:** the fixed 60-column timeline at `Main_Dashboard` G21:BO56, `tblManualProjection` (B59:G60), `tblPatientReimbursement` (B66:H67), and the C5 Project Display Name selector.

## 2. Approved dashboard design (summary; the spec wins if they differ)

- **Project dates:** from `qProjects_SelectedTest` (IREx, IO Code Type "Main"). `tblSelectedProjectMapping` dates are the SOA budget LOA dates and only a fallback; Load warns when they differ.
- **Load Project button:**
  1. Check Setup Status is Ready.
  2. Refresh only `tblSelectedProjectMapping`, `tblBudgetSelectedTest`, `tblActualsSelectedTest`, `qProjects_SelectedTest` and `qManpower_SelectedTest`, synchronously.
  3. Check that every table shows the selected project (selector cells: `modTimeline` `SEL_*` constants).
  4. Build the grid.
  5. Load the saved projections.
  6. Write "Dashboard loaded by <user> at <time>" beside the selector.
- **Columns:** Line (B, overflowing to E) | IO·Code (F) | Budget (G) | months from H with a quarter total after each quarter | all FY totals at the far right | Project Total | Remarks (typed per line, saved per project).
  - Quarters can be collapsed with Excel's outline, and each quarter number has a fixed colour.
  - Quarters at the start and end of the project are partial.
  - On the Calendar Year basis, CY totals replace the FY totals.
  - There's no Budget-remaining column. The macro doesn't touch frozen panes; the user's own panes are kept.
  - Rows 2–19 are grouped on every build (collapse with the +/− beside row 1).
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
  - Variance is shown only up to the Latest SOA Month. After it nothing has been posted yet, so a variance would show the whole future plan as underspend. Actual/Variance cells for periods wholly after the Latest SOA Month are left blank (§7 below).
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

## 3. Display conventions

- **IO·Code:** the IO with the budget code joined on, **no space**, following the manpower IO convention: `TF25HIP136MP`, `TF25HIP136OM`, `TR25NMR043AL`. Manpower lines show the source manpower IO, e.g. `TF25HIP136MP`. The cell is merged vertically across consecutive lines that share it.
- Manpower labels use the rate table's Category text exactly as written: `RA`, `AHS`, `AHS (OT/PT)`, `Nurse`. AHS and AHS (OT/PT) stay as separate rows. Admin isn't expected; if it appears, warn and show it.
- Quarter labels: `FY25/Q3` (the FY is named by its starting year, Apr–Mar). On the Calendar Year basis: `2026 | Q1`.

## 4. Data facts and known data risks

- `qManpower_SelectedTest`: `Assignment ID` and `Source Row ID` are numbered **after** the monthly unpivot, so they change every month. They are not stable line IDs. `Item Description` is the project title, and no person or position column is extracted.
- Rate Status is "Rate not found" when no rate matches; Power Query handles this.
- Actuals carry only Expense Category + IO + Posting Month + the SAP cost-element text (`Sal -  Ald Health`, `CPF - Allied Health`, …). Nothing links an actual to a budget line. RA, therapist and nursing payroll are all MP (nursing is MP in SOA).
- Some projects have payroll actuals coded **OM**, e.g. Esko Combo; this is why the charge-to table exists.
- **Manpower lines on a non-MP IO land in MP (open, 28 Oct 2026).** Esko Combo's manpower service fee is on `TF25HIP136OM` in the IREx source (correct), but the macro puts every manpower line in MP unless `tblManpowerChargeTo` maps it, and charge-to doesn't apply to "Rate not found" lines. Workaround: a charge-to row for `TF25HIP136OM` → OM. Fix options: `docs/incidents/2026-10-28-esko-combo-om-service-fee-in-mp.md`.
- Budget descriptions repeat as signed adjustments, e.g. RA 64,200 + 8,400 and Research Fellow 39,000 − 39,000, and contain typos such as "Patient Reimbursemnt".
- **Project dates:** `tblProjectIOCatalogue` / `tblSelectedProjectMapping` Project Start/End are the min/max of the **SOA budget LOA dates**, not the project dates. The authoritative project dates are in `qProjects_SelectedTest` (IREx). For tDCS they differ (SOA 31 Dec 2026, project 30 Sep 2027). The manpower `Tag Start`/`Tag End` are each person's LOA.
- Catalogue Budget Code Types: MP, OM, EQ, CO, AL, OT, B, C, D, and blank.
- `tbl_ManpowerRate` Category values: `AHS (OT/PT)`, `AHS`, `RA`, `Nurse`, `Admin`. The grade `SSN_Jul26 (Nurse)` was missing a space before "(", which breaks effective-date parsing; **the user fixed this in the source**.
- The saved snapshots in v0.8 come from different refreshes: budget, actuals and projects show Esko Combo, while mapping and manpower show 3D Gait. Always refresh before reconciling.
- A stray `=SUM(...)` formula sat under the manpower table (sheet row 15). It is removed in v0.9_working.
- **Stale cross-workbook formulas in the SOA file (data risk, found 12 Oct 2026).** The SOA file fills some columns, such as the project name (and the LOA dates that feed the catalogue), with **XLOOKUP formulas that point at another workbook**. Power Query reads the SOA file **as saved on SharePoint**: it gets each formula's *stored value from the last save*, not what Excel shows on screen. Having both files open doesn't help. If the SOA file was saved before its external links updated (links not refreshed, the source changed later, saved in Excel for the web or mobile, or not yet synced), Power Query reads stale values. Effects:
  - stale **LOA dates** shift the catalogue and mapping dates (the cause of the tDCS end date, now guarded because the timeline uses IREx dates);
  - a stale **project name** makes budget and actual rows **drop out** of the selected-test joins without any error.
  - **User routine:** in the SOA file, Data → Edit Links → Update Values (or open the source workbook), check the values, save, let SharePoint sync, then click Load Project.
  - **Post-UAT option:** have Power Query read the lookup's source table directly and do the match in M (no dependency on saved formula values), or paste that column as values after each update. Not started: it changes queries, so it follows the one-at-a-time promotion rule.

## 5. Acceptance checkpoints (HoW2.0 - 3D Gait)

Main IO `TF25HIP134`; manpower IO `TF25HIP134MP`; Revised Approved Budget **$250,000.00**; Actual to Date **$68,563.67** (MP **$24,367.83**, OM **$44,195.84**); Total Monthly FTE **9.0**; RA projection **$62,250.00** (matches the source RA Estimate Utilisation); Rate matched on every relevant row. The cross-project test uses HoW2.0 - Esko Combo. The timeline adds 18 acceptance tests (spec §13).

## 6. Working techniques

- **Read the Power Query M code:** unzip the `.xlsm`, base64-decode `customXml/item1.xml` (UTF-16 `DataMashup`), open the inner ZIP and read `Formulas/Section1.m`. The query metadata (e.g. `FillTarget`) follows the package part in the same binary.
- **Read VBA:** `olevba <file>`.
- **Read cell values and formulas:** openpyxl, read-only (`data_only=True` for cached values). Never save with it.
- **Change the workbook:** edit the package XML (see `build/build_v09_working.py`), then run `build/verify_v09_working.py <v0.8> <output>`.
- **Grid logic:** `tests/grid_model.py <v0.8 or v0.9 workbook> <scratch dir>` ports the macro's grid to Python and recalculates it in LibreOffice (a throwaway copy only).
- **After an Excel save** (test 13): `build/compare_pq.py <v0.8> <saved v0.9>`.
- **Docx edits:** unzip, edit `word/document.xml`, rezip, validate (docx skill); keep the existing styles.

## 7. Why Actual and Variance are blank after the Latest SOA Month

The **Latest SOA Month** is the last month that has posted actuals in SOA (`MAX(tblActualsSelectedTest[Posting Month])`, shown in C9; for tDCS it's Aug 2026). For months after it, **no actuals have been posted yet**. That doesn't mean nothing was spent; the spending just hasn't been recorded.

If the dashboard calculated Variance = Actual − Projected for those months:
- Actual would be 0 (nothing posted), so every future month would show Variance = −Projected.
- Quarter, FY and Project Total variances would add up all those negatives, and the project would look **heavily underspent**. For tDCS, every planned dollar from Sep 2026 to Sep 2027 would count as "underspend".
- That figure would be wrong and would change every month as SOA catches up. It would hide the real question: how is spending tracking against plan **for the months we actually have data for**?

So the rule (spec §14.2, accepted by the user on 11 Oct 2026 and reconfirmed on 12 Oct 2026: **keep blank**) is:
- **Months, quarters and FY columns that lie entirely after the Latest SOA Month** get no Actual, Variance ($) or Variance (%) cells. They stay blank, not 0.
- **A period that straddles the Latest SOA Month** (e.g. FY26/Q2 = Jul–Sep 2026 with SOA at Aug) shows Actual for its posted months, and Variance compares those actuals only with the projections **up to the Latest SOA Month**.
- **The Project Total** Variance is actuals to date minus projections to date. Variance (%) = Variance ÷ projections to date.
- **The Projected rows** always run for the whole project, so the full plan stays visible.
- **The cut-off moves forward on its own:** when the next month's SOA is posted and the user clicks Load Project, that month's (and quarter's) Actual and Variance appear.

Code: `MonthLE` / `YearLE` and `WriteBlockSubtotals` / `WriteProjectTotal` in `vba/modTimeline.bas`. Tests: the "No variance after the Latest SOA Month" and "variance uses projections to SOA only" checks in `tests/grid_model.py`.
