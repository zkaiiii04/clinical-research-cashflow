# Handover: Clinical Research Cash Flow Dashboard Timeline

**From session:** 11 October 2026 (design discovery → spec approval)
**Where:** `main` branch of `zkaiiii04/clinical-research-cashflow` (renamed from `wardrobe-atlas` on 11 Oct 2026; GitHub redirects the old name). Design work was merged in PR #1; the unrelated Wardrobe Atlas page was removed in PR #2.
**Status:** spec **v1.0 approved** (revision v1.0.4). v0.9_working is built and **in UAT**. Both 11 Oct issues are closed and confirmed in desktop Excel. 12 Oct: the selector moved to row 5 (VBA updated), and the user asked why variance is blank after the Latest SOA Month (explained in §0.1). Next: finish UAT (§0), then promote the selected-test queries.

---

## 0. START HERE: UAT in progress (as of 12 Oct 2026)

**Where things stand:**
- v0.9_working (spec v1.0.3) is built. `release/CashFlow_v0.9_package.zip` on `main` has the workbook and its SHA-256 in `README_FIRST.txt`, plus the current `vba/` modules (CRLF; import them from the zip, not GitHub "Raw").
- **Timeline end-date defect: closed.** The timeline and C7/C8 now use the IREx project dates in `qProjects_SelectedTest`. The SOA budget LOA dates in `tblSelectedProjectMapping` are only a fallback, and Load warns when they differ. The user confirmed tDCS works in desktop Excel. The stale 31 Dec 2026 came from a source-file column Power Query reads (SOA LOA end). See `docs/incidents/2026-10-11-v0.9-project-end-date.md`.
- **queryTable9 repair: closed.** The rebuilt copy, which keeps the v0.8 manpower names, opens and loads without a repair prompt (§0b).

**12 Oct 2026 changes:**
- **Data risk: stale XLOOKUP values in the SOA file.** The SOA file fills some columns, such as the project name (and the LOA dates behind the catalogue), with XLOOKUPs to another workbook. Power Query reads the file **as saved on SharePoint**, so it gets each formula's stored value from the last save, not what's on screen. Having both files open doesn't help. This is the likely root cause of the stale tDCS end date.
  - A stale project name silently drops rows from the selected-test joins.
  - **User routine:** in the SOA file, Data → Edit Links → Update Values, check the values, save, let it sync, then Load Project.
  - **Post-UAT option:** read the lookup source in Power Query and match in M, or paste that column as values.
  - Details: `docs/PROJECT_REFERENCE.md` §4.
- **Selector moved to row 5 (user, in their working copy).** The VBA now reads the selector cells from constants at the top of `vba/modTimeline.bas` (`SEL_PROJECT` = E5, `SEL_MAINIO` = F5, `SEL_STATUS` = G5; `NOTICE_CELL` = E7, `BASIS_CELL` = C6, `REFRESHED_CELL` = C10). `modStore` and `modSetup` use them too.
  - The workbook formulas that must point at the new cells: `Lists!N4:P4` (`=Main_Dashboard!E5/F5/G5`), the Setup Status formula (now G5, referencing E5/F5), and the data validations on E5 (`CatalogueProjectNames`) and F5 (`SelectedProjectIOCodes`). These follow automatically if the cells were cut and pasted.
  - **Mismatch to fix before any rebuild:** `build/build_v09_working.py` and the release workbook still put the selector in row 6. Either teach the build the row-5 layout (ask the user for their current workbook to copy the exact layout) or set the constants back to row 6 for a fresh install.
- **Blank Actual/Variance after the Latest SOA Month** (tDCS: FY26/Q3 onwards and the FY27 Total, in every category): this is the agreed rule, not missing formulas. See §0.1 for the explanation to give the user. The Projected rows are complete. **Open decision for the user:** keep as is, or show Actual as 0 in future periods while Variance stays blank.

### 0.1 Why Variance (and Actual) is left blank after the Latest SOA Month

The **Latest SOA Month** is the last month that has posted actuals in SOA (`MAX(tblActualsSelectedTest[Posting Month])`, shown in C9; for tDCS it's Aug 2026). For months after it, **no actuals have been posted yet**. That doesn't mean nothing was spent; the spending just hasn't been recorded.

If the dashboard calculated Variance = Actual − Projected for those months:
- Actual would be 0 (nothing posted), so every future month would show Variance = −Projected.
- Quarter, FY and Project Total variances would add up all those negatives, and the project would look **heavily underspent**. For tDCS, every planned dollar from Sep 2026 to Sep 2027 would count as "underspend".
- That figure would be wrong and would change every month as SOA catches up. It would hide the real question: how is spending tracking against plan **for the months we actually have data for**?

So the rule (spec §14.2, accepted by the user on 11 Oct 2026) is:
- **Months, quarters and FY columns that lie entirely after the Latest SOA Month** get no Actual, Variance ($) or Variance (%) cells. They stay blank, not 0.
- **A period that straddles the Latest SOA Month** (e.g. FY26/Q2 = Jul–Sep 2026 with SOA at Aug) shows Actual for its posted months, and Variance compares those actuals only with the projections **up to the Latest SOA Month**.
- **The Project Total** Variance is actuals to date minus projections to date. Variance (%) = Variance ÷ projections to date.
- **The Projected rows** always run for the whole project, so the full plan stays visible.
- **The cut-off moves forward on its own:** when the next month's SOA is posted and the user clicks Load Project, that month's (and quarter's) Actual and Variance appear.

Code: `MonthLE` / `YearLE` and `WriteBlockSubtotals` / `WriteProjectTotal` in `vba/modTimeline.bas`. Tests: the "No variance after the Latest SOA Month" and "variance uses projections to SOA only" checks in `tests/grid_model.py`.

**Next session:**
1. Ask the user for their UAT results: the 16 tests in `tests/UAT_test_script.md`, including test 16 (tDCS project dates) and the Debug → Compile result. Fix what fails. If grid logic changes, keep `tests/grid_model.py` in step with `vba/modTimeline.bas` and re-run it (43 checks).
2. **Open source-data question (tDCS OT vs RA):** the manpower table shows RA = 43,575 (from `DB_RA_Distrib`) and AHS (OT/PT) = 11,150 (from `DB_AHS_Nurse_Distrib`), the reverse of what the user expected (budget OT 42,000 / RA 10,800). The staff group comes from which IREx sheet a row is on. Ask whether they checked IREx. If the OT person moves to the AHS sheet, their amount changes (priced at their AHS grade, not AH11 RA). This isn't a workbook bug unless the source proves otherwise.
3. After UAT sign-off: promote the selected-test queries one at a time (CLAUDE.md business rules) and delete `Legacy_Timeline`.
4. For test 13: `python3 -I build/compare_pq.py <v0.8> <user's saved v0.9>`.

## 0b. v0.9 repair error (queryTable9): closed 11 Oct 2026 (history)

**Fix session outcome (11 Oct 2026):**
- The rename is dropped: the manpower table and sheet keep the v0.8 names `qManpower_SelectedTest` / `PQManpower_SelectedTest`. The VBA reads the name from `modTimeline.MANPOWER_TABLE`.
- `build/verify_v09_working.py` now checks cross-part links (`check_links`) with no exemptions. It flags the first v0.9 (3 problems) and passes v0.8 and the rebuild.
- Rebuilt from v0.8 (SHA-256 `db3bdd5b…dc81ab0`) and repackaged as `release/CashFlow_v0.9_package.zip`. The workbook's SHA-256 is in its `README_FIRST.txt`.
- **Correction:** the shipped file *was* the build output. It says "Excel iOS" because **v0.8 itself was last saved on an iPhone/iPad**. Its PQ and VBA parts pass every check. The build no longer refuses such input; it prints the app and relies on the verifier.
- **Next:** ask the user whether the rebuilt file opened in desktop Excel with no repair prompt. If it did, close the incident and go to §9. If not, ask for the full repair log.

The original analysis, kept for the record:

When the user opened the v0.9 package workbook from SharePoint in desktop Excel, Excel said it had to repair it and logged: `Removed Part: /xl/queryTables/queryTable9.xml part. (External data range)`. Full analysis: **`docs/incidents/2026-10-11-v0.9-querytable9-repair.md`**.

**Preliminary root cause:**
1. **Definite:** the build renamed the sheet `PQManpower_SelectedTest` → `PQ_Manpower_SelectedTest` but left the hidden defined name `ExternalData_3` pointing at `PQManpower_SelectedTest!$A$1:$R$13`, a sheet that no longer exists. Excel can't place the query table and removes it. A clean rebuild with `build/build_v09_working.py` reproduces this.
2. **Likely:** the Power Query metadata still has `FillTarget = qManpower_SelectedTest`, while the table was renamed to `tblManpowerSelectedTest` by an XML edit.
3. **Contributing:** the `.xlsm` in `release/CashFlow_v0.9_package.zip` was re-saved by **Excel iOS** before it was zipped, so it isn't the verified build output. The verifier also exempted the renamed table and didn't check defined names, so it reported PASS on a broken file.

**Do this first in the next session:**
1. Ask the user:
   - not to save the repaired copy (and whether AutoSave already wrote it to SharePoint);
   - to upload v0.8 again;
   - whether they're happy to **keep the v0.8 names** (`qManpower_SelectedTest` / `PQManpower_SelectedTest`). This is recommended. They can rename later in the Excel UI, which keeps the links in step.
2. Remove the table/sheet rename from `build/build_v09_working.py`, and revert it in `vba/`, `build/compare_pq.py`, `tests/grid_model.py`, `CLAUDE.md` and the docs. Or, only if the user insists on the new names, also rewrite `ExternalData_3` and patch the DataMashup `FillTarget` (riskier).
3. Harden `build/verify_v09_working.py`:
   - defined-name sheet references must exist;
   - query-table ↔ table ↔ DataMashup `FillTarget` must agree;
   - no exemptions.
4. Rebuild from v0.8, verify, and repackage **the exact build output**, with a SHA-256 in `README_FIRST.txt`. Replace `release/CashFlow_v0.9_package.zip` (the current one is bad; don't hand it out).
5. Tell the user to open it **only in desktop Excel for Windows**, with no phone, tablet or web Excel before the macros are installed. Then resume UAT (§9).


## 1. Read these first, in this order

1. `CLAUDE.md`: goals, rules and a directory of the repo. Then `docs/PROJECT_REFERENCE.md`: architecture, design summary, display conventions, data facts and data risks.
2. `specs/2026-10-11-cashflow-dynamic-timeline-spec.md`: **the build spec (authoritative).** §10 covers the macros, §12 the KPI repointing and migration, §13 the 16 acceptance tests.
3. `brainstorms/2026-10-11-cashflow-dynamic-timeline.md`: why each decision was made (Q1–Q33). The summary at the top is reconciled; the later Q entries supersede earlier ones where they conflict, and those are marked.
4. `prototypes/cashflow-timeline-prototype.html`: what the user approved visually (https://claude.ai/artifact/AYqLAxM4obRV3eJAXDBLDv).

Supporting documents, updated this session: `docs/Clinical_Research_Cash_Flow_PRD_v1.8.docx`, `docs/Clinical_Research_Cash_Flow_TECHNICAL_SPECIFICATION.md`, `docs/Clinical_Research_Cash_Flow_Build.md`.

## 2. What to ask the user for at the start

- **Upload `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.8.xlsm`.** It isn't in the repo by design, and this session's container is gone.
- **Confirm they want the build to start now.** They approved the design but explicitly said not to create the workbook yet.
- **Agree how the VBA reaches the workbook:**
  - (a) Claude writes `.bas` and `.cls` files plus a Python-built `.xlsm` with the new sheets and tables, and the user imports the modules in Excel (VBE → File → Import). This is the safer route.
  - (b) Claude injects the VBA into `vbaProject.bin` directly. That's riskier, and only with the user's consent.
- **Optionally:** a fresh refresh of v0.8 in Excel before uploading, because the saved snapshots mix projects (§5).

## 3. What was decided (one screen)

- **Platform:** one shared `.xlsm` on SharePoint, desktop Excel. Several users at once, but only one drives the dashboard at a time.
- **Driver:** E6 Source Project + F6 Main IO. The timeline reads `tblSelectedProjectMapping`, `tblBudgetSelectedTest`, `qManpower_SelectedTest` and `tblActualsSelectedTest`. C5 is retired. The KPI cards are repointed to the grid totals.
- **Load Project:** checks Ready → refreshes only those 5 tables synchronously → checks they all match E6/F6 → builds the grid sized to the project → loads saved projections → writes the "loaded by" notice. **Never edits queries**, and there are no event handlers.
- **Columns:** Line | IO·Code | Budget | months + a quarter total after each quarter (collapsible outline, fixed colour per quarter number) | all FY/CY totals at the far right | Project Total. No Budget-remaining column, no frozen panes.
- **IO·Code:** the IO with the budget code joined, no space (`TF25HIP136OM`). It's merged vertically across consecutive identical values.
- **Blocks:** MP, OM, EQ, CO, OTHER (AL, OT, B, C, D; blank shows as `(no code)`).
- **Non-MP rows:** budget lines merged by IO + category + normalised description; net budget; $0 lines hidden; typed monthly projections; no ad-hoc lines.
- **MP block:** budget lines (budget only), shown side by side with read-only manpower lines grouped by **Rate Category + IO**, labelled `RA (Projected)`, `AHS (OT/PT) (Projected)`, `AHS (Projected)`, `Nurse (Projected)`. MP is never typed into. "Rate not found" rows are shown and flagged.
- **Charge-to:** `tblManpowerChargeTo` on `Dashboard_Settings` moves a manpower line into another block, e.g. OM.
- **Actuals:** category level only. Projected / Actual / Variance $ (Actual − Projected) / Variance %. Variance only up to the Latest SOA Month. Actuals outside the project dates go into the Project Total with a note.
- **Store:** very hidden `tblProjectionStore`, keyed by Source Project + Main IO. Save replaces only that project's rows. Saved By/At stamps. Prompt before switching with unsaved edits. Conflict warning. Blank isn't stored; 0 is.
- **Unmatched saved rows:** shown, counted in totals, and can be re-assigned or deleted.
- **Cells:** values for data, grid-local formulas for totals.
- **Protection:** no password; inputs shaded `#FFFFE0`, numbers only. Format: whole dollars, red brackets for negatives, no colour coding of variance by meaning.
- **Retired:** `tblManualProjection` and `tblPatientReimbursement` (patient reimbursement is typed on its OM line). The old area moves to a hidden `Legacy_Timeline` sheet for UAT.

## 4. Build plan (done 11 Oct 2026, except steps 7-8, which the user runs in Excel; see §9)

1. Copy v0.8 to `..._v0.9_working.xlsm`. **Never write to v0.8.**
2. Housekeeping:
   - move or remove the stray `=SUM(qManpower_SelectedTest[...])` in `PQManpower_SelectedTest` row 15;
   - keep `qManpower_SelectedTest` loaded.
3. Sheets and tables:
   - `ProjectionStore` (very hidden) holding `tblProjectionStore`;
   - `Dashboard_Settings` holding `tblManpowerChargeTo`, with data validation on Charge To;
   - `Legacy_Timeline` (hidden): move G21:BO56 plus labels, `tblManualProjection` and `tblPatientReimbursement` there.
4. Names: `TL_Anchor` = `Main_Dashboard!$B$21`. `TL_TotalBudget`, `TL_TotalActual` and `TL_TotalProjected` are redefined on each build. Hidden session-state names: loaded key, loaded time, input snapshot.
5. VBA (new modules only; `Module1` untouched):
   - `modTimeline`: `LoadProject`, grid builder, outline, colours, formats, validation, protection, warnings.
   - `modStore`: `SaveProjections`, `ReassignUnmatched`, `DeleteUnmatched`, store read/write, conflict check.
   - `ThisWorkbook.Workbook_Open`: `EnableOutlining = True`, then `Protect UserInterfaceOnly:=True`.
   - Buttons (Form controls) beside the selector: Load Project, Save Projections, Re-assign unmatched, Delete unmatched.
6. KPI and info cells per spec §12: C7/C8 from the mapping table's Main IO row; C9 = MAX of the actuals' Posting Month; C14, F14 and F17 = `TL_*`; C17 and C19 are ratios. Retire C5.
7. Verify what can be verified here:
   - the `.xlsm` opens with openpyxl (`keep_vba=True`);
   - Power Query M is byte-identical before and after (compare `customXml/item1.xml` → `Formulas/Section1.m`);
   - the VBA compiles in principle (careful static review; no Excel here).
8. Give the user a **desktop-Excel test script** for the acceptance tests in spec §13 (now 16). The key figures for 3D Gait: Budget **$250,000.00**; Actual **$68,563.67** (MP **$24,367.83**, OM **$44,195.84**); MP Projected **$62,250.00**; 17 months Nov 2025–Mar 2027; 6 quarters; FY25/FY26 totals.
9. Commit the `.bas` files, the build script and the test script to the branch. Don't commit the workbook unless the user asks.

## 5. Gotchas found this session

- **Power Query reads saved files, not open ones.** Formula values in a source workbook (e.g. SOA XLOOKUPs to another file) are whatever was stored at its last save to SharePoint. Ask the user to update links, save and sync before refreshing.

- `Assignment ID` / `Source Row ID` in `qManpower_SelectedTest` change every month, because they're numbered after the unpivot. **Don't use them as line keys.** Group by Rate Category + IO.
- The saved snapshots in v0.8 come from different refreshes: budget, actuals and projects show Esko Combo, while mapping and manpower show 3D Gait. The macro's consistency check exists because of this.
- Budget output Project Group comes from *Project Display Name*, while manpower Project Group comes from *Source Project Name*. Each table has to be checked against the right field.
- Payroll actuals can be coded OM (e.g. Esko Combo `Sal -  Ald Health` under OM). This is why the charge-to table exists.
- Budget descriptions contain typos and trailing spaces ("Therapy Associate Service " vs "Therapy Associate Service"). Normalise before merging and keying.
- Outline +/− under sheet protection needs `EnableOutlining = True`, which Excel doesn't keep between sessions, so it must be set in `Workbook_Open`.
- `alert` / `confirm` don't work in the published HTML prototype, so it uses in-page dialogs. The real workbook uses `MsgBox`.
- The rate grade `SSN_Jul26 (Nurse)` spacing has been fixed in the source. If any "Rate not found" appears for nurses, check for grade-text spacing again.

## 6. Open / future (not blocking the build)

- MP parent/child view (budget line → manpower lines through an admin mapping). The user picked "side by side **for now**".
- Promote the selected-test queries over `tblBudget` / `tblActuals` / `tblProjects` / `tblManpower` after the build and UAT.
- Wider UAT, multi-user SharePoint testing, and deleting `Legacy_Timeline` after sign-off.

## 7. Environment notes

- GitHub pushes to this repo failed with 403 for part of the session, then worked once access was reconnected. If pushes fail again, the user reconnects GitHub at https://claude.ai/connect-github.
- Excel isn't available in the cloud container, so the build is done by script and tested by the user in desktop Excel.
- The user follows along on a phone. Keep replies short, and use numbered options when asking questions; their answers come in a shorthand like "a1, b1".

## 8. Suggested opening message for the next chat

> Continue the Clinical Research Cash Flow dashboard work in zkaiiii04/clinical-research-cashflow (main). Read `HANDOVER.md` (start at §0) and `CLAUDE.md`. Here are my UAT results and my current v0.9_working (selector now in row 5).

## 9. Build session (11 Oct 2026): what was done and what's next

- **Workbook:** `build/build_v09_working.py` builds v0.9_working from v0.8 by editing the package XML.
  - openpyxl was **not** used to save: an openpyxl save drops the Power Query DataMashup, the connections and the query tables.
  - `build/verify_v09_working.py` passes: the M code is identical, the 34 PQ parts are byte-identical, and `Module1` is untouched.
  - The workbook isn't in the repo. The user has the copy that was sent in chat; rebuild it from v0.8 with the script if needed.
- **VBA route (a):** `vba/modTimeline.bas`, `vba/modStore.bas` and `vba/modSetup.bas` are imported by the user. `vba/ThisWorkbook_code.txt` is pasted into ThisWorkbook. Then the user runs `Setup_Dashboard_v09` once to add the buttons and protection.
- **Layout decision:** the Line label spans B:E (overflow), so IO·Code is in F, Budget in G and months start at H. Collapsing a quarter never hides E6:G6 or the KPI cards. `TL_Anchor` is still B21.
- **Checks done here:**
  - `tests/grid_model.py`: a Python port of the grid, with 34 checks recalculated in LibreOffice, all passing.
  - A static review of the VBA.
- **Next session:**
  1. Ask for the UAT results (`tests/UAT_test_script.md`), plus the compile result from Debug → Compile.
  2. Fix what fails. Re-run `tests/grid_model.py` if the grid logic changes, and keep it in step with `modTimeline.bas`.
  3. For test 13, run `python3 -I build/compare_pq.py <v0.8> <user's saved v0.9>`.
  4. After UAT: promote the selected-test queries one at a time (CLAUDE.md business rules), then delete `Legacy_Timeline`.
- **Rename reverted:** the first build renamed the manpower table to `tblManpowerSelectedTest` (sheet `PQ_Manpower_SelectedTest`) by editing XML, which caused the queryTable9 repair (§0). The rebuild keeps the v0.8 names. If the user still wants the new names, they rename in the Excel UI and the `MANPOWER_TABLE` constant in `modTimeline.bas` is changed to match.
- **Gotchas:**
  - Give the user `.bas` files from the release zip or as chat attachments. GitHub stores them with LF and its "Raw" download serves LF, which File → Import doesn't handle reliably.
  - The `.bas` files must stay ASCII with CRLF line endings. Unicode labels use `ChrW()`.
  - ThisWorkbook code has to be pasted, not imported.
  - Never open or save the workbook in LibreOffice; it strips the queries.
