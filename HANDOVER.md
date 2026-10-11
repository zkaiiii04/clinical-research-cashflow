# Handover: Clinical Research Cash Flow Dashboard Timeline

**From session:** 11 October 2026 (design discovery → spec approval)
**Where:** `main` branch of `zkaiiii04/clinical-research-cashflow` (renamed from `wardrobe-atlas` on 11 Oct 2026; GitHub redirects the old name). Design work was merged in PR #1; the unrelated Wardrobe Atlas page was removed in PR #2.
**Status:** spec **v1.0 approved** (v1.0.2). v0.9_working rebuilt 11 Oct 2026 after the queryTable9 repair fix; the user then loaded it in desktop Excel with no repair prompt. **Open defect: the timeline uses the SOA budget LOA dates instead of the project dates, so months after the LOA end drop out of the grid (found on tDCS). Fix it first (§0).** UAT continues after that (§9).

---

## 0. START HERE: timeline end date defect (logged 11 Oct 2026, not fixed)

Full write-up: **`docs/incidents/2026-10-11-v0.9-project-end-date.md`**.

**What the user saw (tDCS, Main IO TF24HIP134):** the timeline stops at Dec 2026. The RA line is missing Jan–Sep 2027, and 26,975 sits on "Unmatched: manpower months outside project dates". The MP Projected Project Total (54,725) and `qManpower_SelectedTest` are correct.

**Cause:** the grid months and C7/C8 take Project Start/End from `tblSelectedProjectMapping`. Those are the min/max of the **SOA budget LOA dates** (`qProjectIOCatalogue`), which give 1 Jan 2025 – 31 Dec 2026. The real project dates are in **`qProjects_SelectedTest`** (IREx): 1 Jan 2025 – 30 Sep 2027. This was a spec error (§10.1 step 5, §12).

**Do this first:**
1. Ask the user:
   - to upload the current v0.9_working, or v0.8 for a rebuild;
   - whether the macros are already installed in their copy (this decides whether they re-import `modTimeline.bas` or take a whole rebuilt package);
   - whether they've checked the **OT vs RA** question below in IREx.
2. Change `vba/modTimeline.bas`:
   - read dates from `qProjects_SelectedTest` (the IO Code Type "Main" row for F6), falling back to the mapping table only when those are blank;
   - warn at Load when the SOA LOA dates and the project dates differ.
3. Change `build/build_v09_working.py` so C7/C8 point to `qProjects_SelectedTest`.
4. Add a regression case to `tests/grid_model.py`: data like tDCS should give 33 months (Jan 25 – Sep 27), no Unmatched manpower row, the RA line at 43,575 across its months, and MP Projected = 54,725.
5. Update spec §10.1 and §12, `CLAUDE.md`, the technical specification and the build log.
6. Rebuild from v0.8, verify, and repackage with a new SHA-256.
7. Re-test tDCS with the user, then carry on with UAT (§9).

**Open question for the user (source data):** the manpower table shows **RA = 43,575** (`DB_RA_Distrib`) and **AHS (OT/PT) = 11,150** (`DB_AHS_Nurse_Distrib`). The user expected the reverse (OT ≈ 43,575, RA ≈ 11,150; budget OT 42,000 / RA 10,800). The staff group comes from which IREx sheet a row is on. Either the tDCS rows are on the wrong sheets in IREx, or the labels were crossed. This isn't a workbook bug unless the source proves otherwise.

**Status of the queryTable9 incident (§0b):** the uploaded v0.9 (v0.8 manpower names) opened and loaded in desktop Excel without a repair prompt. Confirm with the user and close it.

## 0b. v0.9 repair error (queryTable9), fixed 11 Oct 2026; the user's later upload suggests it's resolved

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

1. `CLAUDE.md`: project rules, architecture, conventions, data quirks.
2. `specs/2026-10-11-cashflow-dynamic-timeline-spec.md`: **the build spec (authoritative).** §10 covers the macros, §12 the KPI repointing and migration, §13 the 15 acceptance tests.
3. `brainstorms/2026-10-11-cashflow-dynamic-timeline.md`: why each decision was made (Q1–Q33). The summary at the top is reconciled; the later Q entries supersede earlier ones where they conflict, and those are marked.
4. `prototypes/cashflow-timeline-prototype.html`: what the user approved visually (https://claude.ai/artifact/AYqLAxM4obRV3eJAXDBLDv).

Supporting documents, updated this session: `docs/Clinical_Research_Cash_Flow_PRD_v1.7.docx`, `docs/Clinical_Research_Cash_Flow_TECHNICAL_SPECIFICATION.md`, `docs/Clinical_Research_Cash_Flow_Build.md`.

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
8. Give the user a **desktop-Excel test script** for the 15 acceptance tests in spec §13. The key figures for 3D Gait: Budget **$250,000.00**; Actual **$68,563.67** (MP **$24,367.83**, OM **$44,195.84**); MP Projected **$62,250.00**; 17 months Nov 2025–Mar 2027; 6 quarters; FY25/FY26 totals.
9. Commit the `.bas` files, the build script and the test script to the branch. Don't commit the workbook unless the user asks.

## 5. Gotchas found this session

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

> Continue the Clinical Research Cash Flow dashboard work on branch `claude/cash-flow-projection-dashboard-x4eqkf` of zkaiiii04/clinical-research-cashflow. Read `HANDOVER.md` (start at §0), `CLAUDE.md` and `docs/incidents/2026-10-11-v0.9-project-end-date.md`. I've uploaded my current v0.9_working (and v0.8). Fix the timeline end-date defect, rebuild, and tell me what to re-import.

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
  - The `.bas` files must stay ASCII with CRLF line endings. Unicode labels use `ChrW()`.
  - ThisWorkbook code has to be pasted, not imported.
  - Never open or save the workbook in LibreOffice; it strips the queries.
