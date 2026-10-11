# Handover: Clinical Research Cash Flow Dashboard Timeline

**From session:** 11 October 2026 (design discovery → spec approval)
**Where:** `main` branch of `zkaiiii04/wardrobe-atlas` (the user plans to rename this repo; GitHub redirects the old name). Design work was merged in PR #1.
**Status:** spec **v1.0 approved**. **Workbook build not started:** the user said "approve, don't create yet".

---

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

## 4. Build plan for the next session (once the user says go)

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

> Continue the Clinical Research Cash Flow dashboard work in this repo (main branch). Read `HANDOVER.md`, `CLAUDE.md` and the approved spec in `specs/`. I've uploaded the v0.8 workbook. Start the build in v0.9_working following the handover build plan.
