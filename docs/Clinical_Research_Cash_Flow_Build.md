## Clinical Research Cash Flow Projection - Build Status

### Status as at 11 October 2026 (updated after dashboard redesign session)

#### Stable baseline

- Workbook: `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.8.xlsm`
- Working copy: `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.9_working.xlsm` (**built 11 Oct 2026**; VBA to be imported and tested in desktop Excel)
- Stable v0.8 remains the fallback and must not be overwritten.

### Completed and reconciled

#### Source connectivity and core queries

- SharePoint connectivity uses Organizational Account authentication.
- Budget, actuals, project metadata, manpower and manpower-rate queries are configured.
- Existing working queries and loaded tables remain unchanged during automatic-selection testing.

#### Tested project

- Current catalogue-selected project: HoW2.0 - 3D Gait
- Obsolete prior user-defined label: HoW2.0 - 3D Gait typedc
- Main IO: `TF25HIP134`
- Manpower IO from manpower source: `TF25HIP134MP`

#### Financial reconciliation

- Revised Approved Budget: **$250,000.00**
- Actual to Date: **$68,563.67**
- MP actuals: **$24,367.83**
- OM actuals: **$44,195.84**
- Total Monthly FTE: **9.0**
- RA manpower projection: **$62,250.00**
- Source RA Estimate Utilisation: **$62,250.00**
- All relevant tested manpower rows: `Rate matched`

#### Timeline (v0.8, to be replaced)

v0.8 has a fixed-capacity, formula-driven timeline at `Main_Dashboard` G21:BO56: 60 columns (15 quarter blocks) plus a Project Total column in BO. Every cell holds SUMIFS/INDEX formulas. It covers MP, OM, EQ, CO and OTHER, with monthly projected and signed actual values, dollar and percentage variance, quarter totals, calendar-year and April-to-March financial-year labels, blanking outside the project dates, the AE:BK post-project error correction, all-category summaries, a full-duration Project Total and a project-aware Latest SOA Month. **The approved redesign replaces this area** (see below).

### Automatic project-selection work completed

- `qProjectIOCatalogue` / `tblProjectIOCatalogue`: implemented at Source Project Name + IO Code + Budget Code Type grain.
- Dashboard Source Project and dependent Main IO controls: implemented.
- Setup Status and one-row `tblProjectSelection`: implemented.
- `qSelectedProjectMapping` / `tblSelectedProjectMapping`: implemented; retains all Budget Code Types and marks the confirmed Main IO.

### Controlled migration tests passed

#### `qBudget_SelectedTest`

- Automatic selected-project filtering passed.
- Signed Allocation retained as Revised Budget ($).
- Reconciled to **$250,000.00**.
- Original `qBudget` and `tblBudget` unchanged.

#### `qActuals_SelectedTest`

- Automatic selected-project filtering passed.
- Signed transactions preserved.
- Reconciled to **$68,563.67** total, **$24,367.83 MP**, and **$44,195.84 OM**.
- Original `qActuals` and `tblActuals` unchanged.

#### `qProjects_SelectedTest`

- Uses `qSelectedProjectMapping` rather than the old manual mapping.
- Duplicate Budget Code Type rows removed at Project + IO grain.
- Confirmed Main IO classified as Main; other IOs classified as Associated.
- Project Title, Alias, Project Start, Project End and Status retained.
- Original `qProjects` and `tblProjects` unchanged.

#### `fnManpowerSheet_SelectedTest` and `qManpower_SelectedTest`

- Passed.
- The old `qProjectMapping` join is removed only in the duplicate function.
- Selected project is read from `qSelectedProjectMapping[Source Project Name]`.
- `Budget Code Type = MP` is used only to confirm manpower eligibility.
- Manpower source rows are filtered by explicit source `Project Title` or `Alias`.
- Exact source IO is retained; no suffix inference is used.
- Existing RA and Allied Health / Nurse rate-resolution logic is preserved.
- HoW2.0 - 3D Gait reconciles to **9.0 FTE**, **$62,250.00**, `TF25HIP134MP`, and `Rate matched`.
- Cross-project switching passed using HoW2.0 - Esko Combo, including both staff-group routes.
- No Budget Code Type row multiplication was observed.
- Original `fnManpowerSheet`, `qManpower`, and `tblManpower` remain unchanged.
- **Correction:** `qManpower_SelectedTest` is **loaded to the `PQManpower_SelectedTest` sheet**, not connection-only as earlier notes said. The new dashboard macro depends on it, so it must stay loaded. `fnManpowerSheet_SelectedTest` is a function.

### Manpower field decisions

- `Tag Start` and `Tag End` represent LOA start and LOA end dates; names remain unchanged for compatibility.
- Keep Assignment ID, Item Description, Alias, Rate Selected and Source Row ID as audit fields.
- Blank RA Rate Selected is acceptable when Resolved Rate Grade, Monthly Rate and Rate Status are populated.
- Temporary RA, AHS/Nurse and relationship inspection queries may be deleted after verifying `qManpower_SelectedTest` has no dependency on them.
- **New finding:** Assignment ID and Source Row ID are numbered after the monthly unpivot, so they change from month to month. The dashboard therefore groups manpower by Rate Category + IO Code, not by these IDs.

### Dashboard redesign phase (11 October 2026)

#### Completed

- Design discovery session: 33 recorded decisions (`brainstorms/2026-10-11-cashflow-dynamic-timeline.md`).
- **Timeline specification v1.0 approved** (`specs/2026-10-11-cashflow-dynamic-timeline-spec.md`).
- HTML prototype built and reviewed (`prototypes/cashflow-timeline-prototype.html`). The user's feedback was incorporated: the IO and code are joined with no space and merged across lines, and manpower rows carry "(Projected)".
- PRD updated to v1.7. Technical specification updated.

#### Approved design in brief

- E6/F6 drives the timeline, which reads the selected-test tables. C5 is retired.
- **Load Project** macro: refreshes the five selected-test tables synchronously, checks they match, builds the grid sized to the project, and loads the saved projections. **Save Projections** writes to a very hidden per-project store.
- Line-item rows. OM/EQ/CO/OTHER projections are typed by users. MP manpower is read-only (FTE × rate), grouped by Rate Category + IO.
- Projected vs actual is compared at category level. Variance runs to the Latest SOA Month.
- Collapsible, colour-coded quarters; FY totals at the far right; Project Total.
- The KPI cards are repointed to the grid totals. The old timeline and input tables move to a hidden `Legacy_Timeline` sheet for UAT.

#### Source data fixes done by the user

- `tbl_ManpowerRate` grade `SSN_Jul26(Nurse)` corrected to `SSN_Jul26 (Nurse)`. The missing space broke effective-date parsing.

#### Pre-build housekeeping

- Done in v0.9_working: the stray `=SUM(...)` under `qManpower_SelectedTest` (P15) is removed.
- Still true: the saved snapshots mix projects (budget/actuals/projects show Esko Combo; mapping/manpower show 3D Gait). Load Project refreshes the five tables and checks they match before building.

### v0.9_working build (11 October 2026)

#### How it was built

- `build/build_v09_working.py` copies every part of the v0.8 package and edits only the workbook, shared-strings, content-types and two sheet parts, then adds three sheets and two tables. **openpyxl was not used to save:** a load/save with openpyxl drops the Power Query DataMashup, `connections.xml`, all query tables and the dynamic-array metadata.
- `build/verify_v09_working.py` confirms the M code (21 queries) is identical, and the 34 Power Query parts (DataMashup, connections, query tables, PQ table definitions) are byte-identical, and nothing else changed unexpectedly. The VBA project (`Module1`) is byte-identical too. LibreOffice opens the file and the `Legacy_Timeline` copy calculates exactly as v0.8 did.

#### Workbook changes (HANDOVER §4 steps 1-6)

- `Main_Dashboard`:
  - Rows 21-71 are cleared; the macro builds the grid from `TL_Anchor` = B21.
  - C5 is retired (cell and validation removed), along with the H3/J3 links to the retired input tables and the hidden L14 helper.
  - C7/C8 read `tblSelectedProjectMapping` (Main IO = Yes), and C9 = MAX of `tblActualsSelectedTest[Posting Month]`.
  - C14/F14/F17 = `TL_TotalBudget` / `TL_TotalActual` / `TL_TotalProjected`; these are defined as `""` until the first Load. C17 = `IFERROR(F14/C14,"")` and C19 = `IFERROR(F17/C14,"")`.
  - C6 has a Financial Year / Calendar Year list.
  - E7 holds the "loaded by" notice, and the four buttons sit at E8:F11.
- `Legacy_Timeline` (hidden): a full copy of the v0.8 dashboard, still driven by its own C5, holding `tblManualProjection` and `tblPatientReimbursement`. Its two obsolete notes (G3, C5) are dropped.
- `Dashboard_Settings`: `tblManpowerChargeTo` (A4:E5), with Charge To validated as MP/OM/EQ/CO/OTHER.
- `ProjectionStore` (very hidden): `tblProjectionStore` (A1:K2), plus the grid layout area at M:S that the macros rewrite on each Load.
- Manpower table: the stray `=SUM(...)` at P15 is removed and the table stays loaded. At the user's request, the table is renamed `qManpower_SelectedTest` → `tblManpowerSelectedTest` and the sheet `PQManpower_SelectedTest` → `PQ_Manpower_SelectedTest`, matching the other selected-test tables. The query keeps its name; only the table's name attributes change, and no query reads that table.
- `calcChain.xml` is removed (Excel rebuilds it), and the workbook is set to recalculate fully on open.

#### Macros (`vba/`, imported by the user; `Module1` untouched)

- `modTimeline`: `LoadProject`, the refresh of the five selected-test tables (synchronous, background setting restored), the consistency check, the grid builder, outline, colours, formats, validation, notes, protection and the `TL_*` names.
- `modStore`: `SaveProjections`, the conflict check, unsaved-edit detection (a hash of the input cells kept in a hidden name), `ReassignUnmatched`, `DeleteUnmatched`, and store read/write.
- `modSetup.Setup_Dashboard_v09`: a one-off that adds the buttons and protects the sheet.
- `ThisWorkbook.Workbook_Open`: enables outlining and protects with `UserInterfaceOnly`.

#### Build decisions within the spec

- **Line label spans B:E.** The label overflows across C:E, IO·Code is in F, Budget in G and months start at H. Otherwise collapsing the first quarter (columns E:G) would hide the E6/F6 selector and the KPI cards. `TL_Anchor` stays at B21.
- **One Unmatched row per orphaned saved line**, labelled `Unmatched: <description> (<reason>)` as in the approved prototype, with reasons "no longer in budget", "line now nets to $0" or "month outside project dates". Manpower months outside the dates go on one `Unmatched: manpower months outside project dates` row per block.
- **Re-assign:** the user clicks a cell in the target budget line (rather than picking from a list). It and Delete act on the store, then rebuild the grid without refreshing. If there are unsaved edits, they offer to save first.
- **Variance (%)** = Variance ÷ (Actual − Variance), which equals Variance ÷ projected-to-SOA, because every actual is on or before the Latest SOA Month.
- Rate categories arrive upper-cased from `qManpowerRates`. They're shown as the rate table writes them: `NURSE` → `Nurse`, `ADMIN` → `Admin`; RA, AHS and AHS (OT/PT) unchanged.
- C10 "Last Refreshed" is set to the Load time, because Load Project refreshes the tables.

#### Verification done here

- `tests/grid_model.py` ports the grid builder to Python, writes the same formulas, recalculates them in LibreOffice, and checks 34 results on the v0.8 snapshot data plus synthetic store, charge-to, Rate-not-found and out-of-range rows:
  - 17 months; FY25/Q3 … FY26/Q4; CY 2025 | Q4 … 2027 | Q1
  - budget $250,000
  - RA $62,250
  - Research Fellow hidden; Therapy Associate Service once at $30,600
  - 0 versus blank in saved projections
  - Unmatched reasons
  - out-of-range actuals in the FY and Project totals
  - variance only up to the Latest SOA Month
- The VBA was statically reviewed (no Excel here). **It has not been compiled or run yet.**

### Current status

Automatic selected-project testing has passed for budget, actuals, projects and manpower. The dashboard timeline design is approved at spec v1.0, but **the workbook build has not started**. Promotion of the selected-test queries over the original outputs is deferred until after the dashboard build and UAT.

### Next phase

1. ~~Create `v0.9_working` from v0.8; housekeeping; new sheets; legacy move; KPI repointing; VBA modules.~~ Done 11 Oct 2026, untested in Excel.
2. The user imports the VBA, runs `Setup_Dashboard_v09`, and works through `tests/UAT_test_script.md` (the 15 acceptance tests, spec §13). Fix what fails.
3. Then promote the selected-test queries one at a time, with KPI, timeline and cross-project regression testing; complete wider UAT and multi-user SharePoint testing; delete `Legacy_Timeline` after sign-off.
