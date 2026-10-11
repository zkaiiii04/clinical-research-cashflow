## Clinical Research Cash Flow Projection - Build Status

### Status as at 11 October 2026 (updated after dashboard redesign session)

#### Stable baseline

- Workbook: `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.8.xlsm`
- Working copy: `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.9_working.xlsm` (**not created yet**: the user approved the design but said "don't create yet")
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

#### Pre-build housekeeping still required

- Move or remove the stray `=SUM(...)` formula under the `qManpower_SelectedTest` table (sheet row 15).
- Refresh all selected-test tables before reconciling. The saved v0.8 snapshots mix projects: budget, actuals and projects show Esko Combo, while mapping and manpower show 3D Gait.

### Current status

Automatic selected-project testing has passed for budget, actuals, projects and manpower. The dashboard timeline design is approved at spec v1.0, but **the workbook build has not started**. Promotion of the selected-test queries over the original outputs is deferred until after the dashboard build and UAT.

### Next phase

1. When the user gives the go-ahead, create `v0.9_working` from v0.8 (v0.8 untouched).
2. Do the pre-build housekeeping above.
3. Add the `ProjectionStore` (very hidden), `Dashboard_Settings` (`tblManpowerChargeTo`) and `Legacy_Timeline` sheets, and move the old timeline and input tables to `Legacy_Timeline`.
4. Write the VBA modules `modTimeline` and `modStore`: Load Project, Save Projections, Re-assign and Delete Unmatched, plus a `Workbook_Open` step that enables outlining under protection.
5. Repoint the KPI cards and C7/C8/C9 per spec §12.
6. Run the 15 acceptance tests (spec §13) in desktop Excel, including cross-project switching and a check that no Power Query definition changed.
7. Then promote the selected-test queries one at a time, with KPI, timeline and cross-project regression testing; complete wider UAT and multi-user SharePoint testing.
