## Clinical Research Cash Flow Projection - Technical Specification

*Updated 11 October 2026 after the dashboard redesign session. Sections 6–8 are new or rewritten. The detailed build specification is `specs/2026-10-11-cashflow-dynamic-timeline-spec.md` (v1.0, approved).*

### 1. Current architecture

The workbook connects to the TTSHIRExAdmin SharePoint site using Organizational Account authentication. Source workbooks remain read-only.

**Stable baseline:** `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.8.xlsm`
**Next working version (not yet created):** `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.9_working.xlsm`

Core loaded outputs remain active during controlled testing:

- `qBudget` → `tblBudget`
- `qActuals` → `tblActuals`
- `qProjects` → `tblProjects`
- `qManpower` → `tblManpower`

Accepted selected-test objects operate in parallel and do not yet replace these outputs. The redesigned dashboard timeline (section 6) reads the selected-test outputs directly.

### 2. Core business rules

- Preserve signed budget and actual values, including negative adjustments.
- Monthly manpower values are FTE, not currency.
- Projected manpower cost equals Monthly FTE multiplied by Monthly Rate.
- Monthly Rate comes from `qManpowerRates` / `tbl_ManpowerRate`.
- Apply the latest eligible rate with Effective Date less than or equal to the projection month.
- Do not hardcode rates.
- Do not derive IO relationships by adding, removing or interpreting suffixes.
- Do not modify source workbooks.
- Dashboard VBA may refresh the selected-test tables but must never create, edit or delete queries, connections or loaded tables.

### 3. Automatic project-selection architecture

#### 3.1 User controls

`Main_Dashboard` provides a Source Project selector (E6), a dependent Main IO selector (F6) and Setup Status (G6). `tblProjectSelection` (`Lists!N3:P4`) exposes Source Project, Main IO Code and Setup Status to Power Query. The older C5 Project Display Name selector is retired by the redesign.

#### 3.2 Project and IO catalogue

`qProjectIOCatalogue` loads as `tblProjectIOCatalogue` at this grain:

`Source Project Name + IO Code + Budget Code Type`

Multiple Budget Code Types for one IO are retained. Source Project Name is trimmed and cleaned; IO Code and Budget Code Type are trimmed, cleaned and uppercased. Budget Code Types present at 11 Oct 2026: MP, OM, EQ, CO, AL, OT, B, C, D, and blank.

#### 3.3 Selected-project mapping

`qSelectedProjectMapping` loads as `tblSelectedProjectMapping` with Project Display Name, Source Project Name, IO Code, Budget Code Type, Main IO, Project Start and Project End. It filters the catalogue to the selected project and identifies the confirmed Main IO.

The current test project name is `HoW2.0 - 3D Gait`. The prior user-defined label `HoW2.0 - 3D Gait typedc` is obsolete.

### 4. Controlled selected-test queries

#### 4.1 `qBudget_SelectedTest` → `tblBudgetSelectedTest`

- Inner join at project, IO and Budget Code Type grain.
- Signed Allocation renamed Revised Budget ($).
- Output: Project Group, IO Code, Expense Category, Budget Description, Revised Budget ($), Source Row ID, Main IO.
- Accepted result: **$250,000.00**.
- Note: Budget Description repeats as signed adjustments (e.g. RA 64,200 + 8,400; Research Fellow 39,000 − 39,000) and is free text with typos.

#### 4.2 `qActuals_SelectedTest` → `tblActualsSelectedTest`

- Inner join at IO and Budget Code Type grain.
- Signed transactions preserved.
- Transaction Description = SOA `Name` | `Cost element descr` (e.g. "Payroll for AUG 2026 | Sal -  Ald Health"). Nothing links an actual to a budget line or a manpower line.
- Accepted results: **$68,563.67** total, **$24,367.83 MP**, **$44,195.84 OM**.

#### 4.3 `qProjects_SelectedTest`

- Reads `qSelectedProjectMapping`.
- Removes duplicate Budget Code Type rows at Project + IO grain before metadata matching.
- Main IO = Yes maps to Main; other IOs map to Associated.
- Preserves Project Group, IO Code, Project Title, Alias, Project Start, Project End, Status and IO Code Type.

### 5. Manpower selected-test specification

#### 5.1 Existing function dependency

The original `fnManpowerSheet` reads the manpower IO explicitly from the selected source worksheet column and previously joined that IO to `qProjectMapping` to obtain Project Group. `qManpower` itself did not contain an IO Code Type = Manpower filter. Its final hardcoded test-project filter was obsolete.

#### 5.2 Duplicate function

`fnManpowerSheet_SelectedTest` preserves all source extraction, monthly FTE unpivoting, date conversion, source IO cleaning and audit fields. It replaces only the old project-attribution join.

Routing rules:

1. Obtain distinct nonblank `Source Project Name` values from `qSelectedProjectMapping`.
2. Continue only when exactly one selected project exists.
3. Confirm at least one selected mapping row has normalized `Budget Code Type = MP`.
4. Filter source manpower rows where normalized source `Project Title` or normalized source `Alias` equals the normalized selected project name.
5. Add Project Group using the selected Source Project Name.
6. Retain the exact manpower `IO Code` from the source worksheet.
7. Do not derive the manpower IO from the Main IO.

#### 5.3 Duplicate manpower query

`qManpower_SelectedTest` calls `fnManpowerSheet_SelectedTest` for:

- `DB_RA_Distrib` / Research Assistant
- `DB_AHS_Nurse_Distrib` / Allied Health / Nurse

It preserves the existing rate-resolution logic:

- RA rows select the latest eligible `qManpowerRates` row for Category RA and Grade Family AH11.
- Allied Health / Nurse rows resolve from the source Rate Selected value.
- Projected Amount ($) = Monthly FTE × Monthly Rate.
- Rate Status = "Rate matched" when Monthly Rate is present, otherwise "Rate not found".

The obsolete hardcoded final filter for `HoW2.0 - 3D Gait typedc` is removed from the duplicate.

#### 5.4 Accepted manpower results

For HoW2.0 - 3D Gait:

- Main IO: `TF25HIP134`
- Source manpower IO: `TF25HIP134MP`
- Total Monthly FTE: **9.0**
- Projected Amount: **$62,250.00**
- Source RA Estimate Utilisation: **$62,250.00**
- Rate Status: `Rate matched` for all relevant rows

Cross-project switching passed with HoW2.0 - Esko Combo, including Research Assistant and Allied Health / Nurse rows. No multiplication from multiple Budget Code Type rows was observed.

#### 5.5 Manpower output and audit fields

Retain: Project Group, IO Code, Staff Group, Assignment ID, Item Description, Alias, Tag Start, Tag End, Month, Monthly FTE, Rate Selected, Resolved Rate Grade, Rate Category, Rate Effective Date, Monthly Rate, Projected Amount ($), Rate Status, Source Row ID.

- `Tag Start` and `Tag End` represent LOA start and LOA end dates. Keep the existing names until downstream dependencies are reviewed.
- RA Rate Selected may be blank because RA grade resolution is automatic.
- **Assignment ID and Source Row ID are generated after the unpivot and change per month.** They are not stable line identifiers and must not be assumed globally unique.
- `Item Description` is the source Project Title, not a role or person, and no person or position column is extracted.

#### 5.6 Load and promotion controls

- `fnManpowerSheet_SelectedTest` and `qManpower_SelectedTest` remain parallel test objects.
- **`qManpower_SelectedTest` is loaded to a sheet and must stay loaded** (table `qManpower_SelectedTest` on sheet `PQManpower_SelectedTest`; v0.9 keeps these names, see incident 2026-10-11), because the dashboard macro reads it. This corrects earlier notes that described it as connection-only. A stray `=SUM(...)` formula under the table (row 15) must be moved or removed.
- Original `fnManpowerSheet`, `qManpower`, and `tblManpower` remain unchanged.
- Temporary inspection queries may be deleted after dependency verification.
- Promotion is deferred until after the dashboard build and UAT.

#### 5.7 Manpower rate table categories (`tbl_ManpowerRate`)

| Category | Example grades |
|---|---|
| AHS (OT/PT) | AH12–AH17 `_Jul25 (AHS)`, AH12–AH14 `_Apr24 (AHS)` |
| AHS | TA10_Jul25 (AHS) |
| RA | AH11–AH13 `_Jul25 (RA)`, AH11–AH12 `_Apr24 (RA)` |
| Nurse | SN11_Jul25 (Nurse), SSN_Jul26 (Nurse) |
| Admin | JIG11–JIG13_Jul25 (Admin) |

Grade text must have a space before "(" for effective-date parsing. `SSN_Jul26(Nurse)` was corrected in the source by the user on 11 Oct 2026. `qManpowerRates` upper-cases Category.

### 6. Dashboard timeline (approved design v1.0, build pending)

This replaces the v0.8 fixed-capacity, formula-driven timeline (G21:BO56). Full detail is in the spec. The key technical points:

#### 6.1 Inputs and driver

- Driver: E6 Source Project + F6 Main IO. Project key = Source Project + Main IO.
- Reads `qProjects_SelectedTest` for the project dates (IO Code Type "Main" row; IREx project information), `tblSelectedProjectMapping` for the project names and Main IO (its Project Start/End are the SOA budget LOA dates and are only a fallback), `tblBudgetSelectedTest`, `qManpower_SelectedTest` and `tblActualsSelectedTest`. Load warns when the SOA LOA dates and the project dates differ.

#### 6.2 Macros (new modules `modTimeline`, `modStore`; `Module1` untouched)

- `LoadProject`:
  - Prompts to save unsaved edits, then checks G6 = Ready.
  - Refreshes only the five selected-test tables with `QueryTable.Refresh BackgroundQuery:=False`, restoring each table's background setting afterwards.
  - Runs a consistency check: every table must reflect E6/F6, otherwise it stops and leaves the grid unchanged.
  - Builds the grid at `TL_Anchor` (B21), loads the store, places Unmatched rows, writes the "loaded by" notice, captures the session state, and shows warnings (Rate not found, Admin rows, unmatched).
- `SaveProjections`: checks the selection still matches; shows the conflict warning if the store's Saved At is later than the load time; replaces this project's store rows.
- `ReassignUnmatched` / `DeleteUnmatched`.
- `Workbook_Open`: only sets `EnableOutlining` and protects with `UserInterfaceOnly:=True`.
- No `Worksheet_Change` / `AfterRefresh` / `SheetCalculate` handlers. Unsaved edits are detected by comparing input cells with a snapshot taken at load.

#### 6.3 Grid

- Columns:
  - Line | IO·Code | Budget.
  - Months from Project Start to Project End, with a quarter total after each quarter. Each quarter's months are grouped in the outline (summary on the right). Quarters at the start and end are partial.
  - Then all FY (or CY) totals, then Project Total.
- Quarter colours, total / month tint: Q1 `#BDD7EE`/`#DDEBF7`, Q2 `#C6E0B4`/`#E2EFDA`, Q3 `#FFE699`/`#FFF2CC`, Q4 `#F8CBAD`/`#FCE4D6`.
- Labels: `FY25/Q3` (FY named by its starting year, Apr–Mar) or `2026 | Q1`.
- IO·Code: the IO with the budget code joined on, no space (`TF25HIP136OM`). Manpower lines show the source manpower IO. Consecutive identical values are merged vertically.
- Blocks MP, OM, EQ, CO, OTHER (OTHER = every other code). Non-MP: budget lines merged by IO + category + normalised description (trimmed, spaces collapsed, non-printing characters removed, upper case), net budget, $0 lines hidden, monthly input cells. MP: budget lines (budget only) and manpower lines grouped by Rate Category + IO, labelled `<Category> (Projected)`. "Rate not found" rows are shown and flagged.
- Each block has Projected, Actual, Variance ($) = Actual − Projected, and Variance (%) = blank when Projected = 0. Variance appears only for months up to the Latest SOA Month, and period variances use projections up to that month. Actuals outside the project dates count in the Project Total with a note.
- Cells hold values for data and grid-local formulas for totals. There are no references to Power Query tables in the grid.
- Format `#,##0;[Red](#,##0)` and % to 1 dp, with no colour coding by meaning. Input cells are shaded `#FFFFE0` and validated as decimals.

#### 6.4 Settings and store

- `Dashboard_Settings!tblManpowerChargeTo`: Source Project, Main IO, Rate Category, IO Code → Charge To (MP / OM / EQ / CO / OTHER). The default is MP.
- `ProjectionStore!tblProjectionStore` (very hidden): Source Project, Main IO, IO Code, Expense Category, Budget Code, Line Key, Budget Description, Month, Amount, Saved By, Saved At. A save replaces the project's rows. Blank is not stored; 0 is stored.
- Unmatched saved rows (line gone, line now $0, month outside dates) are shown on an Unmatched row, counted in totals, and can be re-assigned or deleted.

#### 6.5 KPI cards and retirement

- C7/C8 come from `qProjects_SelectedTest` (IO Code Type "Main"), falling back to `tblSelectedProjectMapping` (Main IO row) only when blank. C9 = MAX of the actuals' Posting Month. C14, F14 and F17 come from the grid's `TL_TotalBudget`, `TL_TotalActual` and `TL_TotalProjected`. C17 and C19 are ratios.
- C5, `tblManualProjection` and `tblPatientReimbursement` are retired. The old area moves to a hidden `Legacy_Timeline` sheet during UAT. Patient reimbursement is typed on its OM budget line.

#### 6.6 Protection

- `Main_Dashboard` is protected with no password. Input cells, E6, F6 and C6 are unlocked. The `ProjectionStore` sheet is very hidden.

### 7. Operating model

- One shared `.xlsm` on SharePoint, opened in desktop Excel; macros don't run in Excel for the web.
- Several people often have the file open at once. The dashboard is driven by one person at a time because the selector and selected-test queries hold one project for the whole file. Saves are per project, stamped, and conflict-checked.

### 8. Change control and next phase

- Save a new workbook version before material changes. Do not overwrite the stable v0.8 baseline.
- All four selected-test areas are passed but not promoted.
- Build the approved dashboard in v0.9_working, then run the spec's 15 acceptance tests.
- Then promote one query at a time, preferring the existing table names, and repeat the KPI, timeline and cross-project reconciliation after each promotion.
