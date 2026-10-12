# Clinical Research Cash Flow — Dynamic Projection Timeline

**Specification v1.0 — APPROVED by the user on 11 October 2026 (current revision v1.0.7). Built as v0.9_working; in UAT.**

**Change log:** v1.0.7 (12 Oct 2026, user request, Stage 2): a **Remarks** column after Project Total on every line row. Remarks are saved per project by Save Projections and follow a line to its Unmatched row (§5.5, §8, §9, §13 test 18). v1.0.6 (12 Oct 2026): the user withdrew v1.0.5 (a), because they didn't want the empty F:G spacer. The Line column is back to B:E (IO·Code F, Budget G); (b) and (c) stay. v1.0.5 (12 Oct 2026, user requests): (a) narrower Line column, withdrawn in v1.0.6: the label is merged across B:C on each line row and wraps (≈40 characters), IO·Code moves to D and Budget to E, F:G are an empty spacer, and the months still start at H, so column widths and the rows above the grid are unchanged; (b) every build groups rows 2–19 (`TOP_FIRST_ROW`/`TOP_LAST_ROW` in `modTimeline`), collapsible with the +/− beside row 1, left expanded; the buttons move and size with their cells so they hide with the rows; (c) the macro doesn't set or clear frozen panes, so panes the user sets are kept. v1.0.4 (12 Oct 2026): the user moved the project selector from E6:G6 to E5:G5 in their working copy. The macros read the selector, notice, Calendar Basis and Last Refreshed cells from constants in `modTimeline` (`SEL_PROJECT`, `SEL_MAINIO`, `SEL_STATUS`, `NOTICE_CELL`, `BASIS_CELL`, `REFRESHED_CELL`). References to E6/F6/G6 in this spec mean those constants. v1.0.3 (11 Oct 2026, defect fix): the timeline's months and C7/C8 take Project Start/End from **`qProjects_SelectedTest`** (IREx project dates, IO Code Type "Main" row), not from `tblSelectedProjectMapping`, whose dates are the min/max of the SOA budget LOA dates. The mapping dates are only a fallback when the project dates are blank, and Load warns when the two differ (`docs/incidents/2026-10-11-v0.9-project-end-date.md`). Acceptance test 16 added. v1.0.2 (11 Oct 2026, incident fix): the v1.0.1 rename is reverted. The manpower table keeps its v0.8 name `qManpower_SelectedTest` on sheet `PQManpower_SelectedTest`, because renaming it by editing the file broke its query link (`docs/incidents/2026-10-11-v0.9-querytable9-repair.md`). It can be renamed later in the Excel UI; the VBA then needs its `MANPOWER_TABLE` constant changed. v1.0.1 (11 Oct 2026, build): the manpower table was renamed `tblManpowerSelectedTest` (user request). v1.0 approval; variance uses the standard number format with no colour coding by meaning. v0.2 merges Code into an IO · Code column (no space, merged across consecutive lines), adds `(Projected)` to manpower row labels, and records the accepted decisions in §14. Prototype: `prototypes/cashflow-timeline-prototype.html`.
**Baseline:** `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.8.xlsm` (not modified)
**Build target:** `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.9_working.xlsm`
**Source of decisions:** `brainstorms/2026-10-11-cashflow-dynamic-timeline.md` (Q1–Q32)

---

## 1. Purpose

Replace the fixed timeline on `Main_Dashboard` (G21:BO56: 60 pre-formatted columns, every cell holding SUMIFS formulas) with a timeline built by a macro for the selected project. The new timeline:

1. Works out the months, quarters and financial-year columns from the project's start and end dates.
2. Lists budget line items (non-MP) and manpower lines (MP) as rows.
3. Lets users type monthly projections for OM, EQ, CO and OTHER budget lines.
4. Saves and loads those projections separately for each project.
5. Shows projected vs actual for each category by month, quarter, financial year and project total.
6. Never changes Power Query definitions, connections or loaded tables.

## 2. Scope

| In scope | Out of scope (this phase) |
|---|---|
| Macro-built timeline grid on `Main_Dashboard` | Changes to any Power Query query or function |
| Projection store (very hidden sheet) | Line-level actuals (actuals are by category only) |
| Charge-to settings table | Linking MP budget lines to manpower lines |
| Load Project / Save Projections / Unmatched Re-assign & Delete | Forecast (actual + remaining) rows |
| Repointing the KPI cards to the new grid | Charts, special print layouts, version history |
| Moving the old timeline and input tables to a hidden comparison sheet | Promoting the selected-test queries over the original tables |

## 3. Users and operating model

- One shared `.xlsm` on SharePoint, opened in **desktop Excel** (macros don't run in Excel for the web).
- Several people often have the file open at the same time. **Only one person drives the dashboard at a time**, because the E6/F6 selector and the selected-test queries hold a single project for the whole file.
- When someone loads a project, a notice appears beside the selector: **"Dashboard loaded by <user> at <time>"**. This is a notice only and doesn't lock anything.
- Anyone with edit access can save projections. Every saved row records who saved it and when.

## 4. Data sources (read-only to the macro)

| Purpose | Object | Fields used |
|---|---|---|
| Selected project | `tblSelectedProjectMapping` | Source Project Name, Project Display Name, IO Code, Budget Code Type, Main IO; Project Start / Project End only as a fallback (they are the SOA budget LOA dates) |
| Project dates | `qProjects_SelectedTest` | IO Code, IO Code Type ("Main"), Project Start, Project End (IREx project information, authoritative) |
| Budget lines | `tblBudgetSelectedTest` | Project Group, IO Code, Expense Category, Budget Description, Revised Budget ($) |
| Manpower projection | `qManpower_SelectedTest` (table on `PQManpower_SelectedTest`, loaded by the query of the same name) | Project Group, IO Code, Staff Group, Month, Monthly FTE, Rate Category, Projected Amount ($), Rate Status |
| Actuals | `tblActualsSelectedTest` | Project Group, Expense Category, Posting Month, Actual Amount ($) |
| Selector | `Main_Dashboard!E6`, `F6`, `G6`; `tblProjectSelection` | Source Project, Main IO Code, Setup Status |
| Calendar basis | `Main_Dashboard!C6` | "Financial Year" / "Calendar Year" |

**Dependency note:** the manpower query's table (`qManpower_SelectedTest`) is **loaded to a sheet**. The documents describe it as connection-only, but the macro needs it loaded, so it must stay loaded. The stray `=SUM(...)` formula under that table (row 15) should be moved above the table or deleted, so that a longer manpower result doesn't run into it.

## 5. Grid layout

### 5.1 Placement

- The grid starts at the old timeline anchor (`Main_Dashboard!B21`), held in the defined name `TL_Anchor`.
- Nothing else may sit below the grid. The old input tables (rows 57–71) are moved away as part of the migration (§12), and the "Back to Top" link is written beneath the grid each time it's built.
- The grid is cleared completely (values, formats, outline, notes) before every build.

### 5.2 Columns

```
| Line | IO · Code | Budget | <months + quarter totals …> | <FY/CY totals …> | Project Total | Remarks |
```


- **IO · Code:** one column holding the IO with its budget code joined on, **no space**, following the manpower IO convention: e.g. `TF25HIP136MP`, `TF25HIP136OM`, `TF25HIP136EQ`, `TR25NMR043AL`. Manpower lines show their manpower IO as it comes from the source (e.g. `TF25HIP136MP`). The cell is **merged vertically** across consecutive lines in a block that share the same value, so the IO isn't repeated on every line. In the MP block, budget lines (`TF25HIP136` + `MP`) and manpower lines (`TF25HIP136MP`) normally share one merged cell. The category name is shown on the block header row rather than in its own column.
- **Months:** every month from the first month of Project Start to the month of Project End, using the project dates in `qProjects_SelectedTest` (the IO Code Type "Main" row for the Main IO, otherwise the first "Main" row). Only if those are blank does the build fall back to the Main IO row of `tblSelectedProjectMapping` (SOA budget LOA dates).
- **Quarter total:** placed straight after the last active month of each quarter. Quarters at the start and end of the project hold only their active months.
- **Quarter labels:**
  - Financial Year (Apr–Mar): `FY25/Q3`, the existing convention, where the FY is named by its starting year.
  - Calendar Year: `2026 | Q1`.
- **Two header rows:** row 1 shows the quarter label across that quarter's months and its total. Row 2 shows `Nov 25`, `Dec 25`, …, `Q Total`.
- **Outline:** each quarter's month columns are grouped (summary column on the right), so the user can collapse to quarter totals and expand back to months.
- **Colours:** each quarter number always has the same colour, whatever the year. The quarter-total header and cells use the full colour, and its months use a light tint. Starting palette (adjustable):

  | Quarter | Total | Month tint |
  |---|---|---|
  | Q1 | `#BDD7EE` | `#DDEBF7` |
  | Q2 | `#C6E0B4` | `#E2EFDA` |
  | Q3 | `#FFE699` | `#FFF2CC` |
  | Q4 | `#F8CBAD` | `#FCE4D6` |

- **FY totals:** all FY total columns sit together at the far right, after the last quarter (`FY25 Total`, `FY26 Total`, …). On the Calendar Year basis they become `2025 Total`, `2026 Total`, … and use a neutral grey.
- **Project Total:** the last figures column, followed by **Remarks** (§5.5).
- There is **no Budget-remaining column**. The macro **doesn't set or clear frozen panes**; panes the user sets are kept (v1.0.5).
- **Top rows (v1.0.5):** each build groups rows 2–19 (selector, KPI cards, buttons) so they collapse with the +/− beside row 1. They are left expanded after a build. The clear before a build removes the whole outline, which is why the macro puts the group back.

Example: HoW2.0 - 3D Gait, Nov 2025 – Mar 2027, Financial Year basis. That's 17 months, 6 quarter totals, 2 FY totals and the Project Total.

```
Nov25 Dec25 [FY25/Q3] Jan26 Feb26 Mar26 [FY25/Q4] Apr26 May26 Jun26 [FY26/Q1] … Jan27 Feb27 Mar27 [FY26/Q4] | FY25 Total | FY26 Total | Project Total
```

### 5.3 Rows

The block order is **MP, OM, EQ, CO, OTHER**, followed by **PROJECT TOTAL**. A block appears when it has budget lines, manpower lines charged to it, unmatched saved projections, or actuals.

**MP block**
1. **MP budget lines.** One row per IO + normalised Budget Description, showing the net Budget only. The month cells are blank and greyed out.
2. **Manpower lines** (read-only).
   - One row per **Rate Category + IO Code**, where the charge-to setting is MP (the default).
   - Label: the rate table's Category text as written followed by **`(Projected)`**, e.g. `RA (Projected)`, `AHS (OT/PT) (Projected)`, `AHS (Projected)`, `Nurse (Projected)`, so the locked projection rows stand apart from the MP budget lines. The manpower IO is shown in the IO · Code column, not in the label.
   - Month value = the sum of Projected Amount ($) for that month.
   - Rows with **Rate Status = "Rate not found"** go on a `<Staff Group> (Projected) ⚠ Rate not found` row, which shows a blank or $0 amount.
   - An **Admin** category isn't expected. If one appears, the row is shown anyway and Load Project warns about it.
3. **Unmatched row,** shown only when needed. It holds manpower months that fall outside the project dates (§9).
4. **Subtotal rows:** Projected, Actual, Variance ($), Variance (%).

**OM, EQ, CO, OTHER blocks**
1. **Budget lines.** One row per IO + Category + normalised Budget Description.
   - Net signed budget, with lines that net to $0 hidden.
   - The month cells are **inputs**.
   - Patient reimbursement is simply its OM budget line.
2. **Manpower lines charged to this block** (read-only, from the charge-to settings in §7). Each looks like an MP manpower line, e.g. `AHS (Projected) ⚙ manpower → OM`.
3. **Unmatched row,** shown only when needed (§9).
4. **Subtotal rows:** Projected, Actual, Variance ($), Variance (%).

OTHER holds every Budget Code Type that isn't MP, OM, EQ or CO (currently AL, OT, B, C, D). The **IO · Code** column shows the code joined to the IO (e.g. `TR25NMR043AL`). A blank code is shown as `(no code)`.

**PROJECT TOTAL:** Budget, Projected, Actual, Variance ($) and Variance (%) across all blocks.

### 5.4 Normalising descriptions

The line key is the Budget Description with leading and trailing spaces removed, repeated internal spaces collapsed, non-printing characters removed, and converted to upper case. Budget records that share Project + IO + Category + line key are added together. The row displays the description text of the first record.

### 5.5 Remarks column (v1.0.7)

- One column straight after Project Total, headed **Remarks**, about 40 characters wide, wrapping.
- **Every line row** has an unlocked (pale yellow) remark cell: typed budget lines, MP budget lines, manpower lines, Rate-not-found lines and Unmatched rows. Block headers and subtotal rows don't have one.
- Remarks are text (one starting with `=` stays text). The column uses the standard font size, so rows keep their usual height and grow when a long remark wraps.
- **Saved per project** by Save Projections and loaded back by Load Project (§8). An edited remark counts as an unsaved change (the switch-project prompt, §10.1).
- **Line key** of a remark: budget lines use the line identity (Block | Code | IO | Line Key, §5.4). Manpower lines use `MPW|<Rate Category>|<IO>`, so changing the charge-to category keeps the remark. Rate-not-found lines use `RNF|<Staff Group>|<IO>`, the MP Unmatched manpower row uses `OOR|<Block>`, and Unmatched rows built from saved amounts use `UNM|<identity>`.
- **Unmatched (§9):** if a remark's line is gone, the remark shows on that line's Unmatched row. If the line has no saved amounts, it gets an Unmatched row of its own, labelled "(… ; remark only)". Re-assign moves the remark to the target line, after any remark that line already has, and Delete removes it. Re-assign needs a typed budget line in the same block, so a remark-only row in the **MP** block (manpower or MP budget line gone) can only be deleted; copy the text first if it's still needed. If a remark had to move to a different row on Load, the grid counts as changed until it is saved.

## 6. Cell contents

| Cell | Content |
|---|---|
| Budget (line) | Value: net signed Revised Budget ($) |
| Manpower month (line) | Value, written by the macro |
| Input month (OM/EQ/CO/OTHER line) | Value loaded from the store, or blank. Unlocked and shaded `#FFFFE0` |
| Unmatched month | Value, read-only |
| Quarter / FY / Project Total (any line) | Local formula, e.g. `=SUM(H30:J30)` (the months in that period) |
| Subtotal Projected (month) | `=SUM(<line rows in block>)` |
| Subtotal Actual (month) | Value: sum of Actual Amount ($) for the category and Posting Month |
| Subtotal Actual (quarter/FY/project) | `=SUM(<that row's months>)` |
| Variance ($) | `=Actual − Projected` (positive = overspend) |
| Variance (%) | `=IF(Projected=0,"",Variance/Projected)` |
| Budget (subtotal / project) | `=SUM(<line budgets>)` |

Formulas only ever point at cells within the grid. **No SUMIFS or other references to Power Query tables appear in the grid.**

**Formats:** whole dollars with thousands separators, negatives shown in red in brackets (`#,##0;[Red](#,##0)`), percentages to 1 decimal place. Variance rows use the same format; there is no extra colour coding by meaning (overspend vs underspend).

**Validation on input cells:** decimal numbers only, negatives allowed. If a blank input is saved, nothing is stored for it. A typed 0 is stored as 0.

## 7. Settings sheet: `Dashboard_Settings`

`tblManpowerChargeTo`, maintained by the admin:

| Source Project | Main IO | Rate Category | IO Code | Charge To |
|---|---|---|---|---|
| HoW2.0 - Esko Combo | TF25HIP136 | AHS | TF25HIP136MP | OM |

- Any manpower line without an entry here is charged to **MP**.
- Charge To must be one of MP, OM, EQ, CO or OTHER (enforced by data validation).
- The setting applies to the whole line (Rate Category + IO), not to individual people.

## 8. Projection store: `ProjectionStore` (very hidden)

`tblProjectionStore`:

| Column | Notes |
|---|---|
| Source Project | E6 at the time of saving |
| Main IO | F6 at the time of saving |
| IO Code | |
| Expense Category | MP / OM / EQ / CO / OTHER |
| Budget Code | The raw code (matters for OTHER: AL, OT, …) |
| Line Key | Normalised description (§5.4) |
| Budget Description | The description text shown when the row was saved |
| Month | First day of the month |
| Amount | Number, 0 allowed |
| Saved By | `Application.UserName` |
| Saved At | `Now()` |

Each project is identified by the key **Source Project + Main IO**.

**Remarks (v1.0.7)** are kept beside the table, in a plain range at `ProjectionStore!W:AD` with a header row: Source Project | Main IO | Line Ref (the remark key, §5.5) | Block | Description | Remark | Saved By | Saved At. `tblProjectionStore` itself is unchanged, and existing saved projections need no migration. The grid layout area that Load writes is now `M:T` (column T holds each row's remark key).

Session state is kept in hidden defined names: the loaded project key, the time it was loaded, and the input values captured at load. The captured values are used to detect unsaved edits **without any sheet-change events**.

## 9. Unmatched saved projections

When a project is loaded, a saved store row is **unmatched** if:
- its line no longer exists (description renamed or removed in SOA, or the IO is gone), or
- its line now nets to $0 and is therefore hidden, or
- its month falls outside the project's current dates.

Unmatched rows are handled like this:
- They appear on the category's **Unmatched** row, are read-only, and **are counted** in the subtotal and Project Total.
- Months outside the project dates have no column. Their amounts go into the Unmatched row's **Project Total**, and a cell note lists them by month.
- **Re-assign:** select the Unmatched row and choose a current budget line in the same category. The store rows are moved to that line, adding to any amount already there for the same month.
- **Delete:** select the Unmatched row and confirm. Those store rows are removed.
- Manpower months outside the project dates follow the same display rule on the MP Unmatched row. They can't be re-assigned or deleted, because they come from the source.

## 10. Macros

All macros go in new standard modules (`modTimeline`, `modStore`). The existing `Module1` is not changed.

### 10.1 `LoadProject` (button beside the selector)

1. If there are unsaved edits for the currently loaded project, ask **"Save changes to <project> first? Yes / No / Cancel"**.
2. Check that `G6` = "Ready". If not, show the status message and stop.
3. Refresh **only** these loaded tables, one at a time and waiting for each to finish:
   `tblSelectedProjectMapping` → `tblBudgetSelectedTest` → `tblActualsSelectedTest` → `qProjects_SelectedTest` → `qManpower_SelectedTest`.
   - Each table's original background-refresh setting is saved first and put back afterwards.
   - `qProjectIOCatalogue` is **not** refreshed here (Refresh All covers it).
4. **Consistency check.** Every refreshed table must contain only the selected project:
   - the mapping table holds E6 with F6 as its Main IO;
   - budget and actuals Project Group = the mapping's Project Display Name;
   - manpower Project Group = Source Project Name.
   If any table doesn't match, stop with a message naming it and leave the grid unchanged.
5. Take the project dates from `qProjects_SelectedTest` (fallback: `tblSelectedProjectMapping`, see §5.2). If the mapping's SOA budget LOA dates differ from the project dates, add a Load warning naming both date ranges. Build the grid (§5–6) and apply the outline, colours, protection and validation.
6. Load the store rows for the project key into the input cells, and place any unmatched rows (§9).
7. Write the "Dashboard loaded by …" notice and capture the session state (§8).
8. Show warnings if any apply: manpower rows with "Rate not found", Admin-category rows, unmatched saved projections.

### 10.2 `SaveProjections` (button)

1. Check that a project is loaded and that E6/F6 still match the loaded key. If they don't, stop: "Selection changed — reload first".
2. **Conflict check:** if the latest Saved At for this key in the store is later than this session's load time, ask **"<Saved By> saved this project at <time> after you loaded it. Overwrite?"**
3. Delete the store rows for this key, then write one row for every non-blank input cell, plus any unmatched rows that are still present.
4. Replace this key's remarks with the non-blank Remarks cells (§5.5). Remarks for other projects are kept.
5. Update the session state.

The conflict check (step 2) also looks at the remarks' Saved At.

### 10.3 `ReassignUnmatched` / `DeleteUnmatched` (buttons)

These act on the selected Unmatched row as described in §9, then rebuild the affected block from the store.

### 10.4 `Workbook_Open`

This only sets `EnableOutlining = True` and protects `Main_Dashboard` with `UserInterfaceOnly:=True`, so the +/− buttons work on the protected sheet; Excel doesn't keep these settings between sessions. It doesn't rebuild or refresh anything.

### 10.5 Power Query safety rules

- The macros never call `Queries.Add`, `.Delete` or `.Formula =`, and never change connection strings, query definitions or table destinations.
- The only Power Query action is `ListObject.QueryTable.Refresh BackgroundQuery:=False` on the five tables listed in §10.1.
- There are no `Worksheet_Change`, `AfterRefresh` or `SheetCalculate` handlers.

## 11. Protection

- `Main_Dashboard` is protected **without a password**. Unlocked cells: the input month cells, E6, F6, C6, and the charge-to table on the settings sheet.
- The outline can be used while the sheet is protected (§10.4).
- `ProjectionStore` is set to `xlSheetVeryHidden`.

## 12. KPI cards and migration

All of this is done in `v0.9_working` only.

1. Move the old timeline (G21:BO56 and its labels), `tblManualProjection` and `tblPatientReimbursement` to a hidden sheet, `Legacy_Timeline`, for comparison during UAT. They'll be deleted after UAT sign-off.
2. Retire the C5 Project Display Name selector.
3. Repoint the KPI cards and project info:

   | Cell | New source |
   |---|---|
   | C7 Project Start / C8 Project End | `qProjects_SelectedTest` (IO Code Type = "Main"); fallback `tblSelectedProjectMapping` (Main IO = "Yes") when blank |
   | C9 Latest SOA Month | `MAX(tblActualsSelectedTest[Posting Month])` |
   | C14 Revised Approved Budget | Grid PROJECT TOTAL Budget (`TL_TotalBudget`) |
   | F14 Actual to Date | Grid PROJECT TOTAL Actual, Project Total column (`TL_TotalActual`) |
   | F17 Projected Total | Grid PROJECT TOTAL Projected, Project Total column (`TL_TotalProjected`) |
   | C17 Overall Utilisation | `=IFERROR(F14/C14,"")` |
   | C19 Projected Utilisation | `=IFERROR(F17/C14,"")` |

   The `TL_*` names are redefined by the macro on every build. Until the first Load, the KPI cards show blank.

## 13. Acceptance tests

| # | Test | Expected |
|---|---|---|
| 1 | Load HoW2.0 - 3D Gait (FY basis) | 17 month columns, Nov 2025 – Mar 2027; quarters FY25/Q3 (2 months) to FY26/Q4; FY25 Total, FY26 Total, Project Total |
| 2 | Same project, Calendar Year basis | Quarters 2025 \| Q4 (2 months) to 2027 \| Q1; 2025/2026/2027 Total columns |
| 3 | Budget reconciliation | PROJECT TOTAL Budget = **$250,000.00** |
| 4 | Actuals reconciliation | Actual Project Total = **$68,563.67**; MP **$24,367.83**; OM **$44,195.84** |
| 5 | Manpower reconciliation | MP Projected = **$62,250.00** (RA \| TF25HIP134MP) |
| 6 | Zero-net line | Esko Combo "Research Fellow" (39,000 − 39,000) hidden; "Therapy Associate Service" shown once at 30,600 |
| 7 | Save → switch → reload | Typed values come back exactly; blank stays blank; 0 stays 0 |
| 8 | Unsaved switch | Save/No/Cancel prompt appears; Cancel keeps the grid |
| 9 | Conflict | User B saves after user A loads → A gets the overwrite warning |
| 10 | Unmatched | Rename a description in the test data → saved amounts appear on the Unmatched row and in totals; Re-assign and Delete work |
| 11 | Charge-to | Map a manpower line to OM → it appears locked in OM, and MP/OM subtotals move accordingly |
| 12 | Collapse | Quarter +/− works on the protected sheet after reopening the file |
| 13 | PQ untouched | Query names and M formulas identical before and after a full Load/Save cycle (compared by script) |
| 14 | Mismatch guard | Change E6 without loading, then Save → blocked with "Selection changed — reload first" |
| 15 | Cross-project | Switch to HoW2.0 - Esko Combo, which has AHS manpower lines and EQ/CO/OM budget lines → correct blocks and figures |
| 16 | Project dates | Load tDCS (SOA budget LOA end 31 Dec 2026, IREx project end 30 Sep 2027) → timeline Jan 2025 – Sep 2027, C7/C8 = 1 Jan 2025 / 30 Sep 2027, no "manpower months outside project dates" row, MP Projected = sum of `qManpower_SelectedTest`, and a Load warning naming both date ranges |
| 17 | Top rows collapse | After Load, the − beside row 1 hides rows 2–19 (buttons hidden too, nothing floating over the grid); + shows them; still works after loading another project and under protection; frozen panes unchanged |
| 18 | Remarks | Type a remark on a typed line and on a manpower line, Save, load another project, reload → both remarks back; edit a remark then switch project → "Save changes?" prompt; a remark on a line that leaves the budget shows on its Unmatched row and moves with Re-assign / goes with Delete |

## 14. Decisions made while writing the spec

The user accepted the recommendations for items 1 and 2 ('noted on the 2 decisions'):

1. **Actuals posted outside the project dates** (e.g. a late posting after Project End). *Decided:* count them in the category's Project Total, and in the FY total if that FY column exists, with a cell note, the same way unmatched projections are handled.
2. **Monthly variance for months with no SOA yet** (after the Latest SOA Month). *Decided:* show Variance ($) and (%) only for months up to the Latest SOA Month, and leave later months blank. Quarter, FY and Project Total variances then compare actuals with projections **up to the Latest SOA Month only**, so a future plan isn't reported as underspend. *Why:* after the Latest SOA Month no actuals have been posted yet, so Actual would read 0 and Variance would equal −Projected for every future month, making the project look heavily underspent. Those periods' Actual and Variance cells are therefore left blank (not 0), and the cut-off moves forward each time a new SOA month is loaded. The Projected rows always run for the whole project. *Reconfirmed by the user on 12 Oct 2026: keep these cells blank, not 0.*
3. **Input shading colour:** light yellow `#FFFFE0` (default; change if there's a house style).

## 15. Future (not in this build)

- MP parent/child view: budget lines (RA, Therapist, …) with manpower lines nested under them, using an admin mapping table. The settings sheet is designed so this can be added later.
- Promoting the selected-test queries over `tblBudget` / `tblActuals` / `tblProjects` / `tblManpower`.
