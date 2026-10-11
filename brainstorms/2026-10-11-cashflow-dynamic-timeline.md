# Cash Flow Dashboard — Dynamic Projection Timeline: Brainstorm / Discovery Notes
Date: 2026-10-11 · Goal: Pin down every requirement for replacing the fixed, formula-filled timeline on Main_Dashboard (G21:BO56) with a macro-built, project-specific, line-item timeline that supports manual projections with per-project save/load, without interfering with Power Query.

## Summary / key decisions
(running synthesis, updated as answers come in)

- **Platform:** single shared .xlsm on SharePoint, desktop Excel → VBA OK; projections stored in-workbook.
- **Driver/data:** E6/F6 + selected-test tables (budget, manpower, actuals, mapping). KPIs unchanged for now.
- **Non-MP line grain:** one row per IO + Category + normalised Budget Description, net budget.
- **Zero-net lines:** always hidden.
- **Ad-hoc lines:** not allowed; budget lines only.
- **Categories:** MP, OM, EQ, CO, OTHER (OTHER is separate; shown when budget lines exist).
- **MP line identity:** macro groups by Staff Group + IO + LOA start/end + Resolved Rate Grade; no PQ change; identical positions may merge.
- **MP structure:** (option 3, 'for now') MP budget lines (budget only) and manpower lines (projection) listed separately in the MP block; no mapping.
- **MP manual entry:** none — MP read-only; manual entry only on OM/EQ/CO/OTHER budget lines.
- **Actuals:** category level only (no line-level actuals).
- **Rows:** line items carry budget + projection; projected-vs-actual comparison on category subtotal rows only.
- **Columns:** Budget | months → quarter total (collapsible, fixed colour per quarter number) … | all FY/CY totals at far right | Project Total. No Budget-remaining column (pending confirm). Same layout for Calendar Year.
- **Store key:** Source Project + Main IO.
- **Concurrency:** multiple simultaneous editors are common → save must be per-project row-level merge with who/when stamps; shared-canvas problem resolved by turn-taking (one selector, one driver at a time).

- User's stated requirements (initial brief):
  1. Number of month/quarter columns computed automatically from the project; months and quarters filled in automatically (no pre-defined columns, no pre-filled formulas).
  2. Manual projection entry directly in the generated area for CO, EQ, OTHER and OM.
  3. Manual projections saved and loaded, unique per project (users can switch project).
  4. Rows show line items: non-MP from the PQ budget selected sheet (PQ_Budget_SelectedTest); MP from the manpower selected sheet (PQManpower_SelectedTest), each line item stated and differentiated.
  5. Clearly show projected vs actual per quarter and for the project total.
  6. Macro acceptable; must not interfere with existing Power Query functionality.

## Findings from reviewing v0.8 (facts, not decisions)

- Current timeline: Main_Dashboard G21:BN56 = 60 fixed columns (15 quarter blocks of Quarter Total + 3 months), BO = "Project Full Duration". Every cell pre-filled with SUMIFS/INDEX formulas; header dates in row 21 driven by C6 (Calendar Basis), C7/C8 (Project Start/End).
- Rows today: per category (MP, OM, EQ, CO, OTHER) 4 summary rows (Projected / Actual / Variance $ / Variance %) + a "Detail placeholder" row with hard 0s; PROJECT TOTAL rows 54–56.
- Dashboard formulas key on **C5 Project Display Name** (Project_Mapping list) and read **original** tables tblBudget / tblActuals / tblManpower / tblProjects. The new auto-selection uses **E6 Source Project / F6 Main IO** → tblProjectSelection. Two selectors currently co-exist.
- Manual inputs today: tblManualProjection (Project Display Name, IO Code, Expense Category, Item Description, Planned Month, Cost) at B59:G60; tblPatientReimbursement (… Patients, Rate, Projected Reimbursement, Expense Category=OM) at B66:H67. KPI "Projected Total" (F17) sums tblManpower + tblManualProjection + tblPatientReimbursement.
- PQ_Budget_SelectedTest (tblBudgetSelectedTest): Project Group, IO Code, Expense Category, Budget Description, Revised Budget ($), Source Row ID, Main IO. Descriptions repeat as signed adjustments (e.g. RA 64,200 + RA 8,400; Research Fellow 39,000 and −39,000; Therapy Associate Service 0 and 30,600). Free-text descriptions with typos ("Patient Reimbursemnt").
- PQManpower_SelectedTest (qManpower_SelectedTest — note: it IS loaded to a sheet, docs say connection-only; a SUM formula sits under the table in row 15). One row per source line × month. **Assignment ID and Source Row ID are generated after unpivot/filter, so they change per month row — there is no stable identifier for a manpower position/person.** Item Description = source Project Title (same for every line). No staff name/position column is extracted. Distinguishing fields available: Staff Group, IO Code, Tag Start/End (LOA), Rate Selected / Resolved Rate Grade, Monthly Rate.
- PQ_Actuals_SelectedTest: actuals carry Expense Category + Transaction Description (e.g. "Payroll for AUG 2026 | Sal - …", CPF, Bonus). **No field links an actual to a budget line item or a manpower line.**
- Saved snapshots are from different refreshes (Budget/Actuals/Projects selected-test show HoW2.0 - Esko Combo; Mapping/Manpower show HoW2.0 - 3D Gait).
- MP budget lines (RA, Therapist, Research Fellow) vs manpower Staff Groups (Research Assistant, Allied Health / Nurse): no mapping exists between them.
- Existing VBA: only Module1 (Install_All_SharePoint_Queries — installs .pq files). No event handlers. Workbook uses dynamic arrays (ANCHORARRAY) → Microsoft 365.

## Q&A log
### Q1 — Usage model / where it runs
- Asked: How will users open and use the workbook (desktop vs web; shared vs own copies)?
- Captured: **One shared .xlsm on SharePoint, opened in desktop Excel.** Macros are viable. Saved projections can live inside the workbook (hidden store sheet) and are shared by everyone using the file.
- Flags: Co-authoring behaviour with .xlsm (two people editing/saving at once) still to be confirmed later.

### Q2 — Concurrency
- Asked: Will multiple people have the shared file open and enter projections at the same time?
- Captured: **Yes, often at once.** Design must tolerate concurrent editors.
- Implication raised: with co-authoring, Main_Dashboard (selector E6/F6, the generated grid) and the PQ selected-test tables are ONE shared canvas — user A switching project changes what user B sees, and a PQ refresh is for one selected project for the whole file. Needs resolving (Q3).
- Flags: none yet.

### Q3 — Shared dashboard view under co-authoring
- Asked: How to stop simultaneous users overwriting each other's selection / generated view?
- Captured: **Keep one selector; take turns.** Single-project PQ design stays. Users coordinate so one person drives the dashboard at a time.
- Design consequence: projection SAVE must still be safe if two people save around the same time → save/load only touches rows for the current project key; who/when stamp on each saved row. Optionally a visible "Dashboard in use by <user> since <time>" indicator (to confirm later).
- Tension noted: Q2 said "often at once"; Q3 accepts turn-taking for the dashboard itself.

### Q4 — Driving selector and data sources
- Asked: Which selector drives the new area and which tables does it read?
- Captured: **E6 Source Project / F6 Main IO drives it.** Reads tblBudgetSelectedTest (budget lines), qManpower_SelectedTest (MP projection), tblActualsSelectedTest (actuals), tblSelectedProjectMapping (Project Start/End). C5 is retired for the new area.
- KPI cards (C14, F14, F17, C17, C19) NOT repointed in this scope — they stay on original tables for now.
- Flags: KPI cards will then disagree with the timeline when C5 ≠ E6 → decide later whether/when to repoint KPIs (likely at promotion). -> user

### Q5 — Project key for persistence
- Asked: What identifies a project in the projection store?
- Captured: **Source Project Name (E6) + Main IO (F6)** composite key.

### Q6 — Budget line-item grain
- Asked: How do budget records become rows?
- Captured: **Merge by IO Code + Expense Category + Budget Description** (description matched trimmed, case-insensitive). Row shows net signed Revised Budget. E.g. OM 'Therapy Associate Service' 0 + 30,600 → one row $30,600.
- Line-item identity for the store = Project key + IO Code + Expense Category + normalised Budget Description.
- Flags: a description edited/renamed in the SOA source would orphan saved projections for that line → handling to be decided (Q later).

### Q7 — Zero-net budget lines
- Asked: Show merged lines whose net budget is $0?
- Captured: **Always hide** $0 net lines.
- Flags: if a projection was previously saved against a line that later nets to $0, it would be hidden → must decide whether it still counts in totals / is warned about (ties into orphan-handling question).

### Q8 — Ad-hoc / unbudgeted projection lines
- Asked: Allow projection lines not in the budget?
- Captured: **No — budget lines only.** Manual projections can only be entered against lines present in tblBudgetSelectedTest. Unbudgeted spend needs a budget line in SOA first.

### Q9 — Category set
- Asked: Is OTHER a real category or is OM 'Other Operating Expenses'?
- Captured: **Keep a separate OTHER** category — some projects use it in SOA. Show it when budget lines exist for it. Categories in scope: MP, OM, EQ, CO, OTHER.
- Flags: exact SOA CodeType value for OTHER (e.g. 'OTHER', 'OT', other?) not seen in current data -> user to confirm. Behaviour when actuals exist for a category with no budget lines -> to ask.

### Q10 — Manpower line identity
- Asked: How to identify/label each MP line given Assignment ID / Source Row ID change per month and no name/position is extracted?
- Captured: **Group in the macro; no PQ change.** One MP line = Staff Group + IO Code + Tag Start (LOA start) + Tag End (LOA end) + Resolved Rate Grade. qManpower_SelectedTest / fnManpowerSheet_SelectedTest stay frozen.
- Accepted risk: two positions with identical Staff Group, IO, LOA dates and grade merge into one row (FTE and $ summed — totals still correct).
- Flags: label format to confirm (proposed 'Research Assistant | AH11_JUL25 (RA) | Nov 2025–Mar 2027 | TF25HIP134MP').

### Q11 — MP section structure
- Asked: How do MP budget lines (RA, Therapist, Research Fellow) relate to manpower lines (Research Assistant, Allied Health / Nurse)?
- STATUS: Reopened and re-confirmed: **Option 3 'for now'** (user's words).
- Captured: **Both, side by side, no mapping.** Within the MP block: (a) MP budget lines from tblBudgetSelectedTest showing budget only; (b) manpower lines from qManpower_SelectedTest showing monthly projection (FTE × rate). No mapping between them is inferred.
- Flags: user said 'for now' — design should leave room to move to parent/child (budget line → manpower lines via an admin mapping table) later without rework.

### Q12 — Manual entry on MP rows
- Asked: Can users type manual projections on MP rows?
- Captured: **MP is read-only.** MP projection comes only from manpower lines (Monthly FTE × Monthly Rate). MP budget lines show budget only. All MP cells locked. Manual entry applies only to OM, EQ, CO, OTHER budget lines.

### Q13 — Where actuals appear
- Asked: Actuals carry only Expense Category + IO + Posting Month + SAP cost-element text (e.g. 'Sal -  Ald Health', 'CPF - Allied Health', 'Skill Dev Fund') — no link to budget lines/manpower lines. Category level, or line level via an admin rule table?
- Initially user described 'sum all cost related to the RA' (= line-level actual). After seeing the data, decided: **"stick with broad category — going into specific line items into detail and labelling will be too challenging."**
- Captured: **Actuals at category level only.** Category subtotal rows show Projected / Actual / Variance; no per-line actuals, no transaction-text rule table.
- Note: OM actuals in Esko Combo are payroll lines ('Sal -  Ald Health') — they will sit in OM category actuals as-is.
- Flags: confirm whether 'too challenging' also applies to the projection/budget rows (Q14).

### Q14 — Rows inside each category block
- Asked: Does 'too challenging' apply only to actuals, or to projection rows too?
- Captured: **Actuals only.** "projections per line item is possible, the comparisons to actual will just be on categorical level which is fine by me."
- Row model per category: line-item rows (budget + monthly projection; typed for OM/EQ/CO/OTHER, calculated read-only for MP manpower lines; MP budget lines budget-only) → category subtotal block: Projected (sum of lines), Actual, Variance ($), Variance (%).

### Q15 — Column layout
- Asked: How to arrange generated columns?
- Captured: **Option 1 — months, then a quarter total after each quarter's months**, then Project Total at the end. Boundary quarters contain only active months.
- **Quarters collapsible**: Excel outline grouping so each quarter collapses to its total and expands to show its months.
- **Each quarter colour-coded** so quarters are easy to tell apart.
- **Extra columns: yes** — financial-year total columns in addition to quarter totals and Project Total.
- Flags: FY-total placement, colour scheme, and whether 'Remaining budget' column is included -> Q16.

### Q16 — FY totals, colours, end columns, calendar basis
- (a) FY totals: **all FY total columns together at the far right** (before Project Total) — "keep the month and quarter view free from clutter".
- (b) Colours: **fixed colour per quarter number** (Q1 always same colour, Q2 another, …); months lighter tint of their quarter.
- (c) End columns: user wrote "project total is enlightened" — interpreted as **"Project Total is enough"** → no Budget-remaining / Budget−Actual columns. (Confirm.)
- (d) Calendar Year basis: **same layout applies** ('2026 | Q1' labels; calendar-year totals replace FY totals).
- Resulting column order: [labels] | Budget | months + quarter totals (collapsible, colour-coded) | FY (or CY) totals | Project Total.

## Open flags (pending input)
- Confirm (c) 'enlightened' = 'enough' → no Budget-remaining columns -> user
- MP may later move to parent/child with a Staff Group→budget-line mapping table -> user (future)
- MP row label format -> user
- Exact SOA CodeType string used for OTHER -> user
- Actuals in a category that has no budget lines (no rows to show them against) -> to ask
- Orphaned projections when a budget description changes/disappears in source — handling? -> user (to ask)
- KPI cards stay on original tables/C5 while timeline uses selected-test/E6 — potential mismatch; decide repoint timing -> user
- "In use by" indicator / soft lock on the dashboard — wanted? -> user
