# Cash Flow Dashboard — Dynamic Projection Timeline: Brainstorm / Discovery Notes
Date: 2026-10-11 · Goal: Pin down every requirement for replacing the fixed, formula-filled timeline on Main_Dashboard (G21:BO56) with a macro-built, project-specific, line-item timeline that supports manual projections with per-project save/load, without interfering with Power Query.

## Summary / key decisions
(running synthesis, updated as answers come in)

- **Platform:** single shared .xlsm on SharePoint, desktop Excel → VBA OK; projections stored in-workbook.
- **Driver/data:** E6/F6 + selected-test tables (budget, manpower, actuals, mapping). KPIs repointed to new grid; C5 retired (Q22).
- **Non-MP line grain:** one row per IO + Category + normalised Budget Description, net budget.
- **Zero-net lines:** always hidden.
- **Ad-hoc lines:** not allowed; budget lines only.
- **Categories:** MP, OM, EQ, CO, OTHER (OTHER is separate; shown when budget lines exist).
- **MP line identity:** macro groups by Staff Group + IO Code only (Q25b); AHS/Nurse split into 'Allied Health' and 'Nurse' via rate category list (Q27b); label 'Staff Group | IO'; no PQ change.
- **Charge-to:** admin table maps manpower lines to MP/OM; OM-charged manpower shows as locked row in OM block (Q27c).
- **MP structure:** (option 3, 'for now') MP budget lines (budget only) and manpower lines (projection) listed separately in the MP block; no mapping.
- **MP manual entry:** none — MP read-only; manual entry only on OM/EQ/CO/OTHER budget lines.
- **Actuals:** category level only (no line-level actuals).
- **Rows:** line items carry budget + projection; projected-vs-actual comparison on category subtotal rows only.
- **Columns:** Budget | months → quarter total (collapsible, fixed colour per quarter number) … | all FY/CY totals at far right | Project Total. No Budget-remaining column (pending confirm). Same layout for Calendar Year.
- **Cell contents:** values for data (MP, actuals, manual inputs) + grid-local SUM/variance formulas for totals; no table SUMIFS.
- **Trigger:** 'Load Project' button = refresh selected-test queries synchronously → consistency check → build → load saved projections. No events; no query edits.
- **Save:** replace-this-project rows; prompt on unsaved switch; conflict warning on later save by someone else; blank not stored, 0 stored.
- **Unmatched saved projections:** kept, shown in an 'Unmatched' row per category, counted in totals; out-of-range months in Project Total only (with note); Re-assign / Delete actions available.
- **Patient reimbursement:** typed $ on the OM budget line; tblPatientReimbursement retired.
- **Migration:** build in v0.9_working; old area + old input tables kept on hidden sheet for UAT comparison, deleted later.
- **Past months:** projection = plan as entered; no forecast row; no locking.
- **Protection:** sheet protected, no password; inputs unlocked + shaded; store sheet very hidden; numeric-only validation (pending confirm).
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

### Q17 — Cell contents
- Asked: What does the macro write into generated cells?
- Captured: **Option 1 — values for data, simple local formulas for totals.** Macro writes plain values for MP manpower projection, category actuals and loaded manual projections; writes short grid-local formulas (=SUM of the row's months for quarter/FY/Project totals; subtotal = SUM of line rows; Variance $ = Actual − Projected; Variance % = IFERROR(Var/Projected,"")) so totals update live as users type. No SUMIFS against PQ tables in the grid.
- Consequence: MP/actual values are a snapshot from the last build → rebuild/reload after a PQ refresh (trigger design in Q18).

### Q18 — Build trigger and PQ interaction
- Asked: What triggers the build; does the macro touch PQ?
- Captured: **One 'Load Project' button.** Steps: (1) check Setup Status = Ready; (2) refresh ONLY the selected-project queries in dependency order, synchronously (qSelectedProjectMapping → qBudget_SelectedTest, qActuals_SelectedTest, qProjects_SelectedTest, qManpower_SelectedTest); (3) verify every selected table reflects the same E6/F6 project — stop with a message if not; (4) build grid; (5) load saved projections for the project key.
- Macro only CALLS refresh — never edits/adds/deletes queries, connections or loaded tables. No event-driven rebuilds.
- Separate 'Save Projections' button (Q19).
- Flags: qProjectIOCatalogue (feeds the E6/F6 dropdowns) not refreshed by Load Project — refreshed separately/as today -> confirm.

### Q19 — Save behaviour
- (a) **Replace this project's rows only.** 'Save Projections' deletes rows for the current project key (Source Project + Main IO) from the hidden store and writes the grid back. Store row = Source Project, Main IO, IO Code, Expense Category, Budget Description (normalised), Month, Amount, Saved By, Saved At. Other projects untouched. No version history.
- (b) **Unsaved edits on Load Project for another project → prompt 'Save changes to <project> first? Yes / No / Cancel'.**
- (c) **Conflict check: yes.** On save, if the store's latest Saved At for this project is later than when this user loaded it → warn '<name> saved this project at <time> after you loaded it. Overwrite?'.
- (d) **Blank ≠ 0.** Blank month not stored; typed 0 stored as 0.

### Q20 — Saved projections that no longer match the grid
- Cases: budget description renamed/removed in SOA; line now nets to $0 (hidden per Q7); Project End changed so saved months fall outside the timeline.
- Captured: **Option 2 — keep them and count them in totals.** Shown in an **'Unmatched' row under each category** (read-only), so category subtotal / Project Total still include them.
- Flags: saved months outside the new project date range have no month column → proposed: show their amount in the Unmatched row's Project Total (and FY total if the FY exists) only -> confirm. Whether users can re-assign/delete unmatched rows from the dashboard -> confirm.

### Q20b — Unmatched row behaviour
- (a) **Saved months outside project dates → counted in the Unmatched row's Project Total only**; month detail kept visible via a cell note/comment. No extra month columns.
- (b) **Users can clean up**: each Unmatched row offers **Re-assign** (pick a current budget line in the same category) and **Delete** via a macro button/form.

### Q21 — Patient reimbursement
- Captured: **Option 1 — typed as dollars on the OM 'Patient Reimbursemnt' budget line like any other line; retire tblPatientReimbursement** (patients × rate calculator dropped).
- Flags: KPI F17 'Projected Total' currently sums tblPatientReimbursement + tblManualProjection → must be re-pointed when those tables retire (ties to KPI flag) -> Q22.

### Q22 — Old area, old input tables, KPI cards
- Captured: **Option 2.** Build in a new working version (v0.9_working; v0.8 untouched as fallback). Old fixed timeline (G21:BO56), tblManualProjection and tblPatientReimbursement moved to a **hidden comparison sheet for UAT**, deleted later. **KPI cards repointed** to the new grid / E6-F6 (Revised Approved Budget, Actual to Date, Projected Total, Overall & Projected Utilisation); C5 retired from the dashboard.
- Note: KPI mismatch flag from Q4 resolved by this answer.

### Q23 — Closed (past) months
- Captured: **Option 1 — projection stays as entered in every month**, past or future. Projected = plan; Projected vs Actual for past months shows plan vs reality. Projected Total KPI = sum of plan. No Forecast row, no locking of past months.

### Q24 — Protection and input validation
- Captured: **Option 2 — protect Main_Dashboard without a password**; only OM/EQ/CO/OTHER line-item month cells unlocked and shaded as inputs; E6/F6 selectors editable; outline +/- works under protection; hidden store sheet set to xlSheetVeryHidden. Guards against accidents, not deliberate edits.
- Input validation: user replied "years" — interpreted as **"yes"** → decimal numbers only, negatives allowed. (Confirm.)

### Q25 — Loose-ends batch (partial)
- (1) OTHER code: user: "there is one called AL, B, C, D". Catalogue check (tblProjectIOCatalogue) found Budget Code Types beyond MP/OM/EQ/CO: **AL** (D.R.E.A.M.T. (Presage)), **OT** (EEG (RESTORE 2.0)), **B, C, D** (Uptale: Grief and Loss Training), and **blank** (D.R.E.A.M.T., FHT 2.0). How these map to blocks -> follow-up Q25a.
- (2) **Yes** — a category with actuals but no budget lines still shows its block with subtotal rows (Actual / Variance) so spend is never hidden.
- (3) User asked how the MP row label is derived → explained; raised that Resolved Rate Grade can change month-to-month when a new rate becomes effective, which would split one position into two rows → proposed dropping grade from the grouping key -> follow-up Q25b.
- (4)–(9) not yet answered.

### Q25b — MP grouping key and label (supersedes Q10 key)
- Captured: **MP line = Staff Group + IO Code only.** Remove Tag Start/End (LOA) and Resolved Rate Grade from both the grouping key and the label, so a rate change mid-LOA never splits a position. Label e.g. 'Research Assistant | TF25HIP134MP'.
- Consequence: all positions of the same Staff Group on the same IO merge into one row (FTE × rate summed per month) — accepted.

### Q26 — Edge case raised by user: staff-group actuals are lumped in SOA
- User: Research Assistant and Therapist actual costs both fall under the 'Allied Health' group in SOA (e.g. 'Sal -  Ald Health', 'CPF - Allied Health'), so if a project budgets for both, projections exist per staff group but actuals are lumped. "How should we resolve this?"
- Response given: already resolved by Q13/Q14 — projected-vs-actual is compared at **category (MP) level only**: MP subtotal Projected = sum of ALL manpower lines (RA + Allied Health / Nurse); MP Actual = sum of ALL MP actuals regardless of cost-element text. No per-staff-group actual is attempted, so the lumping doesn't create a mismatch.
- Related risk raised: Esko Combo shows payroll lines coded **OM** ('Sal -  Ald Health' under OM). If a person on the manpower sheet is paid through OM, their projection lands in MP but actuals land in OM → MP looks under-spent, OM over-spent. -> follow-up question to user.

### Q26 — Manpower staff paid under OM (answer)
- Captured: **Yes, this happens** — some manpower-sheet staff are charged to OM in SOA. User OK to add a per-line "charged to OM" setting.
- User question: if a manpower line is set to OM, will OM show it like manpower? → Proposed answer: yes — the line moves into the OM block as a read-only calculated row (FTE × rate) next to OM's typed budget lines; OM subtotal Projected includes it and compares to OM actuals. Mechanism / granularity -> follow-up Q27.

### Q26c — Nursing (raised by user)
- User: "nursing costs are handled separately under nursing category in actual SOA spending — is this covered?"
- Issue: manpower Staff Group 'Allied Health / Nurse' comes from ONE source sheet (DB_AHS_Nurse_Distrib), so AHS and nurse projections are on one MP row. Whether covered depends on how SOA codes nursing actuals:
  - if nursing actuals are CodeType MP (only cost-element text differs) → covered by MP-level comparison;
  - if nursing actuals have their own CodeType (a separate category) → NOT covered: projection in MP, actuals in another block.
- Possible fix if separate: split the AHS/Nurse row by the rate's Rate Category (from qManpowerRates) and treat nursing as its own category/charge target. -> follow-up Q27.

### Q27 — Nursing code, AHS/Nurse split, charge-to setting
- (a) **Nursing actuals use CodeType MP** (SOA states this clearly) → nursing is covered by the MP-level projected-vs-actual comparison.
- (b) **27b1 — split 'Allied Health / Nurse' into 'Allied Health' and 'Nurse' rows using the rate's Rate Selected / Rate Category**, via a small admin list of which rate categories count as nursing. No PQ change. (Since nursing is MP, this split is for clarity/labelling and for the charge-to setting, not for actuals matching.) MP line key becomes Staff Group (split) + IO Code.
- (c) **27c1 — 'charge to' setting lives in an admin table** on a settings sheet: Source Project, Main IO, Staff Group, IO Code → Charge To (MP / OM / …). Macro reads it at Load Project; a manpower line charged to OM appears as a locked FTE × rate row in the OM block and counts in OM subtotal.
- Limitation stated: the setting is per Staff Group + IO line, not per person. User has not yet said whether that's acceptable.

## Open flags (pending input)
1. Charge-to granularity: OK that it's per Staff Group + IO (not per person)? Do OM-paid staff sit on a different IO? -> user
2. List of rate categories/grades that count as 'Nurse' (for AHS/Nurse split) -> user / admin to supply
3. How codes AL / OT / B / C / D / blank map to dashboard blocks (Q25a) -> user
4. Confirm 'enlightened' = 'enough' → no Budget-remaining column -> user
5. Load Project does not refresh qProjectIOCatalogue (dropdown list) — leave to Refresh All? -> user
6. 'Dashboard loaded by <user> at <time>' notice (no lock) — wanted? -> user
7. Variance sign: keep Actual − Projected (positive = overspend)? -> user
8. Manpower FTE months outside project dates → MP Unmatched row, Project Total only? -> user
9. Confirm 'years' = 'yes' for numeric-only validation on input cells -> user
10. Anything not covered: charts, print/export, number format (whole $ vs cents), who may save -> user
- Future (not blocking): MP may later move to parent/child with a Staff Group → budget-line mapping table.
