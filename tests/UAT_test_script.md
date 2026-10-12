# v0.9_working: install and acceptance test script (desktop Excel)

For `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.9_working.xlsm`, built from v0.8 by `build/build_v09_working.py`. Spec: `specs/2026-10-11-cashflow-dynamic-timeline-spec.md` §13.

You need desktop Excel for Windows (Microsoft 365). The macros don't run in Excel for the web, and Mac Excel lacks `Scripting.Dictionary`.

Don't open or save the workbook in LibreOffice or Google Sheets. A LibreOffice save removes the Power Query connections.

---

## Part A: install (once)

The illustrated version of these steps is in `docs/v0.9_Install_and_User_Guide.md` (and the .docx in the package).

1. Save `..._v0.9_working.xlsm` somewhere local first, not on SharePoint yet. If Windows shows "This file came from another computer", right-click the file → Properties → tick **Unblock** → OK.
2. Open it in desktop Excel and click **Enable Content**. If Excel offers to repair the file instead, click **No**, stop, and report it (a screenshot of the message, plus the repair log if one was shown). Don't save a repaired copy.
3. Press **Alt+F11** to open the VBA editor, then **File → Import File…** and import, one at a time:
   - `vba/modTimeline.bas`
   - `vba/modStore.bas`
   - `vba/modSetup.bas`
4. In the Project pane, double-click **ThisWorkbook** (under Microsoft Excel Objects). Paste in the code from `vba/ThisWorkbook_code.txt`, from `Private Sub Workbook_Open()` to `End Sub`.
5. Click **Debug → Compile VBAProject**. Nothing should happen, meaning no error message.
   - If a line is highlighted, send me the module name, the highlighted line and the message. **Stop here.**
6. Close the VBA editor. Press **Alt+F8**, choose `Setup_Dashboard_v09` and click **Run**. You should see "Setup complete".
   - Four buttons appear at E8:F11: Load Project, Save Projections, Re-assign unmatched, Delete unmatched.
7. **File → Save** (keep `.xlsm`), close Excel, then reopen the file and click **Enable Content**.

What changed in the file (built by script; Power Query untouched):
- `Main_Dashboard` rows 21 and below are empty until the first Load. C5 is gone. KPI cards C14/F14/F17/C17/C19 are blank until the first Load. C7/C8/C9 come from the selected-test tables. E7 says "Dashboard not loaded yet…". C6 now has a Financial Year / Calendar Year dropdown.
- New sheets:
  - `Dashboard_Settings` (`tblManpowerChargeTo`)
  - `Legacy_Timeline` (hidden; the full v0.8 dashboard, still driven by its own C5, for comparison)
  - `ProjectionStore` (very hidden)
- The manpower table and sheet keep their v0.8 names, `qManpower_SelectedTest` / `PQManpower_SelectedTest`. The stray `=SUM(...)` under the table is removed.

---

## Part B: acceptance tests (spec §13)

Record **Pass / Fail** for each, plus a screenshot or the message text for any failure.

**Expected figures for HoW2.0 - 3D Gait** (Main IO `TF25HIP134`, manpower IO `TF25HIP134MP`):

| Figure | Expected |
|---|---|
| Budget (PROJECT TOTAL, Budget column; KPI C14) | **$250,000.00** |
| Actual to Date (PROJECT TOTAL Actual, Project Total; KPI F14) | **$68,563.67** |
| MP Actual | **$24,367.83** |
| OM Actual | **$44,195.84** |
| MP Projected (`RA (Projected)`, `TF25HIP134MP`) | **$62,250.00** |
| Projected Total (KPI F17) before typing anything | $62,250.00 plus any saved projections |
| Overall Utilisation (C17) | 27.4% (68,563.67 / 250,000) |
| Months | 17 (Nov 2025 – Mar 2027) |

### 1. Load (FY basis)
1. Set C6 = Financial Year, E6 = HoW2.0 - 3D Gait, F6 = TF25HIP134. G6 should show Ready.
2. Click **Load Project**. It takes a few seconds while the five tables refresh.

Expected:
- Month columns run from **Nov 25 to Mar 27**, 17 in all, starting at column H.
- Each quarter is followed by a quarter total: `FY25/Q3` (2 months: Nov, Dec), `FY25/Q4`, `FY26/Q1` … `FY26/Q4`.
- At the far right are `FY25 Total`, `FY26 Total`, then `Project Total`.
- E7 shows "Dashboard loaded by <you> at <time>".

### 2. Calendar Year basis
1. Set C6 = Calendar Year and click **Load Project**.

Expected:
- Quarters run `2025 | Q4` (2 months), `2026 | Q1` … `2027 | Q1`.
- Totals are `2025 Total`, `2026 Total`, `2027 Total`, then Project Total.

Set C6 back to Financial Year and load again.

### 3. Budget reconciliation
- **PROJECT TOTAL Projected** row, **Budget** column (G) = **250,000**, and KPI C14 = **$250,000**.

### 4. Actuals reconciliation
- **PROJECT TOTAL Actual**, Project Total column = **68,564** (KPI F14 shows **$68,563.67**).
- **MP Actual** = **24,368**.
- **OM Actual** = **44,196**.

### 5. Manpower reconciliation
- The MP block has `RA (Projected)` with IO·Code `TF25HIP134MP`. Its Project Total = **62,250**.
- **MP Projected** Project Total = **62,250**.
- The RA row's month cells are grey and locked; typing there is refused.

### 6. Zero-net line (HoW2.0 - Esko Combo)
1. Load HoW2.0 - Esko Combo (E6, then its Main IO in F6, then **Load Project**).

Expected:
- **Research Fellow** (39,000 − 39,000) does **not** appear.
- **Therapy Associate Service** appears **once**, in OM, with Budget **30,600**.

### 7. Save → switch → reload
1. With Esko Combo loaded, type in yellow OM cells:
   - HBRA, Jan 26 = `500`
   - HBRA, Feb 26 = `0`
   - Leave HBRA, Mar 26 blank.
   - Patient Reimbursemnt, Jun 26 = `-120`
2. Click **Save Projections**. You should see "Saved 3 projection value(s)…".
3. Load 3D Gait, then load Esko Combo again.

Expected:
- HBRA Jan 26 = 500, Feb 26 = **0** (not blank), Mar 26 **blank**.
- Patient Reimbursemnt Jun 26 = **(120)** in red.
- OM Projected and PROJECT TOTAL include them.

### 8. Unsaved switch
1. With Esko Combo loaded, type `999` in any yellow cell. Don't save.
2. Change E6/F6 to 3D Gait and click **Load Project**. You get "Save changes to HoW2.0 - Esko Combo first?" with Yes / No / Cancel.
3. Click **Cancel**. The Esko grid stays, with 999 still there.
4. Click **Load Project** again and choose **No**. 3D Gait loads and the 999 is discarded.
5. Load Esko again. The 999 cell is blank.

### 9. Conflict warning
This simulates "someone else saved after you loaded".
1. Load Esko Combo and click **Save Projections** once.
2. Press **Alt+F11**, then **Ctrl+G** for the Immediate window. Paste this line and press Enter:
   `ThisWorkbook.Names.Add Name:="TL_LoadedAt", RefersTo:="=""1""", Visible:=False`
3. Back in Excel, change a yellow cell and click **Save Projections**.

Expected:
- "<your name> saved this project at <time> after you loaded it. Overwrite?"
- **No** saves nothing; **Yes** saves.

On two PCs, the real test is: A loads, B loads and saves, A saves, and A gets the warning.

### 10. Unmatched rows, Re-assign and Delete
1. Load Esko Combo, type HBRA Jan 26 = `500` and HBRA Sep 26 = `200`, and save.
2. Simulate a renamed SOA line:
   1. **Alt+F11**. In the Project pane, select the **ProjectionStore** sheet. In the Properties window (F4), set **Visible** = `-1 - xlSheetVisible`.
   2. In Excel, on ProjectionStore, change **Line Key** from `HBRA` to `HBRA OLD` on the HBRA rows only.
   3. Back in the VBA editor, set Visible = `2 - xlSheetVeryHidden`.
3. Click **Load Project** for Esko Combo.

Expected:
- OM shows an italic row **"Unmatched: HBRA (no longer in budget)"** holding 500 (Jan 26) and 200 (Sep 26).
- OM Projected and PROJECT TOTAL include them.
- Load Project warns about 1 unmatched line.

4. Click a cell in that Unmatched row, then **Re-assign unmatched**. When asked, click any cell in the **HBRA** row and confirm.
   - The amounts move onto HBRA and the Unmatched row disappears.
5. Repeat step 2 to make another Unmatched row. Select it and click **Delete unmatched**, then **Yes**.
   - The row disappears and the totals drop by those amounts.

### 11. Charge-to
1. On **Dashboard_Settings**, in `tblManpowerChargeTo`, enter one row:
   - Source Project `HoW2.0 - Esko Combo`
   - Main IO `TF25HIP136`
   - Rate Category `AHS` (or `AHS (OT/PT)`, whichever Esko's AHS line shows)
   - IO Code: the manpower IO shown in that line's IO·Code cell
   - Charge To `OM`
2. Load Esko Combo.

Expected:
- That line leaves MP and appears in OM as **"AHS (Projected) ⚙ manpower → OM"**, grey and locked.
- MP Projected drops and OM Projected rises by the same amount. PROJECT TOTAL is unchanged.

Delete the settings row afterwards.

### 12. Collapse under protection
1. Save, close and reopen the file, then Load any project.
2. Click the **−** above a quarter, beside the column letters. Its months collapse to the quarter total. **+** expands them.
3. Try typing in a total cell or a label. Excel refuses because the sheet is protected.
4. Typing in yellow cells and changing C6/E6/F6 still work.

### 13. Power Query untouched
1. After tests 1–12, save the workbook and send it back to me.
2. I run `python3 -I build/compare_pq.py v0.8.xlsm <your file>`. It checks that the M code, the connections and every query-table destination are identical.
   - Also check yourself: **Data → Queries & Connections** lists the same queries as v0.8.

### 14. Mismatch guard
1. Load 3D Gait. Change E6 to Esko Combo (and F6), but **don't** click Load Project.
2. Click **Save Projections**.

Expected: **"Selection changed — reload first."** Nothing is saved.

### 15. Cross-project
1. Load HoW2.0 - Esko Combo.

Expected:
- Blocks appear in the order MP, OM, EQ, CO.
- MP: RA (72,600) and Therapist (88,000) budget lines, plus its manpower lines (AHS and RA rows as the source has them).
- OM: Patient Reimbursemnt 9,000, HBRA 4,000, CRIO/ Biostatistics Support 4,000, Publication 3,000, Therapy Associate Service 30,600.
- EQ: Technological Development 38,500.
- CO: Stationary, Alcohol wips, Paper, Files 300.
- IO·Code reads `TF25HIP136MP` / `TF25HIP136OM` / `TF25HIP136EQ` / `TF25HIP136CO`, merged down each run.
- Its payroll actuals coded OM show in **OM Actual**.

### Note on blank cells
Actual, Variance ($) and Variance (%) are deliberately blank for months, quarters and FY columns entirely after the Latest SOA Month (C9). Nothing has been posted for them yet, so a variance would show the whole future plan as underspend. They fill in after the next SOA month is posted and you click Load Project. This is expected, not a failure.

### 16. Project dates come from the project, not the SOA budget LOA (tDCS)
1. Select tDCS (Main IO `TF24HIP134`) and click Load Project.

Expected (if the SOA budget LOA dates for tDCS still end 31 Dec 2026; if you've corrected them in the SOA file, the warning simply doesn't appear):
- A Load warning: "The SOA budget LOA dates (1 Jan 2025 - 31 Dec 2026) differ from the project dates in qProjects_SelectedTest (1 Jan 2025 - 30 Sep 2027). The timeline uses the project dates…".
- C7/C8 show 1 Jan 2025 and 30 Sep 2027.
- The timeline runs Jan 2025 to Sep 2027 (33 months; FY24 Total to FY27 Total at the far right).
- No "Unmatched: manpower months outside project dates" row.
- The two manpower lines' Project Totals add up to MP Projected, and MP Projected equals the total of `Projected Amount ($)` in `qManpower_SelectedTest` (54,725 with the data on 11 Oct 2026).

---

## What to send back

- The test numbers with Pass / Fail.
- For any failure: a screenshot and the exact message.
- The saved workbook, for test 13. Upload it here; it isn't committed to the repo.
