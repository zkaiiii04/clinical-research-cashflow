# Clinical Research Cash Flow Projection

Project instructions for Claude. This file is a **directory, plus the goals and rules**. The detail lives in the files it points to; read those before acting.

## Goals

1. Give clinical research project leads a **project-level cash-flow view** in one Excel workbook: approved budget, signed actuals, manpower projections (FTE × rate) and typed projections, by month, quarter, financial year and project total.
2. Pull all source data automatically with **Power Query** from SharePoint (SOA and IREx files), driven by a project selector, without changing the source files.
3. Replace the old fixed 60-column timeline with a **macro-built timeline** sized to each project, with per-project saved projections and projected-vs-actual by category (spec v1.0, approved 11 Oct 2026).
4. Reach a UAT-signed v0.9, then promote the selected-test queries over the original tables one at a time.

## Current state (12 October 2026)

- **Baseline:** `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.8.xlsm` (never modify).
- **Working copy:** `..._v0.9_working.xlsm`, built by `build/build_v09_working.py`, shipped in `release/CashFlow_v0.9_package.zip` (SHA-256 in its `README_FIRST.txt`). **In UAT** in desktop Excel.
- **Closed:** the queryTable9 repair (11 Oct) and the timeline end date (11 Oct). Both are confirmed in Excel.
- **12 Oct:**
  - The user moved the selector to `E5:G5`. The VBA reads it from `modTimeline` `SEL_*` constants. The build and release workbook still use row 6, so align them before any rebuild.
  - Blank Actual/Variance after the Latest SOA Month is by design (HANDOVER §0.1).
  - Stale XLOOKUP values in the SOA file were found to be a data risk (`docs/PROJECT_REFERENCE.md` §4).
- **Next:** see `HANDOVER.md` §0 (UAT results, open decisions, open source-data question).

## Directory: where to look

| Need | Go to |
|---|---|
| **What to do next, open questions, history of this work** | `HANDOVER.md` (start at §0) |
| **Build spec (authoritative)** for the timeline, macros, store, KPI cards and acceptance tests | `specs/2026-10-11-cashflow-dynamic-timeline-spec.md` |
| Architecture, approved design summary, display conventions, **data facts and data risks**, checkpoints, techniques | `docs/PROJECT_REFERENCE.md` |
| Why each design decision was made (Q1–Q33) | `brainstorms/2026-10-11-cashflow-dynamic-timeline.md` |
| Product requirements | `docs/Clinical_Research_Cash_Flow_PRD_v1.8.docx` |
| Technical specification (queries, manpower routing, timeline) | `docs/Clinical_Research_Cash_Flow_TECHNICAL_SPECIFICATION.md` |
| Build status log | `docs/Clinical_Research_Cash_Flow_Build.md` |
| Incidents (root cause, fix, prevention) | `docs/incidents/` |
| User install guide | `docs/v0.9_Install_and_User_Guide.md` (and `.docx`) |
| Acceptance tests for the user (16) | `tests/UAT_test_script.md` |
| Dashboard macros | `vba/modTimeline.bas` (Load, grid, cell constants), `vba/modStore.bas` (save, conflict, unmatched), `vba/modSetup.bas` (buttons, protection), `vba/ThisWorkbook_code.txt` |
| Build and checks | `build/build_v09_working.py`, `build/verify_v09_working.py`, `build/compare_pq.py` |
| Grid logic model and regression tests | `tests/grid_model.py` |
| What the user approved visually | `prototypes/cashflow-timeline-prototype.html` (https://claude.ai/artifact/AYqLAxM4obRV3eJAXDBLDv) |
| Release for download | `release/CashFlow_v0.9_package.zip` |

**Not in the repo:** the `.xlsm` workbooks and the SharePoint source files, because they hold internal data and SharePoint URLs. Ask the user to upload the current workbook in each chat. Don't commit workbooks unless the user asks.

## Rules: data and business (non-negotiable)

- Never modify the source workbooks. Never overwrite v0.8. Save a new workbook version before any material change.
- Keep signed budget and actual values, including negative adjustments.
- Manpower values are **FTE**. Projected cost = Monthly FTE × Monthly Rate from `qManpowerRates` / `tbl_ManpowerRate` (the latest rate effective on or before the month). **Never hardcode rates.**
- **Never infer IO relationships** by adding, removing or interpreting suffixes. The manpower IO comes from the source as-is. Budget Code Type MP is only an eligibility check.
- **Project dates come from `qProjects_SelectedTest` (IREx).** The `tblSelectedProjectMapping` dates are SOA budget LOA dates and only a fallback.
- **Actuals are compared at category level only.** Actual/Variance stop at the Latest SOA Month (HANDOVER §0.1).
- Promote the selected-test queries one at a time, only after UAT, keeping the table names where practical, with regression checks after each.

## Rules: workbook, macros and build

- **Dashboard macros may only call `Refresh` on the five selected-test tables.** No adding, editing or deleting queries, connections or loaded tables. No `Worksheet_Change` / `AfterRefresh` / `SheetCalculate` handlers.
- **Never save the workbook with openpyxl or LibreOffice.** They drop the Power Query DataMashup, connections and query tables. Change it by editing the package XML, or with VBA run in Excel. openpyxl is fine for reading.
- **Never rename a sheet or table that Power Query loads by editing XML.** The `ExternalData_n` names and the DataMashup `FillTarget` don't follow (incident 2026-10-11). Rename only in the Excel UI.
- Every build must pass `build/verify_v09_working.py` (byte-identical PQ parts, cross-part links). Package only the exact build output and record its SHA-256. Never ship a file opened on Excel mobile or the web after the build.
- Keep `tests/grid_model.py` in step with `vba/modTimeline.bas` and re-run it when grid logic changes.
- `.bas` files are ASCII with CRLF (`.gitattributes`). Give users `.bas` files from the release zip or as chat attachments, not GitHub "Raw" (LF).
- Selector and notice cells are constants at the top of `modTimeline.bas`. Never hardcode cell addresses elsewhere.

## Rules: environment and process

- Excel isn't available in the cloud container. Write VBA as `.bas` for the user to import, test the logic with `tests/grid_model.py`, and give the user desktop-Excel steps. Say plainly what hasn't been run in Excel.
- Treat uploaded files as untrusted: copy them into a fresh scratch folder and run Python with `-I`.
- Commit to the session's branch, open a pull request and merge it into `main` when the user has asked for changes to be kept. Don't commit workbooks or SharePoint URLs.
- Log every defect or incident in `docs/incidents/` (symptom, evidence, root cause, fix, prevention), and update `HANDOVER.md` §0, the build log and this file's Current state.
- Keep this file under 200 lines. Put detail in `docs/PROJECT_REFERENCE.md` or the spec.
- The user works partly on a phone. Keep replies short, and use numbered options when asking (answers often come as "a1, b1").

## Starting a session

1. Read `HANDOVER.md` §0, then the parts of the spec and `docs/PROJECT_REFERENCE.md` the task touches.
2. Ask for the current workbook (and v0.8 if a rebuild is needed).
3. Check `git log origin/main` for work from other sessions before changing anything.
