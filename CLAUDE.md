# Clinical Research Cash Flow Projection

Project instructions for Claude. This file is a **directory, plus the goals and rules**. The detail lives in the files it points to; read those before acting.

## Goals

1. Give clinical research project leads a **project-level cash-flow view** in one Excel workbook: approved budget, signed actuals, manpower projections (FTE × rate) and typed projections, by month, quarter, financial year and project total.
2. Pull all source data automatically with **Power Query** from SharePoint (SOA and IREx files), driven by a project selector, without changing the source files.
3. Replace the old fixed 60-column timeline with a **macro-built timeline** sized to each project, with per-project saved projections and projected-vs-actual by category (spec v1.0, approved 11 Oct 2026).
4. Reach a UAT-signed v0.9, then promote the selected-test queries over the original tables one at a time.

## Current state (12 October 2026)

- **Baseline:** `Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.8.xlsm` (never modify). **Working copy:** `..._v0.9_working.xlsm`, **in UAT** in desktop Excel (spec v1.0.7).
- The user's copy has the selector in **row 5** (VBA `SEL_*` constants). The build script still writes row 6, so align it before any rebuild. The release zip has **no workbook** for now (modules and docs only).
- **Done and tested in Excel by the user (12 Oct):** rows 2–19 grouped on every build (UAT 17) and the **Remarks** column after Project Total, saved per project in `ProjectionStore!W:AD` (UAT 18). Stage 1 (narrower Line column) was **withdrawn by the user** (empty F:G). Future Actual/Variance stay blank (decided). Open: UAT 1–16 results, then promotion.
- **Next:** the latest handover in `handovers/README.md`.

## Directory: every file, and when to read it

Read only what the task needs; don't scan the whole repo.

| File | What it is | Read when |
|---|---|---|
| `handovers/README.md` | Index of session handovers, routine and template | **Every session start**: open the LATEST handover it lists |
| `handovers/2026-10-12-stage1-label-width-and-row-group.md` | Latest handover: row group and Remarks done (UAT 17–18 pass), Stage 1 withdrawn | Every session start (until a newer one exists) |
| `handovers/2026-10-12-session-handover-structure-and-ui-requests.md` | Earlier 12 Oct handover: Stage 1–3 findings and options | Background on the UI requests |
| `handovers/2026-10-12-v0.9-design-build-uat.md` | History of 11–12 Oct: design, build, fixes (was `HANDOVER.md`) | You need background on how something came to be |
| `specs/2026-10-11-cashflow-dynamic-timeline-spec.md` | **Authoritative build spec** (layout, macros, store, KPI cards, §13 tests, §14 decisions) | Any change to the timeline, macros or workbook |
| `docs/PROJECT_REFERENCE.md` | Architecture, design summary, display conventions, data facts and risks (§4), checkpoints, techniques, variance cut-off (§7) | Data questions, reconciling figures, how to read or modify the workbook |
| `brainstorms/2026-10-11-cashflow-dynamic-timeline.md` | Q&A log of design decisions (Q1–Q34) | Someone asks why a design choice was made |
| `docs/Clinical_Research_Cash_Flow_PRD_v1.9.docx` | Product requirements (v1.9) | Updating requirements or the change log |
| `docs/Clinical_Research_Cash_Flow_TECHNICAL_SPECIFICATION.md` | Queries, manpower routing, rate rules, timeline tech spec | Power Query or manpower logic questions |
| `docs/Clinical_Research_Cash_Flow_Build.md` | Build status log | Recording what was built or changed |
| `docs/incidents/2026-10-11-v0.9-querytable9-repair.md` | Incident (closed): Excel repair from an XML rename | Renaming anything, packaging, repair prompts |
| `docs/incidents/2026-10-11-v0.9-project-end-date.md` | Defect (closed): timeline used SOA LOA dates | Project dates, timeline length |
| `docs/v0.9_Install_and_User_Guide.md` / `.docx` | User install and usage guide | Changing install steps or user-facing behaviour |
| `tests/UAT_test_script.md` | 18 acceptance tests for desktop Excel | UAT results, adding tests |
| `tests/grid_model.py` | Python port of the grid; LibreOffice recalc; 44 checks | Any grid-logic change (keep in step with `modTimeline`) |
| `vba/modTimeline.bas` | Load Project, refresh, checks, project dates, grid builder, layout offsets, **cell constants** | Timeline layout or behaviour changes |
| `vba/modStore.bas` | Save, conflict check, unsaved edits, Unmatched re-assign/delete | Saving, the store, Remarks-type persistence |
| `vba/modSetup.bas` | One-off setup: buttons, protection | Buttons or protection changes |
| `vba/ThisWorkbook_code.txt` | `Workbook_Open` (outline under protection) | Startup behaviour |
| `build/build_v09_working.py` | Builds v0.9 from v0.8 by package-XML edits (C7/C8, sheets, names) | Rebuilding the workbook |
| `build/verify_v09_working.py` | Verifies a build (PQ byte-identical, cross-part links) | After every build |
| `build/compare_pq.py` | Test 13: PQ unchanged after an Excel save | A user returns a saved workbook |
| `release/CashFlow_v0.9_package.zip` | Release for download: current modules, docs, tests, handovers; **no workbook** for now | Packaging or download questions |
| `prototypes/cashflow-timeline-prototype.html` | Approved visual prototype (https://claude.ai/artifact/AYqLAxM4obRV3eJAXDBLDv) | Visual layout questions |
| `.gitattributes` / `.gitignore` | CRLF for `.bas`/`.txt`; workbooks ignored | Line-ending or commit problems |

**Not in the repo:** the `.xlsm` workbooks and the SharePoint source files, because they hold internal data and SharePoint URLs. Ask the user to upload the current workbook in each chat. Don't commit workbooks unless the user asks.

## Rules: data and business (non-negotiable)

- Never modify the source workbooks. Never overwrite v0.8. Save a new workbook version before any material change.
- Keep signed budget and actual values, including negative adjustments.
- Manpower values are **FTE**. Projected cost = Monthly FTE × Monthly Rate from `qManpowerRates` / `tbl_ManpowerRate` (the latest rate effective on or before the month). **Never hardcode rates.**
- **Never infer IO relationships** by adding, removing or interpreting suffixes. The manpower IO comes from the source as-is. Budget Code Type MP is only an eligibility check.
- **Project dates come from `qProjects_SelectedTest` (IREx).** The `tblSelectedProjectMapping` dates are SOA budget LOA dates and only a fallback.
- **Actuals are compared at category level only.** Actual/Variance stop at the Latest SOA Month (`docs/PROJECT_REFERENCE.md` §7).
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
- Log every defect or incident in `docs/incidents/` (symptom, evidence, root cause, fix, prevention).
- **End of every session: write a new handover** `handovers/YYYY-MM-DD-<topic>.md` from the template in `handovers/README.md`, add it to the index as LATEST, and update this file's Current state and directory. Don't rewrite old handovers.
- **Whenever a handover is asked for, first update every document the session's changes touch.** Check each file in the directory above against what changed, including the spec, PRD (bump the version and add a change-log row), technical spec, `docs/PROJECT_REFERENCE.md`, build log, UAT script, install guide (md, then regenerate the docx with pandoc), brainstorm log, incidents and release zip. Say which docs were updated and which were checked and needed nothing.
- When adding a file to the repo, add it to the directory table above.
- Keep this file under 200 lines. Put detail in `docs/PROJECT_REFERENCE.md` or the spec.
- The user works partly on a phone. Keep replies short, and use numbered options when asking (answers often come as "a1, b1").

## Starting a session

1. Read the **latest handover** (via `handovers/README.md`), then only the files the task needs, using the directory above.
2. Ask for the current workbook (and v0.8 if a rebuild is needed).
3. Check `git log origin/main` for work from other sessions before changing anything.
