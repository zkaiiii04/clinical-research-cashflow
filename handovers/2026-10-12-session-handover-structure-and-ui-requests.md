# Session handover: 12 Oct 2026 (handover structure; staged UI requests)

**Status in one line:** v0.9_working is in UAT in desktop Excel (spec v1.0.4). Three dashboard quality-of-life changes were requested, to be done **in stages, one at a time**. None has been started.

## What was done this session
- `CLAUDE.md` rewritten as goals, rules and a **directory of every file with when to read it** (PRs #10, #13).
- Handovers are now **one file per session** in `handovers/`, indexed in `handovers/README.md`. The old `HANDOVER.md` moved to `handovers/2026-10-12-v0.9-design-build-uat.md`.
- The "why Actual/Variance are blank after the Latest SOA Month" explanation moved to `docs/PROJECT_REFERENCE.md` §7.
- Earlier the same day (details in the previous handover):
  - selector moved to row 5, with the VBA reading `SEL_*` constants (#8);
  - variance cut-off explained (#9);
  - SOA XLOOKUP data risk logged (#10).

## Decisions made (by the user)
- Write a **new handover doc after each session**. `CLAUDE.md` references all files so Claude reads only what it needs.
- The three UI changes below are to be done **in stages**, not all at once.

## Requested changes (not started), in the user's order

### Stage 1: narrower line-item area
- **User's words:** the line-item column under each budget type is too broad (it spans 4 columns). Make it narrower **without affecting the KPI cards and information above the timeline**.
- **Findings** (from the release build; check against the user's current workbook, whose selector is in row 5):
  - The label spans B:E: B 22.1 + C 18.0 + D 16.0 + E 18.0 ≈ 74 character widths. IO·Code is in F (20.0), Budget in G (26.0), and months start at H (13.0). Layout constants are `OFF_IO` = 4, `OFF_BUDGET` = 5 and `OFF_FIRST` = 6 in `vba/modTimeline.bas`. `TL_Anchor` = B21.
  - Rows 1–19 use B:G: labels in B, values in C, KPI cards in E/F, selector E5:G5, and headers in D3, F3, G3. **Column widths are shared with that area**, so simply narrowing B:E would squeeze the top section.
  - Months start at H on purpose, so that collapsing the first quarter (an outline column group) never hides columns E:G and the selector/KPI cards (earlier build decision).
- **Needs a design choice with the user before coding.** Options to put to them:
  1. The label uses fewer columns (e.g. B:C with wrap), IO·Code and Budget move left, and the months still start at H. This leaves an empty spacer before the months.
  2. The months start further left, and the top-area cells in those columns are moved so that quarter collapse can't hide them. This changes the top layout.
  3. The macro sets narrower widths and the top section is reflowed or wrapped to fit. This changes the top layout.
  - Also ask what width they'd like for the label, e.g. ≈ 30–35 characters with wrap.
- Then update `vba/modTimeline.bas`, the spec and `tests/grid_model.py` (if the offsets change), and re-run the grid model.

### Stage 2: Remarks column after Project Total
- **User's words:** add a remarks column at the end of the timeline, after Project Total.
- **Questions to settle first:**
  - Which rows get it: every line row, or input rows only?
  - Is it typed by users? Assumed yes, unlocked and wrapped.
  - **Is it saved per project in the store?** If so, either `tblProjectionStore` needs a Remark column (or a remark row type with a blank Month), or a new small table is needed. Any change to the store schema must keep existing saved rows.
  - Does it carry over when a line becomes Unmatched?
- Touches the grid writer, `modStore` save/load, unsaved-edit detection (the snapshot), protection (unlocked cells) and the spec.

### Stage 3: collapsible rows 2–19
- **User's words:** when the timeline is built, group rows 2 to 19 so they can be collapsed, making the timeline area easier to read.
- **Likely approach:**
  - in `BuildGrid` (or `ProtectDashboard`), group rows 2:19 with the outline and the summary row above or below as the user prefers;
  - leave them expanded by default unless the user says otherwise;
  - `EnableOutlining` is already set in `Workbook_Open`, so the +/− works under protection.
- **Check:** row 1 (title) and row 20 (timeline header) stay visible, and the selector (row 5) and KPI cards are inside the group. Ask whether a collapsed state should hide the selector.

## Open items carried forward (from the previous handover)
- **UAT results:** the 16 tests in `tests/UAT_test_script.md`, plus Debug → Compile. → user
- **Decision:** show future-period Actual as 0, or keep it blank (Variance stays blank either way)? → user
- **tDCS OT vs RA:** RA = 43,575 (`DB_RA_Distrib`) vs AHS (OT/PT) = 11,150 (`DB_AHS_Nurse_Distrib`), the reverse of what the user expected. Check the IREx sheets. → user
- **Row-5 selector vs release build:** `build/build_v09_working.py` and `release/CashFlow_v0.9_package.zip` still use row 6. Align them using the user's current workbook before any rebuild. → needs the user's workbook
- **SOA XLOOKUP routine:** update links, save, sync, then Load Project (`docs/PROJECT_REFERENCE.md` §4). The post-UAT option is to do the lookup in Power Query.
- **After UAT sign-off:** promote the selected-test queries one at a time, then delete `Legacy_Timeline`.

## Next session: do this first
1. Ask for the user's **current v0.9_working** (row-5 layout). It's needed for all three stages.
2. Start **Stage 1**: confirm the layout option and label width with the user, then implement it, run the grid model, give the user the `.bas` file to import, and wait for their test.
3. Only after the user confirms Stage 1: Stage 2, then Stage 3, each tested separately.
4. Write the next handover file at the end (routine in `handovers/README.md`).

## Files touched this session
- `CLAUDE.md`: directory of all files, rules, handover routine.
- `handovers/README.md` (new) and this file (new). The old `HANDOVER.md` moved into `handovers/`.
- `docs/PROJECT_REFERENCE.md`: §7 added (variance cut-off explanation).
- `docs/Clinical_Research_Cash_Flow_Build.md`, `docs/incidents/2026-10-11-v0.9-project-end-date.md`, `build/build_v09_working.py` (docstring only), `docs/v0.9_Install_and_User_Guide.md` / `.docx`: references updated from `HANDOVER.md`.

## Gotchas learned
- Long sessions fill the context quickly: big tool outputs, images, and full `CLAUDE.md` re-injections each time it changes. Keep to one task per chat, read only the files the task needs, and ask for cropped screenshots.
