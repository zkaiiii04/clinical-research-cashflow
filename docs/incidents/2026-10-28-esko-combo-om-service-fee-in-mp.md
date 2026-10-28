# Defect (open): Esko Combo OM manpower service fee shows in the MP block

**Status:** open, noted for a future change. No code changed yet.
**Reported:** 28 Oct 2026 by the user.
**Project:** HoW2.0 - Esko Combo (Main IO `TF25HIP136`).

## Symptom
Esko Combo has a **manpower service fee** charged to `TF25HIP136OM`. The IREx manpower source (the database Excel) has the right IO, `TF25HIP136OM`, but the dashboard shows the line in the **MP (manpower) block**, not in OM.

## Evidence
- User report, 28 Oct 2026. The source IO is correct, so the source isn't the problem.
- `vba/modTimeline.bas` `ReadManpower`: every manpower row starts with `blk = "MP"`. Only a matching `tblManpowerChargeTo` row (Rate Category + IO Code) moves it to another block. The IO Code itself isn't used to choose the block.

## Root cause
Working as designed (spec §5.3, §7): **every manpower line goes to MP unless the charge-to table says otherwise.** This follows the CLAUDE.md rule never to infer IO relationships from suffixes, so the macro doesn't read "OM" at the end of `TF25HIP136OM` as meaning the OM block. As a result, a manpower-source line on a non-MP IO lands in MP and counts in the MP subtotal.

Two more limits of the current mechanism:
- Charge-to only applies to **rate-matched** lines. A "Rate not found" row always stays in MP (`RNF|…` key, no charge-to lookup).
- Charge-to is per Rate Category + IO, so the service fee can be moved only if it has its own Rate Category or IO. Here the IO `TF25HIP136OM` is distinct, so a row can target it.

## Workaround (no code change)
Add a row to `Dashboard_Settings!tblManpowerChargeTo`:

| Source Project | Main IO | Rate Category | IO Code | Charge To |
|---|---|---|---|---|
| HoW2.0 - Esko Combo | TF25HIP136 | *(the line's Rate Category, as shown on the dashboard)* | TF25HIP136OM | OM |

Then click Load Project. The line moves to OM as a locked `… (Projected) ⚙ manpower → OM` row. This won't work if the line shows "⚠ Rate not found".

## Fix (to decide)
Options recorded for a future change. The user hasn't chosen yet.
1. **Keep the charge-to table** as the only routing mechanism and document the Esko Combo row (workaround above).
2. **Default the block from the catalogue:** if the manpower IO exactly equals a catalogue IO + Budget Code Type pair for the selected project (e.g. `TF25HIP136` + `OM`), default the line to that block. The charge-to table still overrides it. This needs the user's sign-off against the "never infer IO relationships" rule, because it matches the IO·Code display convention rather than the source.
3. Also let charge-to apply to "Rate not found" lines.

Any fix changes `ReadManpower` / `ReadChargeTo` in `vba/modTimeline.bas`, so `tests/grid_model.py`, spec §5.3/§7 and the UAT test 11 must change with it.

## Prevention
When a new project is set up, check whether any manpower-source IO isn't an MP IO, and add its charge-to row before UAT.
