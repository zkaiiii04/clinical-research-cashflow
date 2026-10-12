# Session handovers

One handover document per working session, newest first. **Read only the latest one** at the start of a session. Older ones are history, to read only when a task needs that background.

| Date | File | Summary |
|---|---|---|
| 12 Oct 2026 | [`2026-10-12-stage1-label-width-and-row-group.md`](2026-10-12-stage1-label-width-and-row-group.md) | **LATEST.** Stage 1 (Line column B:C) tried and withdrawn by the user (empty F:G). Rows 2–19 grouped on every build, and the Remarks column (Stage 2) written; both waiting on the user's Excel test (UAT 17–18). |
| 12 Oct 2026 | [`2026-10-12-session-handover-structure-and-ui-requests.md`](2026-10-12-session-handover-structure-and-ui-requests.md) | New per-session handover routine. Three staged dashboard changes requested (line-item width, Remarks column, collapsible rows 2–19), none started. Carried-forward open items. |
| 11–12 Oct 2026 | [`2026-10-12-v0.9-design-build-uat.md`](2026-10-12-v0.9-design-build-uat.md) | Design discovery → spec v1.0 → v0.9 build → queryTable9 and end-date fixes → selector moved to row 5 → SOA XLOOKUP data risk. (Was `HANDOVER.md`.) |

## Routine (every session)

1. **Start:** read `CLAUDE.md`, then the latest handover here, then only the files the task needs (`CLAUDE.md` lists every file and when to read it).
2. **End:** write a **new** file `handovers/YYYY-MM-DD-<short-topic>.md` from the template below. Add it at the top of the table above, mark it **LATEST**, and remove LATEST from the previous one. Don't rewrite earlier handovers.
3. Also update `CLAUDE.md` "Current state" (3–6 lines) and the build log if anything was built or changed.

## Template

```markdown
# Session handover: <date> (<topic>)

**Status in one line:** <where the work stands>

## What was done this session
- <change, with file paths and PR numbers>

## Decisions made (by the user)
- <decision>

## Open items / waiting on the user
- <item> → <who / what is needed>

## Next session: do this first
1. <first step>

## Files touched
- <path>: <why>

## Gotchas learned
- <anything a future session must not repeat>
```
