Attribute VB_Name = "modStore"
'==============================================================================
' modStore - Save Projections, Re-assign / Delete unmatched, projection store
'
' Spec: specs/2026-10-11-cashflow-dynamic-timeline-spec.md (v1.0) sections 8-10
'
' tblProjectionStore (very hidden ProjectionStore sheet), one row per saved
' month: Source Project | Main IO | IO Code | Expense Category | Budget Code |
' Line Key | Budget Description | Month | Amount | Saved By | Saved At.
' A project is keyed by Source Project + Main IO. Saving replaces only that
' project's rows. Blank cells are not stored; 0 is stored.
'
' Session state is kept in hidden names (TL_LoadedProject, TL_LoadedMainIO,
' TL_LoadedAt, TL_InputHash, TL_GridSpan). Unsaved edits are found by
' comparing a hash of the input cells with the hash taken at load or save, so
' no sheet events are needed.
'
' The grid layout (which sheet row is which line) is written to
' ProjectionStore!M:T by Load Project and read back here.
'
' Remarks (spec 5.5) are kept in a plain range on ProjectionStore, W:AD, one
' row per line that has a remark: Source Project | Main IO | Line Ref | Block |
' Description | Remark | Saved By | Saved At. tblProjectionStore is unchanged.
' Line Ref is the line's remark key (modTimeline TLine.RemarkRef). Every saved
' remark is shown on Load (on its line, or on an Unmatched row), so Save
' replaces the project's remarks with the Remarks column as a whole.
'==============================================================================
Option Explicit

Public Const S_SRC As Long = 1
Public Const S_MAINIO As Long = 2
Public Const S_IO As Long = 3
Public Const S_CAT As Long = 4
Public Const S_CODE As Long = 5
Public Const S_KEY As Long = 6
Public Const S_DESC As Long = 7
Public Const S_MONTH As Long = 8
Public Const S_AMT As Long = 9
Public Const S_BY As Long = 10
Public Const S_AT As Long = 11
Private Const S_COLS As Long = 11

Private Const LAYOUT_COL As Long = 13       ' ProjectionStore!M
Private Const LAYOUT_COLS As Long = 8       ' Row, Kind, Block, IO, Code, Line Key, Description, Remark Ref

Public Const R_SRC As Long = 1
Public Const R_MAINIO As Long = 2
Public Const R_REF As Long = 3
Public Const R_BLOCK As Long = 4
Public Const R_DESC As Long = 5
Public Const R_TEXT As Long = 6
Public Const R_BY As Long = 7
Public Const R_AT As Long = 8
Private Const R_COLS As Long = 8
Private Const REMARK_COL As Long = 23       ' ProjectionStore!W

Private Const NAME_PROJECT As String = "TL_LoadedProject"
Private Const NAME_MAINIO As String = "TL_LoadedMainIO"
Private Const NAME_AT As String = "TL_LoadedAt"
Private Const NAME_HASH As String = "TL_InputHash"
Private Const NAME_SPAN As String = "TL_GridSpan"

Private mSavedCount As Long
Private mSavedRemarks As Long

'==============================================================================
' Button: Save Projections (spec 10.2)
'==============================================================================
Public Sub SaveProjections()
    On Error GoTo Fail
    If SaveCore(False) Then
        MsgBox "Saved " & mSavedCount & " projection value(s) and " & mSavedRemarks & _
               " remark(s) for " & LoadedProject() & ".", vbInformation, "Save Projections"
    End If
    Exit Sub
Fail:
    MsgBox "Save Projections stopped: " & Err.Description, vbExclamation, "Save Projections"
End Sub

' fromLoad = True when Load Project saves the previous project before
' switching: E6/F6 already show the new project, so the selection check is
' skipped and the loaded key is used.
Public Function SaveCore(ByVal fromLoad As Boolean) As Boolean
    Dim src As String, mainIO As String, lastBy As String, lastAt As Date
    Dim ws As Worksheet, data As Variant, n As Long, lay As Variant, nLay As Long
    Dim inputSet As Object, monthSet As Object, mCols() As Long, mDates() As Date, nM As Long
    Dim out() As Variant, nOut As Long, i As Long, j As Long, c As Long, v As Variant
    Dim key As String, d As Date, nowStamp As Date, who As String, skipped As Long

    src = LoadedProject()
    mainIO = LoadedMainIO()
    If src = "" Then
        MsgBox "Load a project first.", vbExclamation, "Save Projections"
        Exit Function
    End If
    If Not fromLoad Then
        If Not SelectionMatchesLoaded() Then
            MsgBox "Selection changed " & ChrW(&H2014) & " reload first." & vbCrLf & vbCrLf & _
                   "The timeline shows " & src & " (" & mainIO & "), but the project selector now shows a different project. " & _
                   "Click Load Project, or set the selector back to save.", vbExclamation, "Save Projections"
            Exit Function
        End If
    End If
    If SavedSinceLoad(src, mainIO, lastBy, lastAt) Then
        If MsgBox(lastBy & " saved this project at " & Format$(lastAt, "d mmm yyyy h:mm AM/PM") & _
                  " after you loaded it. Overwrite?", vbYesNo + vbExclamation + vbDefaultButton2, _
                  "Save Projections") <> vbYes Then Exit Function
    End If

    Set ws = modTimeline.DashSheet()
    nLay = ReadLayout(lay)
    nM = GridMonths(mCols, mDates)
    Set inputSet = InputIdentities(lay, nLay)
    Set monthSet = MonthSetOf(mDates, nM)
    n = ReadStoreData(data)
    ReDim out(1 To n + nLay * nM + 1, 1 To S_COLS)
    nowStamp = Now
    who = Application.UserName

    ' Keep other projects' rows and this project's unmatched rows
    For i = 1 To n
        If RowIsKey(data, i, src, mainIO) Then
            key = StoreIdentity(data, i)
            If modTimeline.ToDate(data(i, S_MONTH), d) Then
                If inputSet.Exists(key) And monthSet.Exists(CLng(MonthStart(d))) Then GoTo NextRow
            End If
        End If
        nOut = nOut + 1
        For j = 1 To S_COLS
            out(nOut, j) = data(i, j)
        Next j
NextRow:
    Next i

    ' One row per non-blank input cell
    mSavedCount = 0
    For i = 1 To nLay
        If lay(i, 2) = "INPUT" Then
            For c = 1 To nM
                v = ws.Cells(CLng(lay(i, 1)), mCols(c)).Value2
                If Not modTimeline.IsBlank(v) Then
                    If IsNumeric(v) And Not IsError(v) Then
                        nOut = nOut + 1
                        out(nOut, S_SRC) = src
                        out(nOut, S_MAINIO) = mainIO
                        out(nOut, S_IO) = lay(i, 4)
                        out(nOut, S_CAT) = lay(i, 3)
                        out(nOut, S_CODE) = lay(i, 5)
                        out(nOut, S_KEY) = lay(i, 6)
                        out(nOut, S_DESC) = lay(i, 7)
                        out(nOut, S_MONTH) = mDates(c)
                        out(nOut, S_AMT) = CDbl(v)
                        out(nOut, S_BY) = who
                        out(nOut, S_AT) = nowStamp
                        mSavedCount = mSavedCount + 1
                    Else
                        skipped = skipped + 1
                    End If
                End If
            Next c
        End If
    Next i

    WriteStore out, nOut
    mSavedRemarks = SaveRemarks(src, mainIO, lay, nLay, who, nowStamp)
    SetNameText NAME_AT, Trim$(Str$(CDbl(Now)))
    SetNameText NAME_HASH, HashInputs()
    If skipped > 0 Then MsgBox skipped & " cell(s) held text instead of a number and were not saved.", _
                               vbExclamation, "Save Projections"
    SaveCore = True
End Function

'==============================================================================
' Buttons: Re-assign / Delete unmatched (spec 9, 10.3)
'==============================================================================
Public Sub ReassignUnmatched()
    Dim lay As Variant, nLay As Long, sel As Long, tgt As Long, target As Range
    Dim src As String, mainIO As String, moved As Long, fromKey As String, toKey As String
    Dim remMoved As Boolean
    On Error GoTo Fail
    If Not PrepareUnmatched("Re-assign unmatched", lay, nLay, sel) Then Exit Sub
    src = LoadedProject(): mainIO = LoadedMainIO()

    If CountInputs(lay, nLay, CStr(lay(sel, 3))) = 0 Then
        MsgBox "There are no budget lines in " & lay(sel, 3) & " to move these amounts to.", _
               vbExclamation, "Re-assign unmatched"
        Exit Sub
    End If

    On Error Resume Next
    Set target = Application.InputBox( _
        "Click any cell in the " & lay(sel, 3) & " budget line that should receive these amounts, then OK." & _
        vbCrLf & vbCrLf & "Amounts in months inside the project dates move to that line, adding to " & _
        "anything already there. Months outside the dates stay unmatched.", _
        "Re-assign unmatched", Type:=8)
    On Error GoTo Fail
    If target Is Nothing Then Exit Sub
    If target.Worksheet.Name <> modTimeline.DASH_SHEET Then
        MsgBox "Click a budget line on " & modTimeline.DASH_SHEET & ".", vbExclamation, "Re-assign unmatched"
        Exit Sub
    End If
    tgt = LayoutIndexOfRow(lay, nLay, target.Cells(1, 1).Row)
    If tgt = 0 Then
        MsgBox "That cell isn't on a budget line.", vbExclamation, "Re-assign unmatched"
        Exit Sub
    End If
    If lay(tgt, 2) <> "INPUT" Or lay(tgt, 3) <> lay(sel, 3) Then
        MsgBox "Choose a budget line in the " & lay(sel, 3) & " block.", vbExclamation, "Re-assign unmatched"
        Exit Sub
    End If

    fromKey = modTimeline.Identity(CStr(lay(sel, 3)), CStr(lay(sel, 5)), CStr(lay(sel, 4)), CStr(lay(sel, 6)))
    toKey = modTimeline.Identity(CStr(lay(tgt, 3)), CStr(lay(tgt, 5)), CStr(lay(tgt, 4)), CStr(lay(tgt, 6)))
    moved = MoveRows(src, mainIO, fromKey, lay, nLay, tgt)
    remMoved = MoveRemark(src, mainIO, LayRef(lay, sel), LayRef(lay, tgt), CStr(lay(tgt, 3)), CStr(lay(tgt, 7)))
    If moved = 0 And Not remMoved Then
        MsgBox "Nothing to move: these saved amounts are all in months outside the project dates. " & _
               "You can delete them instead.", vbInformation, "Re-assign unmatched"
        Exit Sub
    End If
    modTimeline.BuildGrid False
    MsgBox moved & " saved amount(s)" & IIf(remMoved, " and the remark", "") & " moved to """ & _
           lay(tgt, 7) & """.", vbInformation, "Re-assign unmatched"
    Exit Sub
Fail:
    MsgBox "Re-assign stopped: " & Err.Description, vbExclamation, "Re-assign unmatched"
End Sub

Public Sub DeleteUnmatched()
    Dim lay As Variant, nLay As Long, sel As Long, removed As Long, label As String
    Dim remDeleted As Boolean
    On Error GoTo Fail
    If Not PrepareUnmatched("Delete unmatched", lay, nLay, sel) Then Exit Sub
    label = CStr(modTimeline.DashSheet().Cells(CLng(lay(sel, 1)), ThisWorkbook.Names("TL_Anchor").RefersToRange.Column).Value)
    If MsgBox(label & vbCrLf & vbCrLf & "Delete this row's saved amounts and remark? This can't be undone.", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Delete unmatched") <> vbYes Then Exit Sub
    removed = DeleteRows(LoadedProject(), LoadedMainIO(), _
        modTimeline.Identity(CStr(lay(sel, 3)), CStr(lay(sel, 5)), CStr(lay(sel, 4)), CStr(lay(sel, 6))), lay, nLay)
    remDeleted = DeleteRemark(LoadedProject(), LoadedMainIO(), LayRef(lay, sel))
    modTimeline.BuildGrid False
    MsgBox removed & " saved amount(s)" & IIf(remDeleted, " and the remark", "") & " deleted.", _
           vbInformation, "Delete unmatched"
    Exit Sub
Fail:
    MsgBox "Delete stopped: " & Err.Description, vbExclamation, "Delete unmatched"
End Sub

' Common checks: a project is loaded, E6/F6 still match it, the active cell is
' on an Unmatched row, no unsaved edits (offers to save), nobody saved since.
Private Function PrepareUnmatched(ByVal title As String, ByRef lay As Variant, _
                                  ByRef nLay As Long, ByRef sel As Long) As Boolean
    Dim lastBy As String, lastAt As Date
    If LoadedProject() = "" Then
        MsgBox "Load a project first.", vbExclamation, title
        Exit Function
    End If
    If Not SelectionMatchesLoaded() Then
        MsgBox "Selection changed " & ChrW(&H2014) & " reload first.", vbExclamation, title
        Exit Function
    End If
    nLay = ReadLayout(lay)
    sel = 0
    If ActiveSheet Is modTimeline.DashSheet() Then sel = LayoutIndexOfRow(lay, nLay, ActiveCell.Row)
    If sel > 0 Then
        If lay(sel, 2) = "MP_OOR" Then
            MsgBox "These manpower months come from the manpower source, so they can't be " & _
                   "re-assigned or deleted here. Correct the dates in the source, or the project dates.", _
                   vbInformation, title
            Exit Function
        End If
        If lay(sel, 2) <> "UNMATCHED" Then sel = 0
    End If
    If sel = 0 Then
        MsgBox "Select a cell in an Unmatched row first (the italic rows that start with ""Unmatched:"").", _
               vbExclamation, title
        Exit Function
    End If
    If IsDirty() Then
        If MsgBox("You have unsaved projection changes. Save them first and continue?", _
                  vbOKCancel + vbQuestion, title) <> vbOK Then Exit Function
        If Not SaveCore(False) Then Exit Function
    ElseIf SavedSinceLoad(LoadedProject(), LoadedMainIO(), lastBy, lastAt) Then
        MsgBox lastBy & " saved this project at " & Format$(lastAt, "d mmm yyyy h:mm AM/PM") & _
               " after you loaded it. Click Load Project first to see their changes.", vbExclamation, title
        Exit Function
    End If
    PrepareUnmatched = True
End Function

' Moves this project's unmatched rows for fromKey that fall in a month column
' to the target line, adding to any amount already saved for that month.
Private Function MoveRows(ByVal src As String, ByVal mainIO As String, ByVal fromKey As String, _
                          ByVal lay As Variant, ByVal nLay As Long, ByVal tgt As Long) As Long
    Dim data As Variant, n As Long, i As Long, j As Long, d As Date, moved As Long
    Dim inputSet As Object, monthSet As Object, mCols() As Long, mDates() As Date, nM As Long
    Dim merged As Object, k As String, out() As Variant, nOut As Long, idx As Long

    nM = GridMonths(mCols, mDates)
    Set inputSet = InputIdentities(lay, nLay)
    Set monthSet = MonthSetOf(mDates, nM)
    n = ReadStoreData(data)
    If n = 0 Then Exit Function
    Set merged = CreateObject("Scripting.Dictionary")
    ReDim out(1 To n, 1 To S_COLS)

    For i = 1 To n
        If RowIsKey(data, i, src, mainIO) And modTimeline.ToDate(data(i, S_MONTH), d) Then
            If StoreIdentity(data, i) = fromKey And Not inputSet.Exists(fromKey) _
               And monthSet.Exists(CLng(MonthStart(d))) Then
                data(i, S_IO) = lay(tgt, 4)
                data(i, S_CAT) = lay(tgt, 3)
                data(i, S_CODE) = lay(tgt, 5)
                data(i, S_KEY) = lay(tgt, 6)
                data(i, S_DESC) = lay(tgt, 7)
                data(i, S_BY) = Application.UserName
                data(i, S_AT) = Now
                moved = moved + 1
            End If
            ' Merge rows that now share line + month
            k = StoreIdentity(data, i) & "|" & CLng(MonthStart(d))
            If merged.Exists(k) Then
                idx = merged(k)
                out(idx, S_AMT) = modTimeline.ToNumber(out(idx, S_AMT)) + modTimeline.ToNumber(data(i, S_AMT))
                If StampOf(data(i, S_AT)) > StampOf(out(idx, S_AT)) Then
                    out(idx, S_BY) = data(i, S_BY)
                    out(idx, S_AT) = data(i, S_AT)
                End If
                GoTo NextRow
            End If
            merged.Add k, nOut + 1
        End If
        nOut = nOut + 1
        For j = 1 To S_COLS
            out(nOut, j) = data(i, j)
        Next j
NextRow:
    Next i
    If moved > 0 Then
        WriteStore out, nOut
        SetNameText NAME_AT, Trim$(Str$(CDbl(Now)))
    End If
    MoveRows = moved
End Function

' Deletes this project's rows for key that are unmatched (line not shown, or
' month outside the month columns).
Private Function DeleteRows(ByVal src As String, ByVal mainIO As String, ByVal key As String, _
                            ByVal lay As Variant, ByVal nLay As Long) As Long
    Dim data As Variant, n As Long, i As Long, j As Long, d As Date, removed As Long
    Dim inputSet As Object, monthSet As Object, mCols() As Long, mDates() As Date, nM As Long
    Dim out() As Variant, nOut As Long, matched As Boolean

    nM = GridMonths(mCols, mDates)
    Set inputSet = InputIdentities(lay, nLay)
    Set monthSet = MonthSetOf(mDates, nM)
    n = ReadStoreData(data)
    If n = 0 Then Exit Function
    ReDim out(1 To n, 1 To S_COLS)
    For i = 1 To n
        If RowIsKey(data, i, src, mainIO) Then
            If StoreIdentity(data, i) = key Then
                matched = False
                If inputSet.Exists(key) And modTimeline.ToDate(data(i, S_MONTH), d) Then
                    matched = monthSet.Exists(CLng(MonthStart(d)))
                End If
                If Not matched Then
                    removed = removed + 1
                    GoTo NextRow
                End If
            End If
        End If
        nOut = nOut + 1
        For j = 1 To S_COLS
            out(nOut, j) = data(i, j)
        Next j
NextRow:
    Next i
    If removed > 0 Then
        WriteStore out, nOut
        SetNameText NAME_AT, Trim$(Str$(CDbl(Now)))
    End If
    DeleteRows = removed
End Function

'==============================================================================
' Store access
'==============================================================================
Private Function StoreTable() As ListObject
    Set StoreTable = modTimeline.FindTable("tblProjectionStore")
    If StoreTable Is Nothing Then Err.Raise vbObjectError + 520, , "tblProjectionStore was not found."
End Function

' All store rows (blank rows skipped) as data(1..n, 1..11). Returns n.
Public Function ReadStoreData(ByRef data As Variant) As Long
    Dim lo As ListObject, v As Variant, i As Long, j As Long, n As Long
    Set lo = StoreTable()
    If lo.DataBodyRange Is Nothing Then Exit Function
    v = lo.DataBodyRange.Value
    ReDim data(1 To UBound(v, 1), 1 To S_COLS)
    For i = 1 To UBound(v, 1)
        If modTimeline.CleanText(v(i, S_SRC)) <> "" Then
            n = n + 1
            For j = 1 To S_COLS
                data(n, j) = v(i, j)
            Next j
        End If
    Next i
    ReadStoreData = n
End Function

' Replaces the whole store with out(1..n). Text columns are formatted as text
' first so descriptions such as "Mar 26" stay text.
Private Sub WriteStore(ByRef out() As Variant, ByVal n As Long)
    Dim lo As ListObject, ws As Worksheet, oldRows As Long, newRows As Long, top As Range
    Dim body As Range, v() As Variant, i As Long, j As Long
    Set lo = StoreTable()
    Set ws = lo.Parent
    Set top = lo.HeaderRowRange.Cells(1, 1)
    oldRows = lo.ListRows.Count
    newRows = n
    If newRows < 1 Then newRows = 1
    lo.Resize ws.Range(top, top.Offset(newRows, S_COLS - 1))
    If oldRows > newRows Then ws.Range(top.Offset(newRows + 1, 0), top.Offset(oldRows, S_COLS - 1)).Clear
    Set body = lo.DataBodyRange
    body.ClearContents
    body.Columns(S_SRC).Resize(, S_DESC).NumberFormat = "@"
    body.Columns(S_BY).NumberFormat = "@"
    body.Columns(S_MONTH).NumberFormat = "mmm yyyy"
    body.Columns(S_AMT).NumberFormat = "#,##0.00"
    body.Columns(S_AT).NumberFormat = "dd mmm yyyy hh:mm:ss"
    If n > 0 Then
        ReDim v(1 To n, 1 To S_COLS)
        For i = 1 To n
            For j = 1 To S_COLS
                v(i, j) = out(i, j)
            Next j
        Next i
        body.Resize(n, S_COLS).Value = v
    End If
End Sub

Public Function RowIsKey(ByRef data As Variant, ByVal i As Long, ByVal src As String, ByVal mainIO As String) As Boolean
    RowIsKey = modTimeline.SameText(data(i, S_SRC), src) And modTimeline.SameText(data(i, S_MAINIO), mainIO)
End Function

' Block|Code|IO|LineKey for a store row (same form as the grid's line identity)
Public Function StoreIdentity(ByRef data As Variant, ByVal i As Long) As String
    Dim blk As String, code As String
    code = UCase$(modTimeline.CleanText(data(i, S_CODE)))
    blk = UCase$(modTimeline.CleanText(data(i, S_CAT)))
    Select Case blk
        Case "MP", "OM", "EQ", "CO", "OTHER"
        Case Else: blk = modTimeline.BlockOf(code)
    End Select
    StoreIdentity = modTimeline.Identity(blk, code, UCase$(modTimeline.CleanText(data(i, S_IO))), _
                                         modTimeline.NormKey(modTimeline.CleanText(data(i, S_KEY))))
End Function

' Latest Saved At for the project; True when it is after this session's load.
Private Function SavedSinceLoad(ByVal src As String, ByVal mainIO As String, _
                                ByRef lastBy As String, ByRef lastAt As Date) As Boolean
    Dim data As Variant, n As Long, i As Long, d As Date, best As Double, loadedAt As Double
    n = ReadStoreData(data)
    For i = 1 To n
        If RowIsKey(data, i, src, mainIO) Then
            If modTimeline.ToDate(data(i, S_AT), d) Then
                If CDbl(d) > best Then
                    best = CDbl(d)
                    lastBy = modTimeline.CleanText(data(i, S_BY))
                End If
            End If
        End If
    Next i
    n = ReadRemarks(data)
    For i = 1 To n
        If RemarkIsKey(data, i, src, mainIO) Then
            If modTimeline.ToDate(data(i, R_AT), d) Then
                If CDbl(d) > best Then
                    best = CDbl(d)
                    lastBy = modTimeline.CleanText(data(i, R_BY))
                End If
            End If
        End If
    Next i
    loadedAt = Val(NameText(NAME_AT))
    If best > 0 And best > loadedAt + 1 / 86400 Then
        lastAt = CDate(best)
        If lastBy = "" Then lastBy = "Someone"
        SavedSinceLoad = True
    End If
End Function

'==============================================================================
' Remarks (ProjectionStore!W:AD, spec 5.5)
'==============================================================================
' All remark rows (blank rows skipped) as rd(1..n, 1..8). Returns n.
Public Function ReadRemarks(ByRef rd As Variant) As Long
    Dim ws As Worksheet, last As Long, v As Variant, i As Long, j As Long, n As Long
    Set ws = ThisWorkbook.Worksheets(modTimeline.STORE_SHEET)
    last = ws.Cells(ws.Rows.Count, REMARK_COL).End(xlUp).Row
    If last < 2 Then Exit Function
    v = ws.Range(ws.Cells(2, REMARK_COL), ws.Cells(last, REMARK_COL + R_COLS - 1)).Value
    ReDim rd(1 To UBound(v, 1), 1 To R_COLS)
    For i = 1 To UBound(v, 1)
        If modTimeline.CleanText(v(i, R_SRC)) <> "" Then
            n = n + 1
            For j = 1 To R_COLS
                rd(n, j) = v(i, j)
            Next j
        End If
    Next i
    ReadRemarks = n
End Function

Public Function RemarkIsKey(ByRef rd As Variant, ByVal i As Long, ByVal src As String, ByVal mainIO As String) As Boolean
    RemarkIsKey = modTimeline.SameText(rd(i, R_SRC), src) And modTimeline.SameText(rd(i, R_MAINIO), mainIO)
End Function

' Replaces the remarks area with out(1..n) and rewrites its header row.
Private Sub WriteRemarkRows(ByRef out() As Variant, ByVal n As Long)
    Dim ws As Worksheet, last As Long, v() As Variant, i As Long, j As Long, body As Range
    Set ws = ThisWorkbook.Worksheets(modTimeline.STORE_SHEET)
    ws.Range(ws.Cells(1, REMARK_COL), ws.Cells(1, REMARK_COL + R_COLS - 1)).Value = _
        Array("Source Project", "Main IO", "Line Ref", "Block", "Description", "Remark", "Saved By", "Saved At")
    last = ws.Cells(ws.Rows.Count, REMARK_COL).End(xlUp).Row
    If last >= 2 Then ws.Range(ws.Cells(2, REMARK_COL), ws.Cells(last, REMARK_COL + R_COLS - 1)).Clear
    If n = 0 Then Exit Sub
    Set body = ws.Range(ws.Cells(2, REMARK_COL), ws.Cells(n + 1, REMARK_COL + R_COLS - 1))
    body.Columns(R_SRC).Resize(, R_BY).NumberFormat = "@"
    body.Columns(R_AT).NumberFormat = "dd mmm yyyy hh:mm:ss"
    ReDim v(1 To n, 1 To R_COLS)
    For i = 1 To n
        For j = 1 To R_COLS
            v(i, j) = out(i, j)
        Next j
    Next i
    body.Value = v
End Sub

' Save: other projects' remarks are kept; this project's are replaced by the
' non-blank Remarks cells of the grid. Returns the number saved. Does nothing
' when the grid was built before remarks existed (no Remarks column).
Private Function SaveRemarks(ByVal src As String, ByVal mainIO As String, ByRef lay As Variant, _
                             ByVal nLay As Long, ByVal who As String, ByVal stamp As Date) As Long
    Dim rd As Variant, n As Long, i As Long, j As Long, rc As Long, ws As Worksheet
    Dim out() As Variant, nOut As Long, v As Variant, ref As String
    rc = RemarkColumn()
    If rc = 0 Then Exit Function
    Set ws = modTimeline.DashSheet()
    n = ReadRemarks(rd)
    ReDim out(1 To n + nLay + 1, 1 To R_COLS)
    For i = 1 To n
        If Not RemarkIsKey(rd, i, src, mainIO) Then
            nOut = nOut + 1
            For j = 1 To R_COLS
                out(nOut, j) = rd(i, j)
            Next j
        End If
    Next i
    For i = 1 To nLay
        ref = LayRef(lay, i)
        v = ws.Cells(CLng(lay(i, 1)), rc).Value2
        If ref <> "" And Not IsError(v) Then
            If Trim$(CStr(v)) <> "" Then
                nOut = nOut + 1
                out(nOut, R_SRC) = src
                out(nOut, R_MAINIO) = mainIO
                out(nOut, R_REF) = ref
                out(nOut, R_BLOCK) = lay(i, 3)
                out(nOut, R_DESC) = lay(i, 7)
                out(nOut, R_TEXT) = CStr(v)
                out(nOut, R_BY) = who
                out(nOut, R_AT) = stamp
                SaveRemarks = SaveRemarks + 1
            End If
        End If
    Next i
    WriteRemarkRows out, nOut
End Function

' Re-assign: the Unmatched row's remark goes to the target line, after any
' remark the target already has. True when a remark was moved.
Private Function MoveRemark(ByVal src As String, ByVal mainIO As String, ByVal fromRef As String, _
                            ByVal toRef As String, ByVal toBlock As String, ByVal toDesc As String) As Boolean
    Dim rd As Variant, n As Long, i As Long, j As Long, iFrom As Long, iTo As Long, iDrop As Long
    Dim out() As Variant, nOut As Long
    If fromRef = "" Or toRef = "" Or fromRef = toRef Then Exit Function
    n = ReadRemarks(rd)
    For i = 1 To n
        If RemarkIsKey(rd, i, src, mainIO) Then
            If modTimeline.SameText(rd(i, R_REF), fromRef) Then iFrom = i
            If modTimeline.SameText(rd(i, R_REF), toRef) Then iTo = i
        End If
    Next i
    If iFrom = 0 Then Exit Function
    If iTo > 0 Then
        rd(iTo, R_TEXT) = CStr(rd(iTo, R_TEXT)) & " | " & CStr(rd(iFrom, R_TEXT))
        iDrop = iFrom
    Else
        rd(iFrom, R_REF) = toRef
        rd(iFrom, R_BLOCK) = toBlock
        rd(iFrom, R_DESC) = toDesc
        iTo = iFrom
    End If
    rd(iTo, R_BY) = Application.UserName
    rd(iTo, R_AT) = Now
    ReDim out(1 To n, 1 To R_COLS)
    For i = 1 To n
        If i <> iDrop Then
            nOut = nOut + 1
            For j = 1 To R_COLS
                out(nOut, j) = rd(i, j)
            Next j
        End If
    Next i
    WriteRemarkRows out, nOut
    SetNameText NAME_AT, Trim$(Str$(CDbl(Now)))
    MoveRemark = True
End Function

' Delete: removes the Unmatched row's remark. True when one was removed.
Private Function DeleteRemark(ByVal src As String, ByVal mainIO As String, ByVal ref As String) As Boolean
    Dim rd As Variant, n As Long, i As Long, j As Long, out() As Variant, nOut As Long
    If ref = "" Then Exit Function
    n = ReadRemarks(rd)
    If n = 0 Then Exit Function
    ReDim out(1 To n, 1 To R_COLS)
    For i = 1 To n
        If RemarkIsKey(rd, i, src, mainIO) And modTimeline.SameText(rd(i, R_REF), ref) Then
            DeleteRemark = True
        Else
            nOut = nOut + 1
            For j = 1 To R_COLS
                out(nOut, j) = rd(i, j)
            Next j
        End If
    Next i
    If DeleteRemark Then
        WriteRemarkRows out, nOut
        SetNameText NAME_AT, Trim$(Str$(CDbl(Now)))
    End If
End Function

'==============================================================================
' Grid layout (ProjectionStore!M:T) and session state
'==============================================================================
Public Sub WriteLayout(ByRef lay() As Variant, ByVal nLay As Long, ByVal r0 As Long, _
                       ByVal lastRow As Long, ByVal firstC As Long, ByVal lastC As Long, _
                       ByVal remarkC As Long)
    Dim ws As Worksheet, last As Long, v() As Variant, i As Long, j As Long
    Set ws = ThisWorkbook.Worksheets(modTimeline.STORE_SHEET)
    last = ws.Cells(ws.Rows.Count, LAYOUT_COL).End(xlUp).Row
    If last >= 2 Then ws.Range(ws.Cells(2, LAYOUT_COL), ws.Cells(last, LAYOUT_COL + LAYOUT_COLS - 1)).Clear
    If nLay > 0 Then
        ReDim v(1 To nLay, 1 To LAYOUT_COLS)
        For i = 1 To nLay
            For j = 1 To LAYOUT_COLS
                v(i, j) = lay(i, j)
            Next j
        Next i
        ws.Range(ws.Cells(2, LAYOUT_COL + 1), ws.Cells(nLay + 1, LAYOUT_COL + LAYOUT_COLS - 1)).NumberFormat = "@"
        ws.Range(ws.Cells(2, LAYOUT_COL), ws.Cells(nLay + 1, LAYOUT_COL + LAYOUT_COLS - 1)).Value = v
    End If
    SetNameText NAME_SPAN, r0 & "|" & lastRow & "|" & firstC & "|" & lastC & "|" & remarkC
End Sub

Private Function ReadLayout(ByRef lay As Variant) As Long
    Dim ws As Worksheet, last As Long
    Set ws = ThisWorkbook.Worksheets(modTimeline.STORE_SHEET)
    last = ws.Cells(ws.Rows.Count, LAYOUT_COL).End(xlUp).Row
    If last < 2 Then Exit Function
    lay = ws.Range(ws.Cells(2, LAYOUT_COL), ws.Cells(last, LAYOUT_COL + LAYOUT_COLS - 1)).Value
    ReadLayout = last - 1
End Function

' Remark key of layout row i ("" for a layout written before remarks existed)
Private Function LayRef(ByRef lay As Variant, ByVal i As Long) As String
    If UBound(lay, 2) >= 8 Then LayRef = modTimeline.CleanText(lay(i, 8))
End Function

' Remarks column on Main_Dashboard (0 for a grid built before remarks existed)
Private Function RemarkColumn() As Long
    Dim parts As Variant
    parts = Split(NameText(NAME_SPAN), "|")
    If UBound(parts) >= 4 Then RemarkColumn = Val(parts(4))
End Function

Private Function LayoutIndexOfRow(ByRef lay As Variant, ByVal nLay As Long, ByVal r As Long) As Long
    Dim i As Long
    For i = 1 To nLay
        If CLng(lay(i, 1)) = r Then
            LayoutIndexOfRow = i
            Exit Function
        End If
    Next i
End Function

Private Function CountInputs(ByRef lay As Variant, ByVal nLay As Long, ByVal blk As String) As Long
    Dim i As Long
    For i = 1 To nLay
        If lay(i, 2) = "INPUT" And lay(i, 3) = blk Then CountInputs = CountInputs + 1
    Next i
End Function

Private Function InputIdentities(ByRef lay As Variant, ByVal nLay As Long) As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To nLay
        If lay(i, 2) = "INPUT" Then
            d(modTimeline.Identity(CStr(lay(i, 3)), CStr(lay(i, 5)), CStr(lay(i, 4)), CStr(lay(i, 6)))) = i
        End If
    Next i
    Set InputIdentities = d
End Function

Private Function MonthSetOf(ByRef mDates() As Date, ByVal nM As Long) As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To nM
        d(CLng(mDates(i))) = i
    Next i
    Set MonthSetOf = d
End Function

' Month columns of the grid: header row 2 holds the month dates.
Private Function GridMonths(ByRef mCols() As Long, ByRef mDates() As Date) As Long
    Dim parts As Variant, ws As Worksheet, r As Long, c As Long, n As Long, v As Variant, d As Date
    parts = Split(NameText(NAME_SPAN), "|")
    If UBound(parts) < 3 Then Exit Function
    Set ws = modTimeline.DashSheet()
    r = CLng(parts(0)) + 1
    ReDim mCols(1 To CLng(parts(3)) - CLng(parts(2)) + 1)
    ReDim mDates(1 To CLng(parts(3)) - CLng(parts(2)) + 1)
    For c = CLng(parts(2)) To CLng(parts(3))
        v = ws.Cells(r, c).Value2
        If VarType(v) = vbDouble Then
            If modTimeline.ToDate(v, d) Then
                n = n + 1
                mCols(n) = c
                mDates(n) = MonthStart(d)
            End If
        End If
    Next c
    GridMonths = n
End Function

Private Function StampOf(ByVal v As Variant) As Double
    Dim d As Date
    If modTimeline.ToDate(v, d) Then StampOf = CDbl(d)
End Function

Private Function MonthStart(ByVal d As Date) As Date
    MonthStart = DateSerial(Year(d), Month(d), 1)
End Function

Public Sub SetSession(ByVal src As String, ByVal mainIO As String)
    SetNameText NAME_PROJECT, src
    SetNameText NAME_MAINIO, mainIO
    SetNameText NAME_AT, Trim$(Str$(CDbl(Now)))
    If src = "" Then
        SetNameText NAME_HASH, ""
    Else
        SetNameText NAME_HASH, HashInputs()
    End If
End Sub

' Makes IsDirty True until the next save or load.
Public Sub MarkDirty()
    SetNameText NAME_HASH, "changed"
End Sub

Public Function LoadedProject() As String
    LoadedProject = NameText(NAME_PROJECT)
End Function

Public Function LoadedMainIO() As String
    LoadedMainIO = NameText(NAME_MAINIO)
End Function

Public Function SelectionMatchesLoaded() As Boolean
    Dim ws As Worksheet
    Set ws = modTimeline.DashSheet()
    SelectionMatchesLoaded = modTimeline.SameText(ws.Range(modTimeline.SEL_PROJECT).Value, LoadedProject()) And _
                             modTimeline.SameText(ws.Range(modTimeline.SEL_MAINIO).Value, LoadedMainIO())
End Function

' True when the input cells differ from what was loaded or last saved.
Public Function IsDirty() As Boolean
    If LoadedProject() = "" Then Exit Function
    IsDirty = (HashInputs() <> NameText(NAME_HASH))
End Function

Private Function HashInputs() As String
    Dim lay As Variant, nLay As Long, mCols() As Long, mDates() As Date, nM As Long
    Dim ws As Worksheet, i As Long, c As Long, v As Variant, s As String, rc As Long
    nLay = ReadLayout(lay)
    nM = GridMonths(mCols, mDates)
    Set ws = modTimeline.DashSheet()
    For i = 1 To nLay
        If lay(i, 2) = "INPUT" Then
            For c = 1 To nM
                v = ws.Cells(CLng(lay(i, 1)), mCols(c)).Value2
                If Not IsEmpty(v) Then
                    If IsError(v) Then
                        s = s & lay(i, 1) & "," & mCols(c) & "=#;"
                    ElseIf VarType(v) = vbString Then
                        s = s & lay(i, 1) & "," & mCols(c) & "=" & v & ";"
                    Else
                        s = s & lay(i, 1) & "," & mCols(c) & "=" & Trim$(Str$(v)) & ";"
                    End If
                End If
            Next c
        End If
    Next i
    rc = RemarkColumn()
    If rc > 0 Then
        For i = 1 To nLay
            v = ws.Cells(CLng(lay(i, 1)), rc).Value2
            If IsError(v) Then
                s = s & "R" & lay(i, 1) & "=#;"
            ElseIf Not IsEmpty(v) Then
                s = s & "R" & lay(i, 1) & "=" & CStr(v) & ";"
            End If
        Next i
    End If
    HashInputs = HashText(s)
End Function

Private Function HashText(ByVal s As String) As String
    Dim i As Long, ch As Long, h1 As Double, h2 As Double
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1)) And &HFFFF&
        h1 = h1 * 31 + ch
        h1 = h1 - Int(h1 / 2147483647#) * 2147483647#
        h2 = h2 * 131 + ch
        h2 = h2 - Int(h2 / 1000000007#) * 1000000007#
    Next i
    HashText = Len(s) & "-" & Format$(h1, "0") & "-" & Format$(h2, "0")
End Function

Private Sub SetNameText(ByVal nm As String, ByVal txt As String)
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="=""" & Replace(txt, """", """""") & """", Visible:=False
End Sub

Private Function NameText(ByVal nm As String) As String
    Dim rt As String
    On Error Resume Next
    rt = ThisWorkbook.Names(nm).RefersTo
    On Error GoTo 0
    If Len(rt) >= 3 And Left$(rt, 2) = "=""" And Right$(rt, 1) = """" Then
        NameText = Replace(Mid$(rt, 3, Len(rt) - 3), """""", """")
    End If
End Function
