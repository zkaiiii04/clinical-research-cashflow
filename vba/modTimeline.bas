Attribute VB_Name = "modTimeline"
'==============================================================================
' modTimeline - Load Project and the macro-built projection timeline
'
' Spec: specs/2026-10-11-cashflow-dynamic-timeline-spec.md (v1.0, approved)
' Workbook: Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.9_working.xlsm
'
' Power Query safety (spec 10.5): the only Power Query action in this module
' is ListObject.QueryTable.Refresh BackgroundQuery:=False on the five
' selected-test tables (RefreshSelectedTables). Nothing here adds, edits or
' deletes queries, connections or loaded tables. There are no
' Worksheet_Change / AfterRefresh / SheetCalculate handlers.
'
' Needs desktop Excel for Windows (uses Scripting.Dictionary).
'==============================================================================
Option Explicit

Public Const DASH_SHEET As String = "Main_Dashboard"
Public Const STORE_SHEET As String = "ProjectionStore"
' The manpower selected-test table keeps its v0.8 name (incident 2026-10-11:
' renaming it outside Excel broke its query link). If it is ever renamed in
' the Excel UI, change this one constant.
Public Const MANPOWER_TABLE As String = "qManpower_SelectedTest"
' Main_Dashboard cells the macros read and write. Change them here if the
' selector or notice is moved on the sheet (the user moved the selector from
' row 6 to row 5 on 12 Oct 2026).
Public Const SEL_PROJECT As String = "E5"     ' Source Project dropdown
Public Const SEL_MAINIO As String = "F5"      ' Main IO Code dropdown
Public Const SEL_STATUS As String = "G5"      ' Setup Status formula
Public Const NOTICE_CELL As String = "E7"     ' "Dashboard loaded by ..." notice
Public Const BASIS_CELL As String = "C6"      ' Calendar Basis (Financial Year / Calendar Year)
Public Const REFRESHED_CELL As String = "C10" ' Last Refreshed
' Rows above the grid (selector, KPI cards, buttons) that each build groups so
' they collapse with the +/- beside row 1 (user request 12 Oct 2026).
Public Const TOP_FIRST_ROW As Long = 2
Public Const TOP_LAST_ROW As Long = 19

' Grid columns are offsets from TL_Anchor (B21). The Line label sits in B and
' overflows across C:E, so IO-Code is in F, Budget in G and the collapsible
' month columns start at H. Collapsing a quarter therefore never hides the
' selector (SEL_* cells) or the KPI cards above the grid.
Private Const OFF_IO As Long = 4
Private Const OFF_BUDGET As Long = 5
Private Const OFF_FIRST As Long = 6

Private Const FMT_MONEY As String = "#,##0;[Red](#,##0)"
Private Const FMT_PCT As String = "0.0%;[Red](0.0%)"
Private Const MAX_MONTHS As Long = 240
Private Const REMARK_WIDTH As Double = 40   ' Remarks column after Project Total (wraps)

Private Type TLine
    Kind As String          ' INPUT, BUDGET_MP, HIDDEN, MANPOWER, RNF, UNMATCHED, MP_OOR
    Block As String         ' MP, OM, EQ, CO, OTHER
    IO As String            ' upper-case IO
    Code As String          ' budget code (upper case); MP for manpower
    LineKey As String       ' normalised description (spec 5.4); rate category for manpower
    Desc As String          ' description text shown and stored
    Label As String         ' text in the Line column
    IOCode As String        ' text in the IO-Code column
    Budget As Double
    HasBudget As Boolean
    Seq As Long
    SortKey As String
    Vals() As Variant       ' 1..nMonths: Empty or Double
    OORAmount As Double     ' amounts in months outside the project dates (Project Total only)
    OORNote As String
    RnfRows As Long
    RnfFTE As Double
    Row As Long
    RemarkRef As String     ' key of the line's saved remark (modStore remarks area)
    Remark As String
End Type

Private L() As TLine
Private nL As Long
Private ord() As Long
Private nOrd As Long
Private lineIndex As Object         ' Block|Code|IO|LineKey -> line index (budget lines)
Private mStart As Date
Private nMonths As Long
Private fyBasis As Boolean
Private soaMonth As Date            ' Latest SOA Month; 0 when the project has no actuals
Private nCols As Long, nQ As Long, nY As Long, pCol As Long
Private colKind() As String, colMonth() As Long, colQ() As Long, colY() As Long
Private qFirstC() As Long, qLastC() As Long, qTotC() As Long, qNum() As Long, qY() As Long
Private qLbl() As String
Private yLbl() As String, yCol() As Long
Private blockAct() As Double        ' (1..5, 1..nMonths)
Private blockHasAct() As Boolean    ' (1..5)
Private yExtra() As Double          ' (1..5, 1..nY): actuals outside the dates, inside an FY/CY column
Private pExtra() As Double          ' (1..5): actuals outside the dates and every FY/CY column
Private actNote() As String         ' (1..5)
Private otherCodes As Object
Private warnRNF As Long, warnAdmin As Long, warnUnmatched As Long, warnCharged As Long
Private warnChargeBad As String
Private dateNote As String          ' set by CheckTables: SOA LOA dates differ from the project dates
Private storeRowsForKey As Long
Private remarkMoved As Boolean      ' a saved remark is shown under a different key (Save re-keys it)

Private mCalc As Long
Private mScreen As Boolean
Private mEvents As Boolean

'==============================================================================
' Button: Load Project (spec 10.1)
'==============================================================================
Public Sub LoadProject()
    Dim ws As Worksheet, v As Variant, status As String, msg As String
    Dim ans As VbMsgBoxResult
    On Error GoTo Fail
    Set ws = DashSheet()

    ' 1. Unsaved edits on the project that is loaded now
    If modStore.IsDirty() Then
        ans = MsgBox("Save changes to " & modStore.LoadedProject() & " first?", _
                     vbYesNoCancel + vbQuestion, "Load Project")
        If ans = vbCancel Then Exit Sub
        If ans = vbYes Then
            If Not modStore.SaveCore(True) Then Exit Sub
        End If
    End If

    ' 2. Setup Status must be Ready
    v = ws.Range(SEL_STATUS).Value
    If IsError(v) Then status = "(error)" Else status = CleanText(v)
    If status <> "Ready" Then
        MsgBox "Setup Status is """ & status & """." & vbCrLf & vbCrLf & _
               "Choose a Source Project in " & SEL_PROJECT & " and its Main IO in " & SEL_MAINIO & ", then try again.", _
               vbExclamation, "Load Project"
        Exit Sub
    End If

    ' 3. Refresh only the five selected-test tables, one at a time
    If Not RefreshSelectedTables(msg) Then
        Application.StatusBar = False
        MsgBox msg & vbCrLf & vbCrLf & "The timeline was not changed.", vbExclamation, "Load Project"
        Exit Sub
    End If

    ' 4-8. Check, build, load saved projections, notice, warnings
    BuildGrid True
    Exit Sub
Fail:
    Application.StatusBar = False
    MsgBox "Load Project stopped: " & Err.Description, vbExclamation, "Load Project"
End Sub

' Refreshes tblSelectedProjectMapping, tblBudgetSelectedTest,
' tblActualsSelectedTest, qProjects_SelectedTest and qManpower_SelectedTest,
' in that order, synchronously. Each connection's background-refresh setting is
' saved first and put back afterwards. qProjectIOCatalogue is not refreshed.
Private Function RefreshSelectedTables(ByRef msg As String) As Boolean
    Dim tblNames As Variant, i As Long, lo As ListObject, cn As Object
    Dim hadBg As Boolean, bg As Boolean, errNo As Long, errDesc As String

    tblNames = Array("tblSelectedProjectMapping", "tblBudgetSelectedTest", _
                  "tblActualsSelectedTest", "qProjects_SelectedTest", MANPOWER_TABLE)
    For i = LBound(tblNames) To UBound(tblNames)
        Set lo = FindTable(CStr(tblNames(i)))
        If lo Is Nothing Then
            msg = "Table " & tblNames(i) & " was not found."
            Exit Function
        End If
        Application.StatusBar = "Load Project: refreshing " & tblNames(i) & " (" & (i + 1) & " of 5)..."

        hadBg = False
        Set cn = Nothing
        On Error Resume Next
        Set cn = lo.QueryTable.WorkbookConnection
        If Not cn Is Nothing Then
            bg = cn.OLEDBConnection.BackgroundQuery
            If Err.Number = 0 Then hadBg = True
        End If
        Err.Clear
        If hadBg Then cn.OLEDBConnection.BackgroundQuery = False
        Err.Clear
        lo.QueryTable.Refresh BackgroundQuery:=False
        errNo = Err.Number
        errDesc = Err.Description
        Err.Clear
        If hadBg Then cn.OLEDBConnection.BackgroundQuery = bg
        On Error GoTo 0

        If errNo <> 0 Then
            msg = "Refreshing " & tblNames(i) & " failed: " & errDesc
            Exit Function
        End If
    Next i
    Application.StatusBar = False
    RefreshSelectedTables = True
End Function

'==============================================================================
' Build the grid from the loaded tables and the store (spec 5, 6, 9).
' isLoad = True: called by Load Project (shows warnings).
' isLoad = False: rebuild after Re-assign / Delete (no refresh, no warnings).
'==============================================================================
Public Function BuildGrid(ByVal isLoad As Boolean) As Boolean
    Dim ws As Worksheet, src As String, mainIO As String
    Dim dispName As String, srcName As String, pStart As Date, pEnd As Date
    Dim msg As String, started As Boolean
    On Error GoTo Fail
    Set ws = DashSheet()
    src = CleanText(ws.Range(SEL_PROJECT).Value)
    mainIO = CleanText(ws.Range(SEL_MAINIO).Value)

    ' 4. Consistency check, before anything on the sheet changes
    If Not CheckTables(src, mainIO, dispName, srcName, pStart, pEnd, msg) Then
        Application.StatusBar = False
        MsgBox msg & vbCrLf & vbCrLf & "The timeline was not changed. Click Load Project again; " & _
               "if this repeats, use Data > Refresh All and check the selected-test queries.", _
               vbExclamation, "Load Project"
        Exit Function
    End If

    BeginWork
    started = True
    Application.StatusBar = "Load Project: building the timeline..."
    ResetModel
    fyBasis = (CleanText(ws.Range(BASIS_CELL).Value) <> "Calendar Year")
    BuildColumns pStart, pEnd
    ReadBudget
    ReadManpower src, mainIO
    ReadActuals
    ReadStoreRows src, mainIO
    ReadRemarkRows src, mainIO
    SortLines

    ' 5-7. Grid, saved projections, notice, session state
    WriteGrid ws, src, mainIO, dispName
    If isLoad Then
        ws.Range(NOTICE_CELL).Value = "Dashboard loaded by " & Application.UserName & " at " & _
                              Format$(Now, "d mmm yyyy h:mm AM/PM")
        ws.Range(NOTICE_CELL).Font.Italic = True
        ws.Range(REFRESHED_CELL).Value = Now
        ws.Range(REFRESHED_CELL).NumberFormat = "dd mmm yyyy hh:mm"
    End If
    modStore.SetSession src, mainIO
    ' A remark that moved to another row counts as an unsaved change, so the
    ' next Save (or the save offered by Re-assign / Delete) stores it there
    If remarkMoved Then modStore.MarkDirty
    ProtectDashboard
    EndWork

    ' 8. Warnings
    If isLoad Then ShowWarnings dispName
    BuildGrid = True
    Exit Function
Fail:
    msg = Err.Description
    If started Then CleanupAfterFail ws
    Application.StatusBar = False
    MsgBox "Building the timeline stopped: " & msg, vbExclamation, "Load Project"
End Function

' Runs outside the error handler so its own On Error Resume Next works and
' Excel's calculation / screen settings are always restored.
Private Sub CleanupAfterFail(ByVal ws As Worksheet)
    On Error Resume Next
    ResetKpiNames
    modStore.SetSession "", ""
    ws.Range(NOTICE_CELL).Value = "Timeline build failed. Click Load Project again."
    ProtectDashboard
    EndWork
End Sub

Private Sub ShowWarnings(ByVal dispName As String)
    Dim w As String
    If dateNote <> "" Then w = w & "- " & dateNote & vbCrLf
    If warnRNF > 0 Then w = w & "- " & warnRNF & " manpower line(s) have Rate Status ""Rate not found"". " & _
        "They are shown with no amount; check the rate grade in tbl_ManpowerRate." & vbCrLf
    If warnAdmin > 0 Then w = w & "- " & warnAdmin & " manpower line(s) use the Admin rate category, " & _
        "which isn't expected on the manpower sheets. They are shown and counted in MP." & vbCrLf
    If warnUnmatched > 0 Then w = w & "- " & warnUnmatched & " unmatched saved projection or remark line(s). " & _
        "They are counted in the totals; select one and use Re-assign or Delete." & vbCrLf
    If warnCharged > 0 Then w = w & "- " & warnCharged & " manpower line(s) are charged to another " & _
        "category (Dashboard_Settings)." & vbCrLf
    If warnChargeBad <> "" Then w = w & "- tblManpowerChargeTo has Charge To values that aren't " & _
        "MP, OM, EQ, CO or OTHER (" & warnChargeBad & "). Those lines stay in MP." & vbCrLf
    If w <> "" Then MsgBox "Loaded " & dispName & "." & vbCrLf & vbCrLf & w, vbInformation, "Load Project"
End Sub

'==============================================================================
' Consistency check (spec 10.1 step 4) and project dates (spec 10.1 step 5).
' The timeline dates come from qProjects_SelectedTest (IREx project
' information): the IO Code Type "Main" row for the Main IO. The dates in
' tblSelectedProjectMapping are the min/max of the SOA budget LOA dates; they
' are used only when the project dates are blank, and a difference is reported
' as a Load warning (incident 2026-10-11, project end date).
'==============================================================================
Private Function CheckTables(ByVal src As String, ByVal mainIO As String, _
        ByRef dispName As String, ByRef srcName As String, _
        ByRef pStart As Date, ByRef pEnd As Date, ByRef msg As String) As Boolean
    Dim lo As ListObject, v As Variant, i As Long, nRows As Long, found As Boolean
    Dim cSrc As Long, cDisp As Long, cIO As Long, cMain As Long, cStart As Long, cEnd As Long
    Dim mapStart As Date, mapEnd As Date, hasMapDates As Boolean, hasProjDates As Boolean

    dateNote = ""
    v = TableData("tblSelectedProjectMapping", lo)
    If Not IsEmpty(v) Then
        cSrc = ColIndex(lo, "Source Project Name")
        cDisp = ColIndex(lo, "Project Display Name")
        cIO = ColIndex(lo, "IO Code")
        cMain = ColIndex(lo, "Main IO")
        cStart = ColIndex(lo, "Project Start")
        cEnd = ColIndex(lo, "Project End")
        For i = 1 To UBound(v, 1)
            If CleanText(v(i, cSrc)) <> "" Then
                nRows = nRows + 1
                If Not SameText(v(i, cSrc), src) Then
                    msg = "tblSelectedProjectMapping shows """ & CleanText(v(i, cSrc)) & _
                          """, not """ & src & """."
                    Exit Function
                End If
                If Not found And SameText(v(i, cMain), "Yes") And SameText(v(i, cIO), mainIO) Then
                    found = True
                    dispName = CleanText(v(i, cDisp))
                    srcName = CleanText(v(i, cSrc))
                    hasMapDates = ToDate(v(i, cStart), mapStart) And ToDate(v(i, cEnd), mapEnd)
                End If
            End If
        Next i
    End If
    If nRows = 0 Then
        msg = "tblSelectedProjectMapping is empty."
        Exit Function
    End If
    If Not found Then
        msg = "tblSelectedProjectMapping has no Main IO = Yes row for " & mainIO & "."
        Exit Function
    End If

    hasProjDates = ProjectDates(mainIO, pStart, pEnd)
    If hasProjDates Then
        If hasMapDates Then
            If mapStart <> pStart Or mapEnd <> pEnd Then
                dateNote = "The SOA budget LOA dates (" & Format$(mapStart, "d mmm yyyy") & " - " & _
                    Format$(mapEnd, "d mmm yyyy") & ") differ from the project dates in " & _
                    "qProjects_SelectedTest (" & Format$(pStart, "d mmm yyyy") & " - " & _
                    Format$(pEnd, "d mmm yyyy") & "). The timeline uses the project dates; " & _
                    "update the LOA dates in the SOA file if they are out of date."
            End If
        End If
    ElseIf hasMapDates Then
        pStart = mapStart
        pEnd = mapEnd
        dateNote = "qProjects_SelectedTest has no Project Start/End for the Main IO " & mainIO & _
            ", so the timeline uses the SOA budget LOA dates (" & Format$(pStart, "d mmm yyyy") & _
            " - " & Format$(pEnd, "d mmm yyyy") & "). Check the project dates in IREx."
    Else
        msg = "Project Start or Project End is missing for " & src & _
              " in qProjects_SelectedTest and tblSelectedProjectMapping."
        Exit Function
    End If
    If pEnd < pStart Then
        msg = "Project End is before Project Start for " & src & "."
        Exit Function
    End If
    If DateDiff("m", pStart, pEnd) + 1 > MAX_MONTHS Then
        msg = "The project runs for more than " & MAX_MONTHS & " months; check its dates."
        Exit Function
    End If

    ' Budget, actuals and projects carry the Project Display Name;
    ' manpower carries the Source Project Name.
    If Not CheckGroup("tblBudgetSelectedTest", dispName, msg) Then Exit Function
    If Not CheckGroup("tblActualsSelectedTest", dispName, msg) Then Exit Function
    If Not CheckGroup("qProjects_SelectedTest", dispName, msg) Then Exit Function
    If Not CheckGroup(MANPOWER_TABLE, srcName, msg) Then Exit Function
    CheckTables = True
End Function

' Project Start/End from qProjects_SelectedTest: the IO Code Type "Main" row
' whose IO Code is the Main IO, otherwise the first IO Code Type "Main" row.
' Returns False when no such row has both dates.
Private Function ProjectDates(ByVal mainIO As String, ByRef pStart As Date, ByRef pEnd As Date) As Boolean
    Dim lo As ListObject, v As Variant, i As Long, pass As Long
    Dim cIO As Long, cType As Long, cStart As Long, cEnd As Long
    v = TableData("qProjects_SelectedTest", lo)
    If IsEmpty(v) Then Exit Function
    cIO = ColIndex(lo, "IO Code")
    cType = ColIndex(lo, "IO Code Type")
    cStart = ColIndex(lo, "Project Start")
    cEnd = ColIndex(lo, "Project End")
    For pass = 1 To 2
        For i = 1 To UBound(v, 1)
            If SameText(v(i, cType), "Main") And (pass = 2 Or SameText(v(i, cIO), mainIO)) Then
                If ToDate(v(i, cStart), pStart) And ToDate(v(i, cEnd), pEnd) Then
                    ProjectDates = True
                    Exit Function
                End If
            End If
        Next i
    Next pass
End Function

Private Function CheckGroup(ByVal tableName As String, ByVal expected As String, _
                            ByRef msg As String) As Boolean
    Dim lo As ListObject, v As Variant, i As Long, c As Long, pg As String
    v = TableData(tableName, lo)
    If Not IsEmpty(v) Then
        c = ColIndex(lo, "Project Group")
        For i = 1 To UBound(v, 1)
            pg = CleanText(v(i, c))
            If pg <> "" And Not SameText(pg, expected) Then
                msg = tableName & " shows """ & pg & """, not """ & expected & """."
                Exit Function
            End If
        Next i
    End If
    CheckGroup = True
End Function

'==============================================================================
' Model
'==============================================================================
Private Sub ResetModel()
    nL = 0
    Erase L
    nOrd = 0
    Erase ord
    Set lineIndex = CreateObject("Scripting.Dictionary")
    Set otherCodes = CreateObject("Scripting.Dictionary")
    ReDim blockHasAct(1 To 5)
    ReDim pExtra(1 To 5)
    ReDim actNote(1 To 5)
    soaMonth = 0
    warnRNF = 0: warnAdmin = 0: warnUnmatched = 0: warnCharged = 0
    warnChargeBad = ""
    storeRowsForKey = 0
End Sub

Private Function MonthDate(ByVal i As Long) As Date
    MonthDate = DateSerial(Year(mStart), Month(mStart) + i - 1, 1)
End Function

Private Function MonthIndex(ByVal d As Date) As Long
    MonthIndex = DateDiff("m", mStart, DateSerial(Year(d), Month(d), 1)) + 1
End Function

Private Function InRange(ByVal mi As Long) As Boolean
    InRange = (mi >= 1 And mi <= nMonths)
End Function

' FY = April to March, named by its starting year (FY25/Q3); CY = 2026 | Q1.
Private Sub PeriodOf(ByVal d As Date, ByRef qn As Long, ByRef qLabel As String, ByRef yLabel As String)
    Dim yk As Long
    If fyBasis Then
        yk = Year(d)
        If Month(d) < 4 Then yk = yk - 1
        qn = ((Month(d) + 8) Mod 12) \ 3 + 1
        qLabel = "FY" & Right$(CStr(yk), 2) & "/Q" & qn
        yLabel = "FY" & Right$(CStr(yk), 2) & " Total"
    Else
        yk = Year(d)
        qn = (Month(d) - 1) \ 3 + 1
        qLabel = yk & " | Q" & qn
        yLabel = yk & " Total"
    End If
End Sub

Private Sub BuildColumns(ByVal pStart As Date, ByVal pEnd As Date)
    Dim i As Long, k As Long, y As Long, qn As Long, ql As String, yl As String
    Dim nextQn As Long, nextQl As String, nextYl As String, prevQl As String

    mStart = DateSerial(Year(pStart), Month(pStart), 1)
    nMonths = DateDiff("m", mStart, DateSerial(Year(pEnd), Month(pEnd), 1)) + 1
    ReDim colKind(1 To nMonths * 2 + 50)
    ReDim colMonth(1 To nMonths * 2 + 50)
    ReDim colQ(1 To nMonths * 2 + 50)
    ReDim colY(1 To nMonths * 2 + 50)
    ReDim qFirstC(1 To nMonths): ReDim qLastC(1 To nMonths): ReDim qTotC(1 To nMonths)
    ReDim qNum(1 To nMonths): ReDim qY(1 To nMonths): ReDim qLbl(1 To nMonths)
    ReDim yLbl(1 To nMonths): ReDim yCol(1 To nMonths)
    nCols = 0: nQ = 0: nY = 0: prevQl = ""

    For i = 1 To nMonths
        PeriodOf MonthDate(i), qn, ql, yl
        If ql <> prevQl Then
            nQ = nQ + 1
            qLbl(nQ) = ql
            qNum(nQ) = qn
            qFirstC(nQ) = nCols + 1
            y = 0
            For k = 1 To nY
                If yLbl(k) = yl Then y = k
            Next k
            If y = 0 Then
                nY = nY + 1
                yLbl(nY) = yl
                y = nY
            End If
            qY(nQ) = y
            prevQl = ql
        End If
        nCols = nCols + 1
        colKind(nCols) = "M": colMonth(nCols) = i: colQ(nCols) = nQ
        If i < nMonths Then
            PeriodOf MonthDate(i + 1), nextQn, nextQl, nextYl
        Else
            nextQl = ""
        End If
        If nextQl <> ql Then
            qLastC(nQ) = nCols
            nCols = nCols + 1
            colKind(nCols) = "Q": colQ(nCols) = nQ
            qTotC(nQ) = nCols
        End If
    Next i
    For k = 1 To nY
        nCols = nCols + 1
        colKind(nCols) = "Y": colY(nCols) = k
        yCol(k) = nCols
    Next k
    nCols = nCols + 1
    colKind(nCols) = "P"
    pCol = nCols

    ReDim blockAct(1 To 5, 1 To nMonths)
    ReDim yExtra(1 To 5, 1 To nY)
End Sub

Private Function AddLine(ByVal kind As String, ByVal blk As String, ByVal io As String, _
                         ByVal code As String, ByVal lk As String, ByVal desc As String) As Long
    nL = nL + 1
    ReDim Preserve L(1 To nL)
    With L(nL)
        .Kind = kind
        .Block = blk
        .IO = io
        .Code = code
        .LineKey = lk
        .Desc = desc
        .Label = desc
        .Seq = nL
        .RemarkRef = Identity(blk, code, io, lk)
    End With
    ReDim L(nL).Vals(1 To nMonths)
    AddLine = nL
End Function

' Budget lines: one per Block + Code + IO + normalised description; net budget.
Private Sub ReadBudget()
    Dim lo As ListObject, v As Variant, i As Long, idx As Long
    Dim cIO As Long, cCat As Long, cDesc As Long, cAmt As Long
    Dim io As String, code As String, blk As String, desc As String, lk As String, key As String

    v = TableData("tblBudgetSelectedTest", lo)
    If IsEmpty(v) Then Exit Sub
    cIO = ColIndex(lo, "IO Code")
    cCat = ColIndex(lo, "Expense Category")
    cDesc = ColIndex(lo, "Budget Description")
    cAmt = ColIndex(lo, "Revised Budget ($)")
    For i = 1 To UBound(v, 1)
        io = UCase$(CleanText(v(i, cIO)))
        If io <> "" Then
            code = UCase$(CleanText(v(i, cCat)))
            blk = BlockOf(code)
            desc = CleanText(v(i, cDesc))
            lk = NormKey(desc)
            key = Identity(blk, code, io, lk)
            If lineIndex.Exists(key) Then
                idx = lineIndex(key)
            Else
                If blk = "MP" Then
                    idx = AddLine("BUDGET_MP", blk, io, code, lk, desc)
                Else
                    idx = AddLine("INPUT", blk, io, code, lk, desc)
                End If
                L(idx).IOCode = IOCodeText(io, code)
                L(idx).HasBudget = True
                lineIndex.Add key, idx
            End If
            L(idx).Budget = L(idx).Budget + ToNumber(v(i, cAmt))
        End If
    Next i
    ' Lines that net to $0 are not shown (spec 5.3)
    For idx = 1 To nL
        If L(idx).HasBudget And Abs(L(idx).Budget) < 0.005 Then L(idx).Kind = "HIDDEN"
    Next idx
End Sub

' tblManpowerChargeTo rows for this project: "RATECATEGORY|IO" -> block
Private Function ReadChargeTo(ByVal src As String, ByVal mainIO As String) As Object
    Dim d As Object, lo As ListObject, v As Variant, i As Long
    Dim cSrc As Long, cMain As Long, cCat As Long, cIO As Long, cTo As Long, target As String
    Set d = CreateObject("Scripting.Dictionary")
    Set ReadChargeTo = d
    v = TableData("tblManpowerChargeTo", lo)
    If IsEmpty(v) Then Exit Function
    cSrc = ColIndex(lo, "Source Project")
    cMain = ColIndex(lo, "Main IO")
    cCat = ColIndex(lo, "Rate Category")
    cIO = ColIndex(lo, "IO Code")
    cTo = ColIndex(lo, "Charge To")
    For i = 1 To UBound(v, 1)
        If SameText(v(i, cSrc), src) And SameText(v(i, cMain), mainIO) Then
            target = UCase$(CleanText(v(i, cTo)))
            Select Case target
                Case "MP", "OM", "EQ", "CO", "OTHER"
                    d(UCase$(CleanText(v(i, cCat))) & "|" & UCase$(CleanText(v(i, cIO)))) = target
                Case Else
                    If InStr(warnChargeBad, """" & target & """") = 0 Then
                        If warnChargeBad <> "" Then warnChargeBad = warnChargeBad & ", "
                        warnChargeBad = warnChargeBad & """" & target & """"
                    End If
            End Select
        End If
    Next i
End Function

' Manpower lines: one per Rate Category + IO (read-only, FTE x rate from Power
' Query). "Rate not found" rows go on a <Staff Group> row with no amount.
Private Sub ReadManpower(ByVal src As String, ByVal mainIO As String)
    Dim lo As ListObject, v As Variant, i As Long, idx As Long, mi As Long, d As Date
    Dim cIO As Long, cSG As Long, cMonth As Long, cFTE As Long, cCat As Long, cAmt As Long, cStat As Long
    Dim chargeTo As Object, mpIndex As Object
    Dim io As String, sg As String, cat As String, disp As String, key As String, blk As String
    Dim isRNF As Boolean, amt As Double

    Set chargeTo = ReadChargeTo(src, mainIO)
    Set mpIndex = CreateObject("Scripting.Dictionary")
    v = TableData(MANPOWER_TABLE, lo)
    If IsEmpty(v) Then Exit Sub
    cIO = ColIndex(lo, "IO Code")
    cSG = ColIndex(lo, "Staff Group")
    cMonth = ColIndex(lo, "Month")
    cFTE = ColIndex(lo, "Monthly FTE")
    cCat = ColIndex(lo, "Rate Category")
    cAmt = ColIndex(lo, "Projected Amount ($)")
    cStat = ColIndex(lo, "Rate Status")

    For i = 1 To UBound(v, 1)
        io = CleanText(v(i, cIO))
        If io <> "" Then
            sg = CleanText(v(i, cSG))
            cat = UCase$(CleanText(v(i, cCat)))
            isRNF = (UCase$(CleanText(v(i, cStat))) = "RATE NOT FOUND")
            blk = "MP"
            If isRNF Then
                key = "RNF|" & UCase$(sg) & "|" & UCase$(io)
            Else
                key = "MPW|" & cat & "|" & UCase$(io)
                If chargeTo.Exists(cat & "|" & UCase$(io)) Then blk = chargeTo(cat & "|" & UCase$(io))
            End If

            If mpIndex.Exists(key) Then
                idx = mpIndex(key)
            ElseIf isRNF Then
                If sg = "" Then sg = "Manpower"
                idx = AddLine("RNF", "MP", UCase$(io), "MP", "", sg)
                L(idx).Label = sg & " (Projected) " & ChrW(&H26A0) & " Rate not found"
                L(idx).IOCode = io
                L(idx).RemarkRef = key
                mpIndex.Add key, idx
                warnRNF = warnRNF + 1
            Else
                disp = DisplayCategory(cat)
                idx = AddLine("MANPOWER", blk, UCase$(io), "MP", cat, disp)
                L(idx).Label = disp & " (Projected)"
                If blk <> "MP" Then
                    L(idx).Label = L(idx).Label & "  " & ChrW(&H2699) & " manpower " & ChrW(&H2192) & " " & blk
                    warnCharged = warnCharged + 1
                End If
                If cat = "ADMIN" Then warnAdmin = warnAdmin + 1
                L(idx).IOCode = io
                L(idx).RemarkRef = key
                mpIndex.Add key, idx
            End If

            If isRNF Then
                L(idx).RnfRows = L(idx).RnfRows + 1
                L(idx).RnfFTE = L(idx).RnfFTE + ToNumber(v(i, cFTE))
            ElseIf ToDate(v(i, cMonth), d) Then
                amt = ToNumber(v(i, cAmt))
                mi = MonthIndex(d)
                If InRange(mi) Then
                    L(idx).Vals(mi) = NzD(L(idx).Vals(mi)) + amt
                Else
                    AddManpowerOutside blk, L(idx).Label, d, amt
                End If
            End If
        End If
    Next i
End Sub

' Manpower months outside the project dates: block's "Unmatched" row,
' Project Total only, listed in a note (spec 5.3, 9).
Private Sub AddManpowerOutside(ByVal blk As String, ByVal lineLabel As String, _
                               ByVal d As Date, ByVal amt As Double)
    Dim i As Long, idx As Long
    For i = 1 To nL
        If L(i).Kind = "MP_OOR" And L(i).Block = blk Then idx = i
    Next i
    If idx = 0 Then
        idx = AddLine("MP_OOR", blk, "", "MP", "", "manpower months outside project dates")
        L(idx).Label = "Unmatched: manpower months outside project dates"
        L(idx).RemarkRef = "OOR|" & blk
    End If
    L(idx).OORAmount = L(idx).OORAmount + amt
    L(idx).OORNote = L(idx).OORNote & vbLf & lineLabel & ", " & Format$(d, "mmm yy") & ": " & FmtMoney(amt)
End Sub

' Actuals by category and Posting Month (spec 6). Actuals outside the project
' dates count in the FY/CY total when that column exists, and always in the
' Project Total, with a note (spec 14.1).
Private Sub ReadActuals()
    Dim lo As ListObject, v As Variant, i As Long, b As Long, mi As Long, y As Long, k As Long
    Dim cCat As Long, cPM As Long, cPD As Long, cAmt As Long
    Dim code As String, d As Date, amt As Double, qn As Long, ql As String, yl As String, hasDate As Boolean

    v = TableData("tblActualsSelectedTest", lo)
    If IsEmpty(v) Then Exit Sub
    cCat = ColIndex(lo, "Expense Category")
    cPM = ColIndex(lo, "Posting Month")
    On Error Resume Next
    cPD = lo.ListColumns("Posting Date").Index
    On Error GoTo 0
    cAmt = ColIndex(lo, "Actual Amount ($)")
    For i = 1 To UBound(v, 1)
        hasDate = ToDate(v(i, cPM), d)
        If Not hasDate And cPD > 0 Then hasDate = ToDate(v(i, cPD), d)
        If hasDate Or Not IsBlank(v(i, cAmt)) Then
            code = UCase$(CleanText(v(i, cCat)))
            b = BlockIndex(BlockOf(code))
            If b = 5 Then otherCodes(CodeText(code)) = True
            amt = ToNumber(v(i, cAmt))
            blockHasAct(b) = True
            If hasDate Then
                d = DateSerial(Year(d), Month(d), 1)
                If d > soaMonth Then soaMonth = d
                mi = MonthIndex(d)
                If InRange(mi) Then
                    blockAct(b, mi) = blockAct(b, mi) + amt
                Else
                    PeriodOf d, qn, ql, yl
                    y = 0
                    For k = 1 To nY
                        If yLbl(k) = yl Then y = k
                    Next k
                    If y > 0 Then
                        yExtra(b, y) = yExtra(b, y) + amt
                        actNote(b) = actNote(b) & vbLf & Format$(d, "mmm yy") & ": " & FmtMoney(amt) & " (also in " & yl & ")"
                    Else
                        pExtra(b) = pExtra(b) + amt
                        actNote(b) = actNote(b) & vbLf & Format$(d, "mmm yy") & ": " & FmtMoney(amt)
                    End If
                End If
            Else
                pExtra(b) = pExtra(b) + amt
                actNote(b) = actNote(b) & vbLf & "(no posting month): " & FmtMoney(amt)
            End If
        End If
    Next i
End Sub

' Saved projections for this project (spec 8, 9). Rows that match a shown
' input line and a month column fill the input cells; the rest become
' Unmatched rows (one per saved line), counted in the totals.
Private Sub ReadStoreRows(ByVal src As String, ByVal mainIO As String)
    Dim data As Variant, n As Long, i As Long, idx As Long, uIdx As Long, mi As Long
    Dim blk As String, code As String, io As String, lk As String, key As String, reason As String
    Dim d As Date, amt As Double, orphans As Object

    Set orphans = CreateObject("Scripting.Dictionary")
    n = modStore.ReadStoreData(data)
    For i = 1 To n
        If modStore.RowIsKey(data, i, src, mainIO) Then
            If ToDate(data(i, modStore.S_MONTH), d) And IsNumeric(data(i, modStore.S_AMT)) _
               And Not IsBlank(data(i, modStore.S_AMT)) Then
                storeRowsForKey = storeRowsForKey + 1
                key = modStore.StoreIdentity(data, i)
                blk = Split(key, "|", 4)(0)
                code = Split(key, "|", 4)(1)
                io = Split(key, "|", 4)(2)
                lk = Split(key, "|", 4)(3)
                amt = CDbl(data(i, modStore.S_AMT))
                mi = MonthIndex(d)
                idx = 0
                If lineIndex.Exists(key) Then idx = lineIndex(key)
                If idx > 0 And InRange(mi) Then
                    If L(idx).Kind <> "INPUT" Then idx = 0
                Else
                    idx = 0
                End If
                If idx > 0 Then
                    L(idx).Vals(mi) = NzD(L(idx).Vals(mi)) + amt
                Else
                    If orphans.Exists(key) Then
                        uIdx = orphans(key)
                    Else
                        If Not lineIndex.Exists(key) Then
                            reason = "no longer in budget"
                        ElseIf L(lineIndex(key)).Kind = "HIDDEN" Then
                            reason = "line now nets to $0"
                        ElseIf L(lineIndex(key)).Kind = "INPUT" Then
                            reason = "month outside project dates"
                        Else
                            reason = "MP lines are not typed"
                        End If
                        uIdx = AddLine("UNMATCHED", blk, io, code, lk, _
                                       CleanText(data(i, modStore.S_DESC)))
                        L(uIdx).Label = "Unmatched: " & L(uIdx).Desc & " (" & reason & ")"
                        L(uIdx).IOCode = IOCodeText(io, code)
                        ' Own remark key: the line itself may still be shown
                        L(uIdx).RemarkRef = "UNM|" & key
                        orphans.Add key, uIdx
                        warnUnmatched = warnUnmatched + 1
                    End If
                    If InRange(mi) Then
                        L(uIdx).Vals(mi) = NzD(L(uIdx).Vals(mi)) + amt
                    Else
                        L(uIdx).OORAmount = L(uIdx).OORAmount + amt
                        L(uIdx).OORNote = L(uIdx).OORNote & vbLf & Format$(d, "mmm yy") & ": " & FmtMoney(amt)
                    End If
                End If
            End If
        End If
    Next i
End Sub

' Saved remarks for this project (spec 5.5). A remark shows on the line with
' its key. If that line is gone, it follows the line to its Unmatched row,
' or gets an Unmatched row of its own ("remark only"), so every saved remark
' is on the grid and Save can replace the project's remarks as a whole.
Private Sub ReadRemarkRows(ByVal src As String, ByVal mainIO As String)
    Dim rd As Variant, n As Long, i As Long, idx As Long, refIndex As Object
    Dim ref As String, stored As String, txt As String, key As String, parts As Variant
    Dim blk As String, code As String, io As String, lk As String, desc As String, reason As String

    remarkMoved = False
    n = modStore.ReadRemarks(rd)
    If n = 0 Then Exit Sub
    Set refIndex = CreateObject("Scripting.Dictionary")
    For i = 1 To nL
        If L(i).Kind <> "HIDDEN" And Not refIndex.Exists(L(i).RemarkRef) Then refIndex.Add L(i).RemarkRef, i
    Next i
    For i = 1 To n
        If modStore.RemarkIsKey(rd, i, src, mainIO) Then
            ref = CleanText(rd(i, modStore.R_REF))
            stored = ref
            txt = ""
            If Not IsError(rd(i, modStore.R_TEXT)) Then txt = CStr(rd(i, modStore.R_TEXT))
            If ref <> "" And Trim$(txt) <> "" Then
                idx = 0
                If refIndex.Exists(ref) Then
                    idx = refIndex(ref)
                ElseIf Left$(ref, 4) = "UNM|" Then
                    ' Its Unmatched row has gone; the line itself may be shown
                    key = Mid$(ref, 5)
                    If refIndex.Exists(key) Then idx = refIndex(key) Else ref = key
                ElseIf refIndex.Exists("UNM|" & ref) Then
                    ' The line is gone and its saved amounts are on an Unmatched row
                    idx = refIndex("UNM|" & ref)
                End If
                If idx = 0 Then
                    desc = CleanText(rd(i, modStore.R_DESC))
                    blk = UCase$(CleanText(rd(i, modStore.R_BLOCK)))
                    If BlockIndex(blk) = 5 Then blk = "OTHER"
                    code = "": io = "": lk = ""
                    parts = Split(ref & "|||", "|")
                    Select Case parts(0)
                        Case "MPW"
                            idx = AddLine("UNMATCHED", blk, CStr(parts(2)), "MP", CStr(parts(1)), desc)
                            L(idx).IOCode = CStr(parts(2))
                            reason = "manpower line no longer in the source"
                        Case "RNF"
                            idx = AddLine("UNMATCHED", "MP", CStr(parts(2)), "MP", "", desc)
                            L(idx).IOCode = CStr(parts(2))
                            reason = "Rate-not-found row no longer in the source"
                        Case "OOR"
                            idx = AddLine("UNMATCHED", blk, "", "MP", "", desc)
                            reason = "no manpower months outside the project dates now"
                        Case Else
                            If BlockIndex(CStr(parts(0))) < 5 Or parts(0) = "OTHER" Then blk = CStr(parts(0))
                            code = CStr(parts(1)): io = CStr(parts(2)): lk = CStr(parts(3))
                            idx = AddLine("UNMATCHED", blk, io, code, lk, desc)
                            L(idx).IOCode = IOCodeText(io, code)
                            If Not lineIndex.Exists(ref) Then
                                reason = "no longer in budget"
                            ElseIf L(lineIndex(ref)).Kind = "HIDDEN" Then
                                reason = "line now nets to $0"
                            Else
                                reason = "line no longer shown"
                            End If
                    End Select
                    L(idx).Label = "Unmatched: " & desc & " (" & reason & "; remark only)"
                    L(idx).RemarkRef = ref
                    refIndex.Add ref, idx
                    warnUnmatched = warnUnmatched + 1
                End If
                If stored <> L(idx).RemarkRef Or L(idx).Remark <> "" Then remarkMoved = True
                If L(idx).Remark = "" Then
                    L(idx).Remark = txt
                ElseIf InStr(1, L(idx).Remark, txt, vbBinaryCompare) = 0 Then
                    L(idx).Remark = L(idx).Remark & " | " & txt
                End If
            End If
        End If
    Next i
End Sub

' Display order inside each block. MP: by IO-Code (budget lines, then the
' manpower lines that share the IO), then Unmatched. Other blocks: budget
' lines, then charged manpower lines, then Unmatched.
Private Sub SortLines()
    Dim i As Long, j As Long, t As Long, rank As Long, b As Long

    ReDim ord(1 To IIf(nL > 0, nL, 1))
    nOrd = 0
    For i = 1 To nL
        If L(i).Kind <> "HIDDEN" Then
            Select Case L(i).Kind
                Case "INPUT", "BUDGET_MP": rank = 1
                Case "MANPOWER": rank = 2
                Case "RNF": rank = 3
                Case "UNMATCHED": rank = 8
                Case Else: rank = 9
            End Select
            b = BlockIndex(L(i).Block)
            If L(i).Block = "MP" And rank < 8 Then
                L(i).SortKey = b & "|0|" & L(i).IOCode & "|" & rank & "|" & Format$(L(i).Seq, "000000")
            Else
                L(i).SortKey = b & "|" & rank & "|" & L(i).IOCode & "|" & Format$(L(i).Seq, "000000")
            End If
            If b = 5 And L(i).Kind <> "MP_OOR" And L(i).Kind <> "MANPOWER" Then otherCodes(CodeText(L(i).Code)) = True
            nOrd = nOrd + 1
            ord(nOrd) = i
        End If
    Next i
    For i = 2 To nOrd
        t = ord(i)
        j = i - 1
        Do While j >= 1
            If L(ord(j)).SortKey <= L(t).SortKey Then Exit Do
            ord(j + 1) = ord(j)
            j = j - 1
        Loop
        ord(j + 1) = t
    Next i
End Sub

'==============================================================================
' Writing the grid
'==============================================================================
Private Sub WriteGrid(ByVal ws As Worksheet, ByVal src As String, ByVal mainIO As String, _
                      ByVal dispName As String)
    Dim anchor As Range, r0 As Long, c0 As Long, firstC As Long, lastC As Long, lastRow As Long
    Dim arr() As Variant, lay() As Variant, nLay As Long
    Dim b As Long, i As Long, k As Long, r As Long, c As Long, sc As Long, q As Long, y As Long
    Dim blocks As Variant, shown(1 To 5) As Boolean, hdrRow(1 To 5) As Long
    Dim lf(1 To 5) As Long, ll(1 To 5) As Long, hasBud(1 To 5) As Boolean
    Dim rP(1 To 6) As Long, rA(1 To 6) As Long, rV(1 To 6) As Long, rVP(1 To 6) As Long
    Dim merges As Collection, runStart As Long, runText As String, it As Variant
    Dim title As String, s As String

    blocks = Array("", "MP", "OM", "EQ", "CO", "OTHER")
    Set anchor = ThisWorkbook.Names("TL_Anchor").RefersToRange
    r0 = anchor.Row
    c0 = anchor.Column
    firstC = c0 + OFF_FIRST
    lastC = firstC + nCols - 1

    ' Row plan
    For i = 1 To nOrd
        shown(BlockIndex(L(ord(i)).Block)) = True
    Next i
    For b = 1 To 5
        If blockHasAct(b) Then shown(b) = True
    Next b
    r = r0 + 2
    For b = 1 To 5
        If shown(b) Then
            hdrRow(b) = r
            r = r + 1
            lf(b) = r
            For i = 1 To nOrd
                If BlockIndex(L(ord(i)).Block) = b Then
                    L(ord(i)).Row = r
                    If L(ord(i)).HasBudget Then hasBud(b) = True
                    r = r + 1
                End If
            Next i
            ll(b) = r - 1
            rP(b) = r: rA(b) = r + 1: rV(b) = r + 2: rVP(b) = r + 3
            r = r + 4
        End If
    Next b
    rP(6) = r: rA(6) = r + 1: rV(6) = r + 2: rVP(6) = r + 3
    lastRow = r + 3

    ClearGrid ws, r0, c0
    ReDim arr(1 To lastRow - r0 + 1, 1 To lastC - c0 + 1)

    ' Header rows (spec 5.2)
    title = dispName & "   |   Main IO " & mainIO & "   |   " & IIf(fyBasis, "Financial Year", "Calendar Year")
    If soaMonth > 0 Then title = title & "   |   Variance to Latest SOA Month " & Format$(soaMonth, "mmm yyyy")
    PutV arr, r0, c0, r0, c0, title
    PutV arr, r0, c0, r0 + 1, c0, "Line"
    PutV arr, r0, c0, r0 + 1, c0 + OFF_IO, "IO" & ChrW(183) & "Code"
    PutV arr, r0, c0, r0 + 1, c0 + OFF_BUDGET, "Budget"
    For c = 1 To nCols
        sc = firstC + c - 1
        Select Case colKind(c)
            Case "M": PutV arr, r0, c0, r0 + 1, sc, CDbl(MonthDate(colMonth(c)))
            Case "Q"
                ' Quarter label sits over its Q Total, so it stays visible when the quarter is collapsed
                PutV arr, r0, c0, r0, sc, qLbl(colQ(c))
                PutV arr, r0, c0, r0 + 1, sc, "Q Total"
            Case "Y"
                PutV arr, r0, c0, r0 + 1, sc, yLbl(colY(c))
                If colY(c) = 1 Then PutV arr, r0, c0, r0, sc, IIf(fyBasis, "Financial year totals", "Calendar year totals")
            Case "P": PutV arr, r0, c0, r0 + 1, sc, "Project Total"
        End Select
    Next c

    ' Blocks
    Set merges = New Collection
    For b = 1 To 5
        If shown(b) Then
            If b = 5 And otherCodes.Count > 0 Then
                PutV arr, r0, c0, hdrRow(b), c0, "OTHER (" & JoinKeys(otherCodes) & ")"
            Else
                PutV arr, r0, c0, hdrRow(b), c0, blocks(b)
            End If
            runStart = 0: runText = ""
            For i = 1 To nOrd
                k = ord(i)
                If BlockIndex(L(k).Block) = b Then
                    WriteLine arr, r0, c0, firstC, k
                    ' IO-Code: written on the first row of each run, merged later
                    If L(k).IOCode <> "" And L(k).IOCode = runText Then
                        ' same run
                    Else
                        If runStart > 0 And L(k).Row - 1 > runStart Then merges.Add Array(runStart, L(k).Row - 1)
                        runStart = 0: runText = L(k).IOCode
                        If runText <> "" Then
                            runStart = L(k).Row
                            PutV arr, r0, c0, L(k).Row, c0 + OFF_IO, runText
                        End If
                    End If
                End If
            Next i
            If runStart > 0 And ll(b) > runStart Then merges.Add Array(runStart, ll(b))
            WriteBlockSubtotals arr, r0, c0, firstC, b, lf(b), ll(b), hasBud(b), rP(b), rA(b), rV(b), rVP(b)
            PutV arr, r0, c0, rP(b), c0, blocks(b) & " Projected"
            PutV arr, r0, c0, rA(b), c0, blocks(b) & " Actual"
            PutV arr, r0, c0, rV(b), c0, blocks(b) & " Variance ($)"
            PutV arr, r0, c0, rVP(b), c0, blocks(b) & " Variance (%)"
        End If
    Next b
    WriteProjectTotal arr, r0, c0, firstC, shown, rP, rA, rV, rVP
    PutV arr, r0, c0, rP(6), c0, "PROJECT TOTAL Projected"
    PutV arr, r0, c0, rA(6), c0, "PROJECT TOTAL Actual"
    PutV arr, r0, c0, rV(6), c0, "PROJECT TOTAL Variance ($)"
    PutV arr, r0, c0, rVP(6), c0, "PROJECT TOTAL Variance (%)"

    ' Write everything in one go (labels and IO-Code as text)
    ws.Range(ws.Cells(r0, c0), ws.Cells(lastRow, c0)).NumberFormat = "@"
    ws.Range(ws.Cells(r0, c0 + OFF_IO), ws.Cells(lastRow, c0 + OFF_IO)).NumberFormat = "@"
    ws.Range(ws.Cells(r0, c0), ws.Cells(lastRow, lastC)).Formula = arr

    FormatGrid ws, r0, c0, firstC, lastC, lastRow, shown, hdrRow, rP, rA, rV, rVP
    For Each it In merges
        With ws.Range(ws.Cells(it(0), c0 + OFF_IO), ws.Cells(it(1), c0 + OFF_IO))
            .Merge
            .VerticalAlignment = xlCenter
        End With
    Next it
    WriteNotes ws, firstC, shown, rA, rP
    WriteRemarks ws, r0, c0, lastC, lastRow, shown, hdrRow, rP, rVP

    ' Outline: each quarter's months collapse into its total (summary on the right)
    ws.Outline.SummaryColumn = xlSummaryOnRight
    For q = 1 To nQ
        ws.Range(ws.Columns(firstC + qFirstC(q) - 1), ws.Columns(firstC + qLastC(q) - 1)).Group
    Next q
    GroupTopRows ws

    ws.Hyperlinks.Add Anchor:=ws.Cells(lastRow + 2, c0), Address:="", _
        SubAddress:="'" & DASH_SHEET & "'!A1", TextToDisplay:="Back to Top"

    ' KPI names (spec 12)
    ThisWorkbook.Names.Add Name:="TL_TotalBudget", _
        RefersTo:="='" & DASH_SHEET & "'!" & ws.Cells(rP(6), c0 + OFF_BUDGET).Address
    ThisWorkbook.Names.Add Name:="TL_TotalActual", _
        RefersTo:="='" & DASH_SHEET & "'!" & ws.Cells(rA(6), firstC + pCol - 1).Address
    ThisWorkbook.Names.Add Name:="TL_TotalProjected", _
        RefersTo:="='" & DASH_SHEET & "'!" & ws.Cells(rP(6), firstC + pCol - 1).Address

    ' Layout map for Save / Re-assign / Delete
    ReDim lay(1 To IIf(nOrd > 0, nOrd, 1), 1 To 8)
    nLay = 0
    For i = 1 To nOrd
        k = ord(i)
        nLay = nLay + 1
        lay(nLay, 1) = L(k).Row
        lay(nLay, 2) = L(k).Kind
        lay(nLay, 3) = L(k).Block
        lay(nLay, 4) = L(k).IO
        lay(nLay, 5) = L(k).Code
        lay(nLay, 6) = L(k).LineKey
        lay(nLay, 7) = L(k).Desc
        lay(nLay, 8) = L(k).RemarkRef
    Next i
    modStore.WriteLayout lay, nLay, r0, lastRow + 2, firstC, lastC, lastC + 1
End Sub

Private Sub PutV(ByRef arr() As Variant, ByVal r0 As Long, ByVal c0 As Long, _
                 ByVal r As Long, ByVal c As Long, ByVal v As Variant)
    arr(r - r0 + 1, c - c0 + 1) = v
End Sub

Private Sub WriteLine(ByRef arr() As Variant, ByVal r0 As Long, ByVal c0 As Long, _
                      ByVal firstC As Long, ByVal k As Long)
    Dim c As Long, sc As Long, r As Long, q As Long, y As Long, list As String
    r = L(k).Row
    PutV arr, r0, c0, r, c0, L(k).Label
    If L(k).HasBudget Then PutV arr, r0, c0, r, c0 + OFF_BUDGET, L(k).Budget
    Select Case L(k).Kind
        Case "BUDGET_MP", "RNF"
            Exit Sub                        ' budget only / no amount
        Case "MP_OOR"
            PutV arr, r0, c0, r, firstC + pCol - 1, L(k).OORAmount
            Exit Sub
    End Select
    For c = 1 To nCols
        sc = firstC + c - 1
        Select Case colKind(c)
            Case "M"
                If Not IsEmpty(L(k).Vals(colMonth(c))) Then PutV arr, r0, c0, r, sc, L(k).Vals(colMonth(c))
            Case "Q"
                q = colQ(c)
                PutV arr, r0, c0, r, sc, "=SUM(" & Ad(r, firstC + qFirstC(q) - 1) & ":" & Ad(r, firstC + qLastC(q) - 1) & ")"
            Case "Y"
                PutV arr, r0, c0, r, sc, "=SUM(" & QTotList(r, firstC, colY(c)) & ")"
            Case "P"
                PutV arr, r0, c0, r, sc, "=SUM(" & YList(r, firstC) & ")" & FmtExtra(L(k).OORAmount)
        End Select
    Next c
End Sub

' Projected / Actual / Variance ($) / Variance (%) for one block (spec 6, 14.2)
Private Sub WriteBlockSubtotals(ByRef arr() As Variant, ByVal r0 As Long, ByVal c0 As Long, _
        ByVal firstC As Long, ByVal b As Long, ByVal lf As Long, ByVal ll As Long, ByVal hasBud As Boolean, _
        ByVal rP As Long, ByVal rA As Long, ByVal rV As Long, ByVal rVP As Long)
    Dim c As Long, sc As Long, q As Long, y As Long, ex As Double, wroteV As Boolean

    If hasBud Then PutV arr, r0, c0, rP, c0 + OFF_BUDGET, _
        "=SUM(" & Ad(lf, c0 + OFF_BUDGET) & ":" & Ad(ll, c0 + OFF_BUDGET) & ")"
    For c = 1 To nCols
        sc = firstC + c - 1
        wroteV = False
        If ll >= lf Then PutV arr, r0, c0, rP, sc, "=SUM(" & Ad(lf, sc) & ":" & Ad(ll, sc) & ")"
        Select Case colKind(c)
            Case "M"
                If MonthLE(colMonth(c)) Then
                    PutV arr, r0, c0, rA, sc, blockAct(b, colMonth(c))
                    PutV arr, r0, c0, rV, sc, "=" & Ad(rA, sc) & "-" & Ad(rP, sc)
                    wroteV = True
                End If
            Case "Q"
                q = colQ(c)
                If MonthLE(colMonth(qFirstC(q))) Then
                    PutV arr, r0, c0, rA, sc, "=SUM(" & Ad(rA, firstC + qFirstC(q) - 1) & ":" & Ad(rA, firstC + qLastC(q) - 1) & ")"
                    PutV arr, r0, c0, rV, sc, "=SUM(" & Ad(rV, firstC + qFirstC(q) - 1) & ":" & Ad(rV, firstC + qLastC(q) - 1) & ")"
                    wroteV = True
                End If
            Case "Y"
                y = colY(c)
                ex = yExtra(b, y)
                If YearLE(y) Or ex <> 0 Then
                    PutV arr, r0, c0, rA, sc, "=SUM(" & QTotList(rA, firstC, y) & ")" & FmtExtra(ex)
                    If ex <> 0 Then
                        PutV arr, r0, c0, rV, sc, "=SUM(" & QTotList(rV, firstC, y) & ")+(" & Ad(rA, sc) & "-SUM(" & QTotList(rA, firstC, y) & "))"
                    Else
                        PutV arr, r0, c0, rV, sc, "=SUM(" & QTotList(rV, firstC, y) & ")"
                    End If
                    wroteV = True
                End If
            Case "P"
                ex = pExtra(b)
                PutV arr, r0, c0, rA, sc, "=SUM(" & YList(rA, firstC) & ")" & FmtExtra(ex)
                If soaMonth > 0 Or ex <> 0 Then
                    If ex <> 0 Then
                        PutV arr, r0, c0, rV, sc, "=SUM(" & YList(rV, firstC) & ")+(" & Ad(rA, sc) & "-SUM(" & YList(rA, firstC) & "))"
                    Else
                        PutV arr, r0, c0, rV, sc, "=SUM(" & YList(rV, firstC) & ")"
                    End If
                    wroteV = True
                End If
        End Select
        If wroteV Then PutV arr, r0, c0, rVP, sc, PctFormula(rA, rV, sc)
    Next c
End Sub

' PROJECT TOTAL rows: sums of the block rows; variance built the same way.
Private Sub WriteProjectTotal(ByRef arr() As Variant, ByVal r0 As Long, ByVal c0 As Long, ByVal firstC As Long, _
        ByRef shown() As Boolean, ByRef rP() As Long, ByRef rA() As Long, ByRef rV() As Long, ByRef rVP() As Long)
    Dim c As Long, sc As Long, q As Long, y As Long, b As Long, ex As Double, wroteV As Boolean
    Dim listP As String, listA As String, listB As String

    For b = 1 To 5
        If shown(b) Then
            listB = listB & IIf(listB = "", "", ",") & Ad(rP(b), c0 + OFF_BUDGET)
        End If
    Next b
    If listB <> "" Then PutV arr, r0, c0, rP(6), c0 + OFF_BUDGET, "=SUM(" & listB & ")"

    For c = 1 To nCols
        sc = firstC + c - 1
        listP = "": listA = ""
        For b = 1 To 5
            If shown(b) Then
                listP = listP & IIf(listP = "", "", ",") & Ad(rP(b), sc)
                listA = listA & IIf(listA = "", "", ",") & Ad(rA(b), sc)
            End If
        Next b
        If listP <> "" Then PutV arr, r0, c0, rP(6), sc, "=SUM(" & listP & ")"
        wroteV = False
        Select Case colKind(c)
            Case "M"
                If MonthLE(colMonth(c)) Then
                    PutV arr, r0, c0, rA(6), sc, "=SUM(" & listA & ")"
                    PutV arr, r0, c0, rV(6), sc, "=" & Ad(rA(6), sc) & "-" & Ad(rP(6), sc)
                    wroteV = True
                End If
            Case "Q"
                q = colQ(c)
                If MonthLE(colMonth(qFirstC(q))) Then
                    PutV arr, r0, c0, rA(6), sc, "=SUM(" & listA & ")"
                    PutV arr, r0, c0, rV(6), sc, "=SUM(" & Ad(rV(6), firstC + qFirstC(q) - 1) & ":" & Ad(rV(6), firstC + qLastC(q) - 1) & ")"
                    wroteV = True
                End If
            Case "Y"
                y = colY(c)
                ex = 0
                For b = 1 To 5
                    ex = ex + yExtra(b, y)
                Next b
                If YearLE(y) Or ex <> 0 Then
                    PutV arr, r0, c0, rA(6), sc, "=SUM(" & listA & ")"
                    If ex <> 0 Then
                        PutV arr, r0, c0, rV(6), sc, "=SUM(" & QTotList(rV(6), firstC, y) & ")+(" & Ad(rA(6), sc) & "-SUM(" & QTotList(rA(6), firstC, y) & "))"
                    Else
                        PutV arr, r0, c0, rV(6), sc, "=SUM(" & QTotList(rV(6), firstC, y) & ")"
                    End If
                    wroteV = True
                End If
            Case "P"
                ex = 0
                For b = 1 To 5
                    ex = ex + pExtra(b)
                Next b
                PutV arr, r0, c0, rA(6), sc, "=SUM(" & listA & ")"
                If soaMonth > 0 Or ex <> 0 Then
                    If ex <> 0 Then
                        PutV arr, r0, c0, rV(6), sc, "=SUM(" & YList(rV(6), firstC) & ")+(" & Ad(rA(6), sc) & "-SUM(" & YList(rA(6), firstC) & "))"
                    Else
                        PutV arr, r0, c0, rV(6), sc, "=SUM(" & YList(rV(6), firstC) & ")"
                    End If
                    wroteV = True
                End If
        End Select
        If wroteV Then PutV arr, r0, c0, rVP(6), sc, PctFormula(rA(6), rV(6), sc)
    Next c
End Sub

' Variance (%) = Variance / Projected-to-SOA. All actuals are on or before the
' Latest SOA Month, so Projected-to-SOA = Actual - Variance.
Private Function PctFormula(ByVal rA As Long, ByVal rV As Long, ByVal sc As Long) As String
    Dim a1 As String, v1 As String
    a1 = Ad(rA, sc): v1 = Ad(rV, sc)
    PctFormula = "=IF(ROUND(" & a1 & "-" & v1 & ",2)=0,""""," & v1 & "/(" & a1 & "-" & v1 & "))"
End Function

Private Function MonthLE(ByVal mi As Long) As Boolean
    If soaMonth > 0 Then MonthLE = (MonthDate(mi) <= soaMonth)
End Function

Private Function YearLE(ByVal y As Long) As Boolean
    Dim q As Long
    For q = 1 To nQ
        If qY(q) = y Then
            If MonthLE(colMonth(qFirstC(q))) Then YearLE = True
        End If
    Next q
End Function

Private Function QTotList(ByVal r As Long, ByVal firstC As Long, ByVal y As Long) As String
    Dim q As Long, s As String
    For q = 1 To nQ
        If qY(q) = y Then s = s & IIf(s = "", "", ",") & Ad(r, firstC + qTotC(q) - 1)
    Next q
    QTotList = s
End Function

Private Function YList(ByVal r As Long, ByVal firstC As Long) As String
    Dim y As Long, s As String
    For y = 1 To nY
        s = s & IIf(s = "", "", ",") & Ad(r, firstC + yCol(y) - 1)
    Next y
    YList = s
End Function

Private Function Ad(ByVal r As Long, ByVal c As Long) As String
    Ad = ColLetter(c) & r
End Function

Private Sub FormatGrid(ByVal ws As Worksheet, ByVal r0 As Long, ByVal c0 As Long, ByVal firstC As Long, _
        ByVal lastC As Long, ByVal lastRow As Long, ByRef shown() As Boolean, ByRef hdrRow() As Long, _
        ByRef rP() As Long, ByRef rA() As Long, ByRef rV() As Long, ByRef rVP() As Long)
    Dim full As Variant, tint As Variant, c As Long, sc As Long, b As Long, i As Long, k As Long, r As Long
    Dim grid As Range, rng As Range, colour As Long, tot As Variant

    full = Array(0, HexColor("BDD7EE"), HexColor("C6E0B4"), HexColor("FFE699"), HexColor("F8CBAD"))
    tint = Array(0, HexColor("DDEBF7"), HexColor("E2EFDA"), HexColor("FFF2CC"), HexColor("FCE4D6"))
    Set grid = ws.Range(ws.Cells(r0, c0), ws.Cells(lastRow, lastC))
    grid.Font.Size = 10

    ' Column colours: months light tint, quarter totals full colour (fixed per
    ' quarter number), FY/CY totals neutral grey, Project Total light grey.
    For c = 1 To nCols
        sc = firstC + c - 1
        Select Case colKind(c)
            Case "M": colour = tint(qNum(colQ(c)))
            Case "Q": colour = full(qNum(colQ(c)))
            Case "Y": colour = HexColor("E1E4E3")
            Case Else: colour = HexColor("EEF2F0")
        End Select
        ws.Range(ws.Cells(r0, sc), ws.Cells(lastRow, sc)).Interior.Color = colour
        ws.Cells(r0 + 1, sc).ColumnWidth = IIf(colKind(c) = "M", 11, IIf(colKind(c) = "P", 14, 12))
        If colKind(c) <> "M" Then ws.Range(ws.Cells(r0 + 2, sc), ws.Cells(lastRow, sc)).Font.Bold = True
    Next c

    ' Headers
    With ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + 1, lastC))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + 1, c0 + OFF_BUDGET)).Interior.Color = HexColor("EEF2F0")
    ws.Cells(r0, c0).HorizontalAlignment = xlLeft
    ws.Cells(r0 + 1, c0).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(r0 + 1, firstC), ws.Cells(r0 + 1, lastC)).NumberFormat = "mmm yy"
    If nY > 0 Then
        ws.Range(ws.Cells(r0, firstC + yCol(1) - 1), ws.Cells(r0, firstC + yCol(nY) - 1)).HorizontalAlignment = xlCenterAcrossSelection
    End If

    ' Numbers
    ws.Range(ws.Cells(r0 + 2, c0 + OFF_BUDGET), ws.Cells(lastRow, lastC)).NumberFormat = FMT_MONEY
    ws.Range(ws.Cells(r0 + 2, c0 + OFF_IO), ws.Cells(lastRow, c0 + OFF_IO)).Font.Size = 9
    ws.Range(ws.Cells(r0 + 2, c0 + OFF_IO), ws.Cells(lastRow, c0 + OFF_IO)).Font.Color = HexColor("4B5A54")

    ' Line rows
    For i = 1 To nOrd
        k = ord(i)
        r = L(k).Row
        Select Case L(k).Kind
            Case "INPUT"
                Set rng = MonthCells(ws, r, firstC)
                rng.Interior.Color = HexColor("FFFFE0")
                rng.Locked = False
                With rng.Validation
                    .Delete
                    .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                         Formula1:="-1000000000000", Formula2:="1000000000000"
                    .IgnoreBlank = True
                    .ErrorTitle = "Numbers only"
                    .ErrorMessage = "Type a dollar amount (negatives allowed), or leave the cell blank."
                End With
            Case "BUDGET_MP"
                With MonthCells(ws, r, firstC).Interior
                    .Pattern = xlPatternLightUp
                    .PatternColor = HexColor("C9D1CD")
                    .Color = HexColor("FFFFFF")
                End With
            Case "MANPOWER", "RNF"
                MonthCells(ws, r, firstC).Interior.Color = HexColor("F1F4F3")
                If L(k).Kind = "RNF" Then ws.Cells(r, c0).Font.Color = HexColor("7A4B00")
            Case "UNMATCHED", "MP_OOR"
                ws.Range(ws.Cells(r, c0), ws.Cells(r, lastC)).Font.Italic = True
                ws.Cells(r, c0).Font.Color = HexColor("7A4B00")
        End Select
    Next i

    ' Block headers, subtotals, project total
    For b = 1 To 5
        If shown(b) Then
            With ws.Range(ws.Cells(hdrRow(b), c0), ws.Cells(hdrRow(b), lastC))
                .Interior.Color = HexColor("D8ECE8")
                .Font.Bold = True
            End With
            With ws.Range(ws.Cells(rP(b), c0), ws.Cells(rVP(b), lastC))
                .Interior.Color = HexColor("F2F5F4")
                .Font.Bold = True
            End With
            ws.Range(ws.Cells(rP(b), c0), ws.Cells(rP(b), lastC)).Borders(xlEdgeTop).Weight = xlMedium
            ws.Range(ws.Cells(rVP(b), c0 + OFF_BUDGET), ws.Cells(rVP(b), lastC)).NumberFormat = FMT_PCT
        End If
    Next b
    With ws.Range(ws.Cells(rP(6), c0), ws.Cells(rVP(6), lastC))
        .Interior.Color = HexColor("17211D")
        .Font.Color = HexColor("FFFFFF")
        .Font.Bold = True
    End With
    ws.Range(ws.Cells(rVP(6), c0 + OFF_BUDGET), ws.Cells(rVP(6), lastC)).NumberFormat = FMT_PCT

    ' Light grid lines; no vertical lines inside B:E so the Line text reads as one column
    With grid.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = HexColor("D5DDD9")
    End With
    With ws.Range(ws.Cells(r0, c0 + OFF_IO), ws.Cells(lastRow, lastC)).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = HexColor("D5DDD9")
    End With
    grid.BorderAround xlContinuous, xlThin, , HexColor("A9B6B0")
    With ws.Range(ws.Cells(r0, c0 + OFF_IO), ws.Cells(lastRow, c0 + OFF_IO)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = HexColor("A9B6B0")
    End With
End Sub

' Rows TOP_FIRST_ROW:TOP_LAST_ROW (selector, KPI cards, buttons) are grouped so
' they collapse with the +/- beside row 1 and the timeline gets more room.
' ClearGrid clears the whole outline, so this runs on every build; the rows
' are left expanded. The buttons are set to move and size with their cells so
' they hide with the rows instead of floating over the timeline. Frozen panes
' are not touched.
Private Sub GroupTopRows(ByVal ws As Worksheet)
    Dim shp As Shape
    ws.Outline.SummaryRow = xlSummaryAbove
    ws.Range(ws.Rows(TOP_FIRST_ROW), ws.Rows(TOP_LAST_ROW)).Group
    For Each shp In ws.Shapes
        If Left$(shp.Name, 6) = "btnTL_" Then shp.Placement = xlMoveAndSize
    Next shp
End Sub

' Remarks column straight after Project Total (spec 5.5). Every line row gets
' an unlocked, wrapped cell; block headers and subtotal rows take the row's
' band colour. The column uses the standard font size, so after the row
' autofit below rows keep their usual height and grow when a long remark is
' typed. Remarks are written as text, so one starting with "=" stays text.
Private Sub WriteRemarks(ByVal ws As Worksheet, ByVal r0 As Long, ByVal c0 As Long, ByVal lastC As Long, _
        ByVal lastRow As Long, ByRef shown() As Boolean, ByRef hdrRow() As Long, _
        ByRef rP() As Long, ByRef rVP() As Long)
    Dim rc As Long, i As Long, k As Long, b As Long, col As Range, cell As Range
    rc = lastC + 1
    Set col = ws.Range(ws.Cells(r0, rc), ws.Cells(lastRow, rc))
    ws.Columns(rc).ColumnWidth = REMARK_WIDTH
    With col
        .NumberFormat = "@"
        .Font.Size = Application.StandardFontSize
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Interior.Color = HexColor("EEF2F0")
    End With
    With ws.Range(ws.Cells(r0, rc), ws.Cells(r0 + 1, rc))
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(r0 + 1, rc).Value = "Remarks"
    For i = 1 To nOrd
        k = ord(i)
        Set cell = ws.Cells(L(k).Row, rc)
        cell.Interior.Color = HexColor("FFFFE0")
        cell.Locked = False
        If L(k).Kind = "UNMATCHED" Or L(k).Kind = "MP_OOR" Then cell.Font.Italic = True
        If L(k).Remark <> "" Then cell.Value = L(k).Remark
    Next i
    For b = 1 To 5
        If shown(b) Then
            ws.Cells(hdrRow(b), rc).Interior.Color = HexColor("D8ECE8")
            ws.Range(ws.Cells(rP(b), rc), ws.Cells(rVP(b), rc)).Interior.Color = HexColor("F2F5F4")
            ws.Cells(rP(b), rc).Borders(xlEdgeTop).Weight = xlMedium
        End If
    Next b
    With ws.Range(ws.Cells(rP(6), rc), ws.Cells(rVP(6), rc))
        .Interior.Color = HexColor("17211D")
    End With
    With col.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = HexColor("D5DDD9")
    End With
    col.BorderAround xlContinuous, xlThin, , HexColor("A9B6B0")
    ws.Range(ws.Rows(r0 + 2), ws.Rows(lastRow)).AutoFit
End Sub

Private Function MonthCells(ByVal ws As Worksheet, ByVal r As Long, ByVal firstC As Long) As Range
    Dim q As Long, rng As Range, part As Range
    For q = 1 To nQ
        Set part = ws.Range(ws.Cells(r, firstC + qFirstC(q) - 1), ws.Cells(r, firstC + qLastC(q) - 1))
        If rng Is Nothing Then Set rng = part Else Set rng = Union(rng, part)
    Next q
    Set MonthCells = rng
End Function

Private Sub WriteNotes(ByVal ws As Worksheet, ByVal firstC As Long, ByRef shown() As Boolean, _
                       ByRef rA() As Long, ByRef rP() As Long)
    Dim i As Long, k As Long, b As Long, ptC As Long, anyAct As Boolean
    ptC = firstC + pCol - 1
    For i = 1 To nOrd
        k = ord(i)
        Select Case L(k).Kind
            Case "UNMATCHED"
                If L(k).OORNote <> "" Then AddNote ws.Cells(L(k).Row, ptC), _
                    "Saved months outside the project dates (counted in the Project Total only):" & L(k).OORNote
            Case "MP_OOR"
                AddNote ws.Cells(L(k).Row, ptC), _
                    "Manpower months outside the project dates (counted in the Project Total only):" & L(k).OORNote
            Case "RNF"
                AddNote ws.Cells(L(k).Row, firstC - OFF_FIRST), _
                    L(k).RnfRows & " manpower month row(s) have Rate Status ""Rate not found"" (" & _
                    Format$(L(k).RnfFTE, "0.0#") & " FTE in total). No amount is projected for them. " & _
                    "Check the rate grade in tbl_ManpowerRate."
        End Select
    Next i
    For b = 1 To 5
        If shown(b) And actNote(b) <> "" Then
            anyAct = True
            AddNote ws.Cells(rA(b), ptC), "Actuals posted outside the project dates (counted here):" & actNote(b)
        End If
    Next b
    If anyAct Then AddNote ws.Cells(rA(6), ptC), _
        "Includes actuals posted outside the project dates; see the notes on the category Actual rows."
End Sub

Private Sub AddNote(ByVal cell As Range, ByVal noteText As String)
    On Error Resume Next
    cell.ClearComments
    cell.AddComment noteText
    cell.Comment.Shape.TextFrame.AutoSize = True
    On Error GoTo 0
End Sub

' Clears everything from TL_Anchor's row down: values, formats, merges,
' notes, validation, links and the whole outline (spec 5.1). GroupTopRows
' puts the row group back after the build.
Private Sub ClearGrid(ByVal ws As Worksheet, ByVal r0 As Long, ByVal c0 As Long)
    Dim lastRow As Long, lastCol As Long, rng As Range
    With ws.UsedRange
        lastRow = .Row + .Rows.Count - 1
        lastCol = .Column + .Columns.Count - 1
    End With
    If lastRow < r0 + 1 Then lastRow = r0 + 1
    If lastCol < c0 + OFF_FIRST Then lastCol = c0 + OFF_FIRST
    Set rng = ws.Range(ws.Rows(r0), ws.Rows(lastRow))
    On Error Resume Next
    rng.Hyperlinks.Delete
    rng.ClearComments
    rng.Validation.Delete
    rng.UnMerge
    ws.Cells.ClearOutline
    On Error GoTo 0
    rng.Clear
    rng.RowHeight = ws.StandardHeight
    With ws.Range(ws.Columns(c0 + OFF_FIRST), ws.Columns(lastCol))
        .Hidden = False
        .ColumnWidth = ws.StandardWidth
    End With
End Sub

Private Sub ResetKpiNames()
    ThisWorkbook.Names.Add Name:="TL_TotalBudget", RefersTo:="="""""
    ThisWorkbook.Names.Add Name:="TL_TotalActual", RefersTo:="="""""
    ThisWorkbook.Names.Add Name:="TL_TotalProjected", RefersTo:="="""""
End Sub

'==============================================================================
' Protection (spec 10.4, 11)
'==============================================================================
Public Sub ProtectDashboard()
    Dim ws As Worksheet
    Set ws = DashSheet()
    On Error Resume Next
    ws.Unprotect
    ws.Range(BASIS_CELL & "," & SEL_PROJECT & "," & SEL_MAINIO).Locked = False
    ws.EnableOutlining = True
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    On Error GoTo 0
End Sub

Public Sub BeginWork()
    mCalc = Application.Calculation
    mScreen = Application.ScreenUpdating
    mEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    DashSheet().Unprotect
End Sub

Public Sub EndWork()
    Application.Calculation = mCalc
    Application.EnableEvents = mEvents
    Application.ScreenUpdating = mScreen
    Application.StatusBar = False
    Application.Calculate
End Sub

'==============================================================================
' Shared helpers (also used by modStore)
'==============================================================================
Public Function DashSheet() As Worksheet
    Set DashSheet = ThisWorkbook.Worksheets(DASH_SHEET)
End Function

Public Function FindTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

' Data body of a table as a 2-D array, or Empty when the table has no rows.
Public Function TableData(ByVal tableName As String, ByRef lo As ListObject) As Variant
    Set lo = FindTable(tableName)
    If lo Is Nothing Then Err.Raise vbObjectError + 513, , "Table " & tableName & " was not found."
    If lo.DataBodyRange Is Nothing Then
        TableData = Empty
    Else
        TableData = lo.DataBodyRange.Value
    End If
End Function

Public Function ColIndex(ByVal lo As ListObject, ByVal header As String) As Long
    On Error Resume Next
    ColIndex = lo.ListColumns(header).Index
    On Error GoTo 0
    If ColIndex = 0 Then Err.Raise vbObjectError + 514, , "Column """ & header & """ was not found in " & lo.Name & "."
End Function

Public Function BlockOf(ByVal code As String) As String
    Select Case UCase$(code)
        Case "MP", "OM", "EQ", "CO": BlockOf = UCase$(code)
        Case Else: BlockOf = "OTHER"
    End Select
End Function

Public Function BlockIndex(ByVal blk As String) As Long
    Select Case blk
        Case "MP": BlockIndex = 1
        Case "OM": BlockIndex = 2
        Case "EQ": BlockIndex = 3
        Case "CO": BlockIndex = 4
        Case Else: BlockIndex = 5
    End Select
End Function

Public Function Identity(ByVal blk As String, ByVal code As String, ByVal io As String, ByVal lk As String) As String
    Identity = blk & "|" & code & "|" & io & "|" & lk
End Function

' IO with the budget code joined on, no space (TF25HIP136OM); blank code -> (no code)
Public Function IOCodeText(ByVal io As String, ByVal code As String) As String
    If code = "" Then IOCodeText = io & " (no code)" Else IOCodeText = io & code
End Function

Private Function CodeText(ByVal code As String) As String
    If code = "" Then CodeText = "(no code)" Else CodeText = code
End Function

' Rate categories arrive upper-cased from qManpowerRates. Show them as the rate
' table writes them: NURSE -> Nurse, ADMIN -> Admin; RA, AHS, AHS (OT/PT) as is.
Private Function DisplayCategory(ByVal cat As String) As String
    If cat = "" Then
        DisplayCategory = "Manpower"
    ElseIf Len(cat) > 3 And Not cat Like "*[!A-Z]*" Then
        DisplayCategory = StrConv(cat, vbProperCase)
    Else
        DisplayCategory = cat
    End If
End Function

' Spec 5.4: trim, collapse repeated spaces, remove non-printing characters, upper case.
Public Function NormKey(ByVal s As String) As String
    Dim i As Long, ch As Long, out As String
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1)) And &HFFFF&
        If ch = 160 Then
            out = out & " "
        ElseIf ch >= 32 And ch <> 127 Then
            out = out & Mid$(s, i, 1)
        End If
    Next i
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    NormKey = UCase$(Trim$(out))
End Function

Public Function CleanText(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    s = Replace(CStr(v), ChrW(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanText = Trim$(s)
End Function

Public Function SameText(ByVal a As Variant, ByVal b As Variant) As Boolean
    SameText = (StrComp(CleanText(a), CleanText(b), vbTextCompare) = 0)
End Function

Public Function IsBlank(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then
        IsBlank = True
    ElseIf VarType(v) = vbString Then
        IsBlank = (Trim$(v) = "")
    End If
End Function

Public Function ToDate(ByVal v As Variant, ByRef d As Date) As Boolean
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate
            d = v
            ToDate = True
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency
            If v > 0 Then
                d = CDate(v)
                ToDate = True
            End If
        Case vbString
            If IsDate(v) Then
                d = CDate(v)
                ToDate = True
            End If
    End Select
End Function

Public Function ToNumber(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsNumeric(v) Then ToNumber = CDbl(v)
End Function

Private Function NzD(ByVal v As Variant) As Double
    If Not IsEmpty(v) Then NzD = CDbl(v)
End Function

Private Function FmtMoney(ByVal x As Double) As String
    FmtMoney = Format$(x, "#,##0;(#,##0)")
End Function

' "+123.45" / "-123.45" for appending a constant to a formula; "" for zero.
Private Function FmtExtra(ByVal x As Double) As String
    If Abs(x) < 0.0000001 Then
        FmtExtra = ""
    ElseIf x > 0 Then
        FmtExtra = "+" & Trim$(Str$(x))
    Else
        FmtExtra = "-" & Trim$(Str$(-x))
    End If
End Function

Private Function JoinKeys(ByVal d As Object) As String
    Dim k As Variant, s As String
    For Each k In d.Keys
        s = s & IIf(s = "", "", ", ") & k
    Next k
    JoinKeys = s
End Function

Public Function ColLetter(ByVal c As Long) As String
    Dim m As Long, s As String
    Do While c > 0
        m = (c - 1) Mod 26
        s = Chr$(65 + m) & s
        c = (c - m - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Function HexColor(ByVal hex6 As String) As Long
    HexColor = RGB(CLng("&H" & Mid$(hex6, 1, 2)), CLng("&H" & Mid$(hex6, 3, 2)), CLng("&H" & Mid$(hex6, 5, 2)))
End Function
