Attribute VB_Name = "modSetup"
'==============================================================================
' modSetup - one-off setup of the v0.9 dashboard after importing the modules
'
' Run Setup_Dashboard_v09 once (Alt+F8), then save the workbook. It:
'   1. checks that the sheets, tables and names the timeline needs exist;
'   2. adds the four buttons beside the selector (E8:F11): Load Project,
'      Save Projections, Re-assign unmatched, Delete unmatched;
'   3. unlocks C6, E6 and F6 and protects Main_Dashboard (no password) with
'      outlining enabled.
' It is safe to run again; it replaces its own buttons. It does not touch
' queries, connections or loaded tables.
'==============================================================================
Option Explicit

Public Sub Setup_Dashboard_v09()
    Dim ws As Worksheet, missing As String, t As Variant, nm As Variant, i As Long

    For Each t In Array("tblSelectedProjectMapping", "tblBudgetSelectedTest", "tblActualsSelectedTest", _
                        "qProjects_SelectedTest", "qManpower_SelectedTest", "tblProjectSelection", _
                        "tblManpowerChargeTo", "tblProjectionStore")
        If modTimeline.FindTable(CStr(t)) Is Nothing Then missing = missing & vbCrLf & "  table " & t
    Next t
    For Each nm In Array("TL_Anchor", "TL_TotalBudget", "TL_TotalActual", "TL_TotalProjected")
        If Not NameExists(CStr(nm)) Then missing = missing & vbCrLf & "  name " & nm
    Next nm
    For Each t In Array(modTimeline.DASH_SHEET, modTimeline.STORE_SHEET, "Dashboard_Settings", "Legacy_Timeline")
        If Not SheetExists(CStr(t)) Then missing = missing & vbCrLf & "  sheet " & t
    Next t
    If missing <> "" Then
        MsgBox "Setup stopped. This workbook is missing:" & missing & vbCrLf & vbCrLf & _
               "Use the v0.9_working file built from v0.8.", vbExclamation, "Setup"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(modTimeline.DASH_SHEET)
    ws.Unprotect
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 6) = "btnTL_" Then ws.Shapes(i).Delete
    Next i
    AddButton ws, "btnTL_Load", "Load Project", "LoadProject", ws.Range("E8:E9")
    AddButton ws, "btnTL_Save", "Save Projections", "SaveProjections", ws.Range("F8:F9")
    AddButton ws, "btnTL_Reassign", "Re-assign unmatched", "ReassignUnmatched", ws.Range("E10:E11")
    AddButton ws, "btnTL_Delete", "Delete unmatched", "DeleteUnmatched", ws.Range("F10:F11")
    modTimeline.ProtectDashboard

    MsgBox "Setup complete." & vbCrLf & vbCrLf & _
           "Save the workbook now (keep the .xlsm format), then choose a project in E6/F6 " & _
           "and click Load Project.", vbInformation, "Setup"
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal btnName As String, ByVal caption As String, _
                      ByVal macro As String, ByVal where As Range)
    Dim b As Object
    Set b = ws.Buttons.Add(where.Left + 2, where.Top + 2, where.Width - 4, where.Height - 4)
    b.Name = btnName
    b.Caption = caption
    b.OnAction = macro
    b.Placement = xlFreeFloating
    b.Font.Bold = True
End Sub

Private Function NameExists(ByVal nm As String) As Boolean
    Dim n As Name
    On Error Resume Next
    Set n = ThisWorkbook.Names(nm)
    NameExists = Not n Is Nothing
    On Error GoTo 0
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
