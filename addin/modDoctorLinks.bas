Attribute VB_Name = "modDoctorLinks"
'==============================================================================
' modDoctorLinks - external link tools
'------------------------------------------------------------------------------
' Finds every link to another workbook, says whether the file is still there,
' shows which cells depend on it, and breaks the links when you decide to.
'
' Depends on: modDoctorCommon, modFile, modApp (from \src)
'==============================================================================
Option Explicit

Private Const MAX_CELLS_LISTED As Long = 300

'------------------------------------------------------------------------------
' ShowExternalLinks   (menu action)
' Read-only.
'------------------------------------------------------------------------------
Public Sub ShowExternalLinks()
    Dim wb As Workbook
    Dim rows As Collection
    Dim sources As Variant
    Dim i As Long
    Dim listed As Long

    Set wb = TargetWorkbook(False)
    If wb Is Nothing Then Exit Sub

    sources = LinkList(wb)

    Set rows = NewReport()
    ReportRow rows, DOCTOR_NAME & " - External links"
    ReportRow rows, "Workbook", wb.Name

    If Not IsArray(sources) Then
        ReportHeading rows, "No links to other workbooks"
        ShowReport rows, "External links"
        MsgBox "'" & wb.Name & "' does not link to any other workbook.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    ReportHeading rows, "Linked workbooks (" & (UBound(sources) - LBound(sources) + 1) & ")"
    ReportRow rows, "Source", "File found", "Last modified"
    For i = LBound(sources) To UBound(sources)
        ReportRow rows, CStr(sources(i)), _
                  IIf(FileExists(CStr(sources(i))), "Yes", "No - link is dead"), _
                  IIf(FileExists(CStr(sources(i))), FileModified(CStr(sources(i))), "")
    Next i

    ReportHeading rows, "Cells that refer to another workbook"
    ReportRow rows, "Sheet", "Cell", "Formula"
    listed = ListLinkedCells(wb, rows)
    If listed = 0 Then
        ReportRow rows, "None found - the links may be used only by names, charts or validation."
    ElseIf listed >= MAX_CELLS_LISTED Then
        ReportRow rows, "(list stopped at " & MAX_CELLS_LISTED & " cells)"
    End If

    ShowReport rows, "External links"
End Sub

'------------------------------------------------------------------------------
' BreakExternalLinks   (menu action)
' Replaces linked formulas with their current values.
'------------------------------------------------------------------------------
Public Sub BreakExternalLinks()
    Dim wb As Workbook
    Dim sources As Variant
    Dim i As Long
    Dim broken As Long, failed As Long
    Dim rows As Collection

    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub

    sources = LinkList(wb)
    If Not IsArray(sources) Then
        MsgBox "'" & wb.Name & "' does not link to any other workbook.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    If Not Confirm("Break " & (UBound(sources) - LBound(sources) + 1) & " link(s) to other " & _
                   "workbooks?" & vbNewLine & vbNewLine & _
                   "Every formula that reads from another workbook is replaced by the value it " & _
                   "shows right now. The numbers stay, the connection goes.") Then Exit Sub
    If Not OfferBackup(wb) Then Exit Sub

    Set rows = NewReport()
    ReportRow rows, DOCTOR_NAME & " - Links broken"
    ReportRow rows, "Workbook", wb.Name
    ReportHeading rows, "Result"
    ReportRow rows, "Source", "Result"

    FastMode True
    For i = LBound(sources) To UBound(sources)
        On Error Resume Next
        wb.BreakLink Name:=CStr(sources(i)), Type:=xlLinkTypeExcelLinks
        If Err.Number = 0 Then
            broken = broken + 1
            ReportRow rows, CStr(sources(i)), "Broken"
        Else
            failed = failed + 1
            ReportRow rows, CStr(sources(i)), "Failed: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    FastMode False

    ShowReport rows, "Links broken"

    MsgBox broken & " link(s) broken." & _
           IIf(failed > 0, vbNewLine & failed & " could not be - see the report.", "") & _
           vbNewLine & vbNewLine & "Save the workbook to keep the change.", vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' LinkList
' The workbook's Excel links, or Empty when there are none.
'------------------------------------------------------------------------------
Public Function LinkList(ByVal wb As Workbook) As Variant
    Dim sources As Variant

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    sources = wb.LinkSources(xlLinkTypeExcelLinks)
    On Error GoTo 0

    If IsArray(sources) Then LinkList = sources
End Function

Public Function LinkCount(ByVal wb As Workbook) As Long
    Dim sources As Variant

    sources = LinkList(wb)
    If IsArray(sources) Then LinkCount = UBound(sources) - LBound(sources) + 1
End Function

'==============================================================================
' Private helpers
'==============================================================================
' Lists cells whose formula points at another workbook. Returns how many it
' listed, stopping at MAX_CELLS_LISTED so one bad sheet cannot fill the report.
Private Function ListLinkedCells(ByVal wb As Workbook, ByVal rows As Collection) As Long
    Dim ws As Worksheet
    Dim formulaCells As Range
    Dim cell As Range
    Dim listed As Long
    Dim formulaText As String

    FastMode True
    For Each ws In wb.Worksheets
        Set formulaCells = Nothing
        On Error Resume Next
        Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not formulaCells Is Nothing Then
            For Each cell In formulaCells
                formulaText = CStr(cell.Formula)
                If InStr(formulaText, "[") > 0 Then
                    listed = listed + 1
                    If listed > MAX_CELLS_LISTED Then Exit For
                    ReportRow rows, ws.Name, cell.Address(False, False), Left$(formulaText, 250)
                End If
            Next cell
        End If
        If listed > MAX_CELLS_LISTED Then Exit For
    Next ws
    FastMode False

    ListLinkedCells = listed
End Function
