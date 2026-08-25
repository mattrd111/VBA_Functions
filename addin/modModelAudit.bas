Attribute VB_Name = "modModelAudit"
'==============================================================================
' modModelAudit - the model integrity review
'------------------------------------------------------------------------------
' Read-only. Nothing here changes the workbook being audited; every result is
' written to a new workbook.
'
' Three actions:
'   AuditModel          the full review, worst findings first
'   ShowFormulaMap      a one-character-per-cell picture of the active sheet
'   SelectFlaggedCells  selects the suspect cells on the active sheet so you can
'                       see them in context
'
' Depends on: modAuditCore, modAuditFormula, modDoctorCommon, and \src
'==============================================================================
Option Explicit

Private Const MAP_MAX_ROWS As Long = 400
Private Const MAP_MAX_COLUMNS As Long = 150
Private Const MAX_ERRORS_LISTED As Long = 40

'------------------------------------------------------------------------------
' AuditModel   (menu action)
'------------------------------------------------------------------------------
Public Sub AuditModel()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim findings As Collection
    Dim sheetStats As Object
    Dim totals As Object
    Dim perSheet As Collection
    Dim rows As Collection
    Dim done As Long

    Set wb = TargetWorkbook(False)
    If wb Is Nothing Then Exit Sub

    If wb.Worksheets.count > 15 Then
        If MsgBox("'" & wb.Name & "' has " & wb.Worksheets.count & " sheets. The review reads " & _
                  "every formula on every one of them and may take a few minutes." & vbNewLine & _
                  vbNewLine & "Carry on?", vbYesNo + vbQuestion, DOCTOR_NAME) <> vbYes Then Exit Sub
    End If

    StartTimer
    FastMode True

    Set findings = NewFindings()
    Set totals = NewStats()
    Set perSheet = New Collection

    For Each ws In wb.Worksheets
        done = done + 1
        ShowProgress done, wb.Worksheets.count, "Reviewing " & ws.Name

        Set sheetStats = NewStats()
        ScanSheet ws, findings, sheetStats
        CheckErrorCells ws, findings, sheetStats

        perSheet.Add Array(ws.Name, sheetStats)
        AddInto totals, sheetStats
    Next ws

    FastMode False
    ClearStatusBar

    Set rows = BuildAuditReport(wb, findings, perSheet, totals)
    ShowReport rows, "Model audit"

    MsgBox SeverityCount(findings, SEV_HIGH) & " high, " & _
           SeverityCount(findings, SEV_MEDIUM) & " medium and " & _
           SeverityCount(findings, SEV_LOW) & " low finding(s) in " & ElapsedText & "." & _
           vbNewLine & vbNewLine & _
           "Nothing in the workbook was changed. The report is in a new workbook, and " & _
           "the cell references in it are clickable while '" & wb.Name & "' stays open.", _
           vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' SelectFlaggedCells   (menu action)
' Runs the consistency check on the active sheet and selects what it finds, so
' you can look at the suspects where they live. Changes nothing.
'------------------------------------------------------------------------------
Public Sub SelectFlaggedCells()
    Dim ws As Worksheet
    Dim flagged As Range

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open a workbook first.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    StartTimer
    FastMode True
    Set flagged = ScanSheetForFlags(ws)
    FastMode False
    ClearStatusBar

    If flagged Is Nothing Then
        MsgBox "Nothing on '" & ws.Name & "' looks out of place." & vbNewLine & vbNewLine & _
               "Checked in " & ElapsedText & ".", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    On Error Resume Next
    flagged.Select
    On Error GoTo 0

    MsgBox flagged.Cells.count & " suspect cell(s) selected on '" & ws.Name & "'." & vbNewLine & _
           vbNewLine & "Run Audit model for the reasons.", vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' ShowFormulaMap   (menu action)
' One character per cell, so the shape of a sheet - where the inputs are, where
' the calculations are, and where they break - is visible in a single screen.
'------------------------------------------------------------------------------
Public Sub ShowFormulaMap()
    Dim ws As Worksheet
    Dim block As Range
    Dim findings As Collection
    Dim stats As Object
    Dim flags As Object
    Dim item As Variant
    Dim values As Variant
    Dim map As Variant
    Dim mapRows As Long, mapCols As Long
    Dim firstRow As Long, firstCol As Long
    Dim r As Long, c As Long
    Dim truncated As Boolean
    Dim reportWs As Worksheet
    Dim book As Workbook

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open a workbook first.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    Set ws = ActiveSheet
    On Error Resume Next
    Set block = ws.UsedRange
    On Error GoTo 0
    If block Is Nothing Then Exit Sub

    firstRow = block.Row
    firstCol = block.column
    mapRows = block.rows.count
    mapCols = block.Columns.count

    If mapRows > MAP_MAX_ROWS Then mapRows = MAP_MAX_ROWS: truncated = True
    If mapCols > MAP_MAX_COLUMNS Then mapCols = MAP_MAX_COLUMNS: truncated = True
    If mapRows < 1 Or mapCols < 1 Then Exit Sub

    StartTimer
    FastMode True

    Set findings = NewFindings()
    Set stats = NewStats()
    ScanSheet ws, findings, stats

    Set flags = NewDictionary(True)
    For Each item In findings
        If IsConsistencyCheck(CStr(item(3))) Then
            If Not flags.Exists(CStr(item(2))) Then
                flags.Add CStr(item(2)), IIf(CStr(item(3)) = "Number typed into a calculated row", "#", "X")
            End If
        End If
    Next item

    values = ws.Range(ws.Cells(firstRow, firstCol), _
                      ws.Cells(firstRow + mapRows - 1, firstCol + mapCols - 1)).Formula

    ReDim map(1 To mapRows + 1, 1 To mapCols + 1)
    map(1, 1) = ws.Name
    For c = 1 To mapCols
        map(1, c + 1) = ColumnLetter(firstCol + c - 1)
    Next c

    For r = 1 To mapRows
        map(r + 1, 1) = firstRow + r - 1
        For c = 1 To mapCols
            map(r + 1, c + 1) = MapCharacter(values, r, c, mapRows, mapCols, flags, _
                                             CellAddress(firstRow + r - 1, firstCol + c - 1))
        Next c
    Next r

    FastMode False
    ClearStatusBar

    Set book = Application.Workbooks.Add
    Set reportWs = DumpToSheet(map, Left$("Map - " & ws.Name, 31), book, False)
    If reportWs Is Nothing Then Exit Sub

    FormatMap reportWs, mapRows + 1, mapCols + 1

    MsgBox "Map of '" & ws.Name & "' drawn in " & ElapsedText & "." & vbNewLine & vbNewLine & _
           "N  a number someone typed (an input)" & vbNewLine & _
           "F  a formula" & vbNewLine & _
           "X  a formula that differs from the rest of its row" & vbNewLine & _
           "#  a number sitting in a row of formulas" & vbNewLine & _
           "T  text     .  empty     !  an error" & _
           IIf(truncated, vbNewLine & vbNewLine & "The sheet was larger than the map; only the " & _
               "top left " & MAP_MAX_ROWS & " x " & MAP_MAX_COLUMNS & " is shown.", ""), _
           vbInformation, DOCTOR_NAME
End Sub

'==============================================================================
' Report
'==============================================================================
Private Function BuildAuditReport(ByVal wb As Workbook, ByVal findings As Collection, _
                                  ByVal perSheet As Collection, ByVal totals As Object) As Collection
    Dim rows As Collection
    Dim item As Variant
    Dim entry As Variant
    Dim stats As Object
    Dim severities As Variant
    Dim s As Long
    Dim shown As Long

    Set rows = NewReport()
    ReportRow rows, DOCTOR_NAME & " - Model audit"
    ReportRow rows, "Workbook", wb.Name
    ReportRow rows, "Folder", IIf(Len(wb.Path) = 0, "(never saved)", wb.Path)
    ReportRow rows, "Checked", Format$(Now, "dd mmm yyyy hh:nn")
    ReportRow rows, "Nothing was changed", "This review is read-only."

    ReportHeading rows, "Verdict"
    ReportRow rows, "High", SeverityCount(findings, SEV_HIGH), "Look at these before you trust a number"
    ReportRow rows, "Medium", SeverityCount(findings, SEV_MEDIUM), "Worth a minute each"
    ReportRow rows, "Low", SeverityCount(findings, SEV_LOW), "Housekeeping and performance"
    ReportRow rows, "Formulas", Format$(DictGet(totals, "Formulas", 0), "#,##0")
    ReportRow rows, "Distinct formulas", Format$(DictGet(totals, "Unique formulas", 0), "#,##0"), _
              "A model with thousands of distinct formulas is a model nobody can check"

    If Application.Iteration Then
        ReportRow rows, "Iterative calculation", "ON", _
                  "Circular references are being resolved silently. Find out which ones."
    End If

    severities = Array(SEV_HIGH, SEV_MEDIUM, SEV_LOW)
    For s = LBound(severities) To UBound(severities)
        shown = SeverityCount(findings, CStr(severities(s)))
        If shown > 0 Then
            ReportHeading rows, CStr(severities(s)) & " (" & shown & ")"
            ReportRow rows, "Sheet", "Cell", "Go to", "Check", "What is wrong", "Formula"
            For Each item In findings
                If StrComp(CStr(item(0)), CStr(severities(s)), vbTextCompare) = 0 Then
                    ReportRow rows, item(1), item(2), GoToFormula(wb.Name, CStr(item(1)), CStr(item(2))), _
                              item(3), item(4), "'" & CStr(item(5))
                End If
            Next item
        End If
    Next s

    If findings.count = 0 Then
        ReportHeading rows, "Nothing found"
        ReportRow rows, "Result", "Every row of every block calculates consistently, with no " & _
                  "numbers buried in formulas."
    End If

    ReportHeading rows, "By sheet"
    ReportRow rows, "Sheet", "Formulas", "Distinct", "Inputs", "Hardcoded numbers", _
              "Volatile calls", "Whole-column refs", "External refs", "IFERROR", "Long formulas", "Error cells"

    For Each entry In perSheet
        Set stats = entry(1)
        ReportRow rows, entry(0), _
                  DictGet(stats, "Formulas", 0), _
                  DictGet(stats, "Unique formulas", 0), _
                  DictGet(stats, "Numbers", 0), _
                  DictGet(stats, "Hardcoded numbers", 0), _
                  DictGet(stats, "Volatile calls", 0), _
                  DictGet(stats, "Whole-column references", 0), _
                  DictGet(stats, "External references", 0), _
                  DictGet(stats, "IFERROR wrappers", 0), _
                  DictGet(stats, "Long formulas", 0), _
                  DictGet(stats, "Error cells", 0)
    Next entry

    ReportHeading rows, "How to read this"
    ReportRow rows, "Formula differs from the row", _
              "The other cells in that run share one formula and this one does not. " & _
              "Sometimes deliberate, usually not."
    ReportRow rows, "Number typed into a calculated row", _
              "Someone replaced a formula with a value. This is the one that costs money."
    ReportRow rows, "Number inside a formula", _
              "An assumption buried where nobody will find it to update. Indexes and " & _
              "counts inside lookup and rounding functions are ignored."
    ReportRow rows, "Limits", _
              "Runs shorter than 4 cells are not judged, and no single check reports more " & _
              "than 150 findings per sheet."

    Set BuildAuditReport = rows
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Sub CheckErrorCells(ByVal ws As Worksheet, ByVal findings As Collection, ByVal stats As Object)
    Dim errorCells As Range
    Dim area As Range
    Dim cell As Range
    Dim total As Long
    Dim listed As Long
    Dim text As String

    On Error Resume Next
    Set errorCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
    On Error GoTo 0
    If errorCells Is Nothing Then Exit Sub

    For Each area In errorCells.Areas
        total = total + area.Cells.count
    Next area
    Bump stats, "Error cells", total

    For Each cell In errorCells
        listed = listed + 1
        If listed > MAX_ERRORS_LISTED Then Exit For

        On Error Resume Next
        text = CStr(cell.text)
        On Error GoTo 0

        If InStr(1, text, "#N/A", vbTextCompare) > 0 Then
            AddFinding findings, SEV_LOW, ws.Name, cell.Address(False, False), "Error showing", _
                       "#N/A - often a lookup that is meant to miss, but check.", CStr(cell.Formula)
        Else
            AddFinding findings, SEV_HIGH, ws.Name, cell.Address(False, False), "Error showing", _
                       text & " - this cell is broken.", CStr(cell.Formula)
        End If
    Next cell
End Sub

Private Function SeverityCount(ByVal findings As Collection, ByVal severity As String) As Long
    Dim item As Variant
    Dim total As Long

    For Each item In findings
        If StrComp(CStr(item(0)), severity, vbTextCompare) = 0 Then total = total + 1
    Next item
    SeverityCount = total
End Function

Private Sub AddInto(ByVal totals As Object, ByVal stats As Object)
    Dim key As Variant

    If totals Is Nothing Or stats Is Nothing Then Exit Sub
    For Each key In stats.keys
        DictIncrement totals, CStr(key), CDbl(stats(key))
    Next key
End Sub

' A clickable jump back to the audited workbook, which works while that
' workbook stays open.
Private Function GoToFormula(ByVal workbookName As String, ByVal sheetName As String, _
                             ByVal address As String) As String
    Dim link As String

    If InStr(address, ":") > 0 Then address = Split(address, ":")(0)
    link = "[" & workbookName & "]'" & Replace$(sheetName, "'", "''") & "'!" & address
    GoToFormula = "=HYPERLINK(""" & link & """,""" & address & """)"
End Function

Private Function MapCharacter(ByRef values As Variant, ByVal r As Long, ByVal c As Long, _
                              ByVal mapRows As Long, ByVal mapCols As Long, _
                              ByVal flags As Object, ByVal address As String) As String
    Dim value As Variant
    Dim text As String

    If flags.Exists(address) Then MapCharacter = CStr(flags(address)): Exit Function

    If mapRows = 1 And mapCols = 1 Then
        value = values
    Else
        value = values(r, c)
    End If

    If IsError(value) Then MapCharacter = "!": Exit Function
    If IsEmpty(value) Or IsNull(value) Then MapCharacter = ".": Exit Function

    If VarType(value) = vbString Then
        text = CStr(value)
        If Len(text) = 0 Then MapCharacter = ".": Exit Function
        If Left$(text, 1) = "=" Then MapCharacter = "F": Exit Function
        If Left$(text, 1) = "#" Then MapCharacter = "!": Exit Function
        If IsNumeric(text) Then MapCharacter = "N": Exit Function
        MapCharacter = "T"
        Exit Function
    End If

    If IsNumeric(value) Then MapCharacter = "N" Else MapCharacter = "T"
End Function

Private Sub FormatMap(ByVal ws As Worksheet, ByVal rowCount As Long, ByVal columnCount As Long)
    Dim body As Range

    On Error Resume Next

    With ws.Range("A1").Resize(rowCount, columnCount)
        .Font.Name = "Consolas"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With

    ws.rows(1).Font.Bold = True
    ws.Columns(1).Font.Bold = True
    ws.Columns(1).HorizontalAlignment = xlRight
    ws.Columns(1).ColumnWidth = 6

    If columnCount > 1 Then
        ws.Range(ws.Columns(2), ws.Columns(columnCount)).ColumnWidth = 2.3
    End If

    If rowCount > 1 And columnCount > 1 Then
        Set body = ws.Range(ws.Cells(2, 2), ws.Cells(rowCount, columnCount))
        body.FormatConditions.Delete
        PaintWhen body, "X", RGB(255, 199, 206), RGB(156, 0, 6)
        PaintWhen body, "#", RGB(255, 235, 156), RGB(156, 87, 0)
        PaintWhen body, "!", RGB(156, 0, 6), RGB(255, 255, 255)
        PaintWhen body, "N", RGB(221, 235, 247), RGB(0, 51, 153)
        PaintWhen body, "F", RGB(242, 242, 242), RGB(89, 89, 89)
        PaintWhen body, "T", RGB(255, 255, 255), RGB(166, 166, 166)
        PaintWhen body, ".", RGB(255, 255, 255), RGB(217, 217, 217)
    End If

    FreezeHeader ws, 1, 1
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Private Sub PaintWhen(ByVal target As Range, ByVal character As String, _
                      ByVal fillColour As Long, ByVal fontColour As Long)
    Dim rule As FormatCondition

    On Error Resume Next
    Set rule = target.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, _
                                           Formula1:="=""" & character & """")
    If rule Is Nothing Then Exit Sub
    rule.Interior.Color = fillColour
    rule.Font.Color = fontColour
    On Error GoTo 0
End Sub
