Attribute VB_Name = "modDoctorAudit"
'==============================================================================
' modDoctorAudit - the read-only health check
'------------------------------------------------------------------------------
' Changes nothing. Reports what is in the workbook, what is making it heavy,
' and which of the other tools would help.
'
' Depends on: modDoctorCommon, modDoctorNames, modDoctorScan, modDoctorStyles,
'             modDoctorSheets, modDoctorLinks, and \src
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' AuditWorkbook   (menu action)
'------------------------------------------------------------------------------
Public Sub AuditWorkbook()
    Dim wb As Workbook
    Dim rows As Collection
    Dim findings As Collection
    Dim names As Collection
    Dim summary As Object
    Dim duplicates As Collection, otherStyles As Collection
    Dim builtInStyles As Long
    Dim wastedRows As Long, junkShapes As Long, emptySheets As Long, cfRules As Long

    Set wb = TargetWorkbook(False)
    If wb Is Nothing Then Exit Sub

    If wb.Worksheets.count > 25 Then
        If MsgBox("'" & wb.Name & "' has " & wb.Worksheets.count & " sheets. The audit reads " & _
                  "every formula in the workbook and may take a minute or two." & vbNewLine & _
                  vbNewLine & "Carry on?", vbYesNo + vbQuestion, DOCTOR_NAME) <> vbYes Then Exit Sub
    End If

    StartTimer
    FastMode True

    Set rows = NewReport()
    Set findings = New Collection

    AddWorkbookSection wb, rows
    AddSheetsSection wb, rows, wastedRows, junkShapes, emptySheets, cfRules

    Set names = AnalyseNames(wb, summary)
    AddNamesSection wb, rows, summary

    builtInStyles = ClassifyStyles(wb, duplicates, otherStyles)
    AddStylesSection wb, rows, builtInStyles, duplicates.count, otherStyles.count

    AddLinksSection wb, rows
    AddOtherSection wb, rows

    BuildFindings wb, findings, summary, duplicates.count, otherStyles.count, _
                  wastedRows, junkShapes, emptySheets, cfRules

    FastMode False
    ClearStatusBar

    ShowReport MergeFindings(rows, findings, wb), "Workbook audit"
End Sub

'==============================================================================
' Sections
'==============================================================================
Private Sub AddWorkbookSection(ByVal wb As Workbook, ByVal rows As Collection)
    Dim sizeBytes As Double

    ReportHeading rows, "Workbook"
    If Len(wb.Path) > 0 Then
        sizeBytes = FileSizeBytes(wb.FullName)
        ReportRow rows, "File size", SizeText(sizeBytes), "As last saved"
    End If
    ReportRow rows, "Sheets", wb.Worksheets.count
    ReportRow rows, "Chart sheets", ChartSheetCount(wb)
    ReportRow rows, "Read only", IIf(wb.ReadOnly, "Yes", "No")
    ReportRow rows, "Structure protected", IIf(wb.ProtectStructure, "Yes", "No")
    ReportRow rows, "Calculation", CalculationText()
    ReportRow rows, "Format", FileExtension(wb.Name)
End Sub

Private Sub AddSheetsSection(ByVal wb As Workbook, ByVal rows As Collection, _
                             ByRef wastedRows As Long, ByRef junkShapes As Long, _
                             ByRef emptySheets As Long, ByRef cfRules As Long)
    Dim ws As Worksheet
    Dim done As Long
    Dim usedLastRow As Long, realLastRow As Long, wasted As Long
    Dim shapes As Long, junk As Long, rules As Long

    ReportHeading rows, "Sheets"
    ReportRow rows, "Sheet", "Visible", "Used range", "Real data", "Wasted rows", _
              "Formulas", "CF rules", "Validation", "Shapes", "Junk shapes", _
              "Comments", "Links", "Merged"

    For Each ws In wb.Worksheets
        done = done + 1
        ShowProgress done, wb.Worksheets.count, "Auditing " & ws.Name

        usedLastRow = UsedLastRow(ws)
        realLastRow = LastRow(ws)
        wasted = usedLastRow - realLastRow
        If wasted < 0 Then wasted = 0
        wastedRows = wastedRows + wasted

        shapes = ShapeCount(ws)
        junk = CountJunkShapes(ws)
        junkShapes = junkShapes + junk

        rules = RuleCount(ws)
        cfRules = cfRules + rules

        If SheetIsEmpty(ws) Then emptySheets = emptySheets + 1

        ReportRow rows, ws.Name, VisibilityText(ws), UsedAddress(ws), DataAddress(ws), wasted, _
                  FormulaCount(ws), rules, ValidationCount(ws), shapes, junk, _
                  CommentCount(ws), HyperlinkCount(ws), MergedText(ws)
    Next ws
End Sub

Private Sub AddNamesSection(ByVal wb As Workbook, ByVal rows As Collection, ByVal summary As Object)
    ReportHeading rows, "Defined names"
    ReportRow rows, "Total", wb.names.count
    ReportRow rows, "In use", DictGet(summary, KIND_IN_USE, 0)
    ReportRow rows, "Reserved", DictGet(summary, KIND_RESERVED, 0), "Print areas, filters - kept"
    ReportRow rows, "Broken (#REF!)", DictGet(summary, KIND_BROKEN, 0), "Safe to delete"
    ReportRow rows, "Broken but referenced", DictGet(summary, KIND_BROKEN_USED, 0), "Fix the formula that uses it"
    ReportRow rows, "External", DictGet(summary, KIND_EXTERNAL, 0), "Left behind by copied sheets"
    ReportRow rows, "Hidden", DictGet(summary, KIND_HIDDEN, 0), "Usually junk"
    ReportRow rows, "Unused", DictGet(summary, KIND_UNUSED, 0), "Nothing appears to refer to it"
End Sub

Private Sub AddStylesSection(ByVal wb As Workbook, ByVal rows As Collection, _
                             ByVal builtInStyles As Long, ByVal duplicateStyles As Long, _
                             ByVal otherStyles As Long)
    ReportHeading rows, "Cell styles"
    ReportRow rows, "Total", wb.Styles.count
    ReportRow rows, "Built in", builtInStyles, "Excel's own"
    ReportRow rows, "Duplicates of built-ins", duplicateStyles, "'Normal 2' and friends - safe to delete"
    ReportRow rows, "Other custom", otherStyles
End Sub

Private Sub AddLinksSection(ByVal wb As Workbook, ByVal rows As Collection)
    Dim sources As Variant
    Dim i As Long

    ReportHeading rows, "External links"
    sources = LinkList(wb)

    If Not IsArray(sources) Then
        ReportRow rows, "None"
        Exit Sub
    End If

    ReportRow rows, "Source", "File found"
    For i = LBound(sources) To UBound(sources)
        ReportRow rows, CStr(sources(i)), IIf(FileExists(CStr(sources(i))), "Yes", "No - dead link")
    Next i
End Sub

Private Sub AddOtherSection(ByVal wb As Workbook, ByVal rows As Collection)
    ReportHeading rows, "Other"
    ReportRow rows, "Pivot caches", SafeCount(wb, "pivotcaches")
    ReportRow rows, "Connections", SafeCount(wb, "connections")
    ReportRow rows, "Custom views", SafeCount(wb, "customviews")
    ReportRow rows, "Note", "Cell styles are workbook-wide, so a style added on one sheet weighs on every sheet."
End Sub

'==============================================================================
' Findings - the "so what" list
'==============================================================================
Private Sub BuildFindings(ByVal wb As Workbook, ByVal findings As Collection, _
                          ByVal summary As Object, ByVal duplicateStyles As Long, _
                          ByVal otherStyles As Long, ByVal wastedRows As Long, _
                          ByVal junkShapes As Long, ByVal emptySheets As Long, _
                          ByVal cfRules As Long)
    Dim deletableNames As Long
    Dim deadLinks As Long
    Dim sizeBytes As Double

    deletableNames = DictGet(summary, KIND_BROKEN, 0) + DictGet(summary, KIND_EXTERNAL, 0) + _
                     DictGet(summary, KIND_HIDDEN, 0)

    If duplicateStyles > 50 Then
        findings.Add Array("Clean styles", Format$(duplicateStyles, "#,##0") & " duplicate styles such as 'Normal 2'", "High")
    ElseIf duplicateStyles > 0 Then
        findings.Add Array("Clean styles", Format$(duplicateStyles, "#,##0") & " duplicate styles", "Low")
    End If

    If otherStyles > 500 Then
        findings.Add Array("Clean styles", Format$(otherStyles, "#,##0") & " other custom styles - consider the 'all custom' option", "Medium")
    End If

    If deletableNames > 0 Then
        findings.Add Array("Clean names", deletableNames & " broken, external or hidden name(s)", _
                           IIf(deletableNames > 20, "High", "Medium"))
    End If

    If DictGet(summary, KIND_BROKEN_USED, 0) > 0 Then
        findings.Add Array("Fix by hand", DictGet(summary, KIND_BROKEN_USED, 0) & _
                           " name(s) refer to #REF! and are still used by formulas", "High")
    End If

    If wastedRows > 5000 Then
        findings.Add Array("Reset used range", Format$(wastedRows, "#,##0") & _
                           " empty rows are still counted as used", "High")
    ElseIf wastedRows > 0 Then
        findings.Add Array("Reset used range", Format$(wastedRows, "#,##0") & " empty rows counted as used", "Low")
    End If

    If junkShapes > 0 Then
        findings.Add Array("Remove invisible objects", junkShapes & " hidden or zero-size drawing object(s)", _
                           IIf(junkShapes > 50, "High", "Low"))
    End If

    If emptySheets > 0 Then
        findings.Add Array("Delete empty sheets", emptySheets & " sheet(s) hold nothing at all", "Low")
    End If

    If cfRules > 500 Then
        findings.Add Array("Conditional formatting", Format$(cfRules, "#,##0") & _
                           " rules - copy and paste has multiplied them", "Medium")
    End If

    deadLinks = DeadLinkCount(wb)
    If deadLinks > 0 Then
        findings.Add Array("External links", deadLinks & " link(s) point at files that are not there", "High")
    End If

    If Len(wb.Path) > 0 Then
        sizeBytes = FileSizeBytes(wb.FullName)
        If sizeBytes > 20000000# Then
            findings.Add Array("File size", SizeText(sizeBytes) & " - work through the findings above, " & _
                               "then save as .xlsb for a smaller file", "Medium")
        End If
    End If

    If Application.Calculation <> xlCalculationAutomatic Then
        findings.Add Array("Calculation", "Calculation is not set to automatic in this Excel session", "Low")
    End If
End Sub

Private Function MergeFindings(ByVal rows As Collection, ByVal findings As Collection, _
                               ByVal wb As Workbook) As Collection
    Dim merged As Collection
    Dim item As Variant
    Dim i As Long

    Set merged = NewReport()

    ' Findings go at the top, where they will be read.
    merged.Add Array(DOCTOR_NAME & " - Workbook audit")
    merged.Add Array("Workbook", wb.Name)
    merged.Add Array("Folder", IIf(Len(wb.Path) = 0, "(never saved)", wb.Path))
    merged.Add Array("Checked", Format$(Now, "dd mmm yyyy hh:nn"))
    merged.Add Array("")
    merged.Add Array("What is worth doing")

    If findings.count = 0 Then
        merged.Add Array("Nothing - this workbook is in good shape.")
    Else
        merged.Add Array("Priority", "Tool", "Why")
        For Each item In findings
            merged.Add Array(item(2), item(0), item(1))
        Next item
    End If

    ' Then the detail.
    For i = 1 To rows.count
        merged.Add rows(i)
    Next i

    Set MergeFindings = merged
End Function

'==============================================================================
' Private measurement helpers - each one absorbs its own errors, because an
' audit must never fall over on a sheet it cannot read.
'==============================================================================
Private Function UsedLastRow(ByVal ws As Worksheet) As Long
    On Error Resume Next
    UsedLastRow = ws.UsedRange.Row + ws.UsedRange.rows.count - 1
    If Err.Number <> 0 Then UsedLastRow = 0: Err.Clear
    On Error GoTo 0
End Function

Private Function UsedAddress(ByVal ws As Worksheet) As String
    On Error Resume Next
    UsedAddress = ws.UsedRange.Address(False, False)
    If Err.Number <> 0 Then UsedAddress = "?": Err.Clear
    On Error GoTo 0
End Function

Private Function DataAddress(ByVal ws As Worksheet) As String
    Dim block As Range

    On Error Resume Next
    Set block = DataRange(ws)
    On Error GoTo 0
    If block Is Nothing Then DataAddress = "(empty)" Else DataAddress = block.Address(False, False)
End Function

Private Function FormulaCount(ByVal ws As Worksheet) As Long
    Dim formulaCells As Range
    Dim area As Range
    Dim total As Long

    On Error Resume Next
    Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If formulaCells Is Nothing Then Exit Function

    For Each area In formulaCells.Areas
        total = total + area.Cells.count
    Next area
    FormulaCount = total
End Function

Private Function ValidationCount(ByVal ws As Worksheet) As Long
    Dim cells As Range
    Dim area As Range
    Dim total As Long

    On Error Resume Next
    Set cells = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    On Error GoTo 0
    If cells Is Nothing Then Exit Function

    For Each area In cells.Areas
        total = total + area.Cells.count
    Next area
    ValidationCount = total
End Function

Private Function ShapeCount(ByVal ws As Worksheet) As Long
    On Error Resume Next
    ShapeCount = ws.Shapes.count
    If Err.Number <> 0 Then ShapeCount = 0: Err.Clear
    On Error GoTo 0
End Function

Private Function CommentCount(ByVal ws As Worksheet) As Long
    Dim total As Long

    On Error Resume Next
    total = ws.Comments.count
    If Err.Number <> 0 Then Err.Clear
    total = total + ws.CommentsThreaded.count          ' newer Excel only
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    CommentCount = total
End Function

Private Function HyperlinkCount(ByVal ws As Worksheet) As Long
    On Error Resume Next
    HyperlinkCount = ws.Hyperlinks.count
    If Err.Number <> 0 Then HyperlinkCount = 0: Err.Clear
    On Error GoTo 0
End Function

Private Function MergedText(ByVal ws As Worksheet) As String
    Dim state As Variant

    On Error Resume Next
    state = ws.UsedRange.MergeCells
    If Err.Number <> 0 Then MergedText = "?": Err.Clear: Exit Function
    On Error GoTo 0

    If IsNull(state) Then
        MergedText = "Some"
    ElseIf state Then
        MergedText = "All"
    Else
        MergedText = "None"
    End If
End Function

Private Function VisibilityText(ByVal ws As Worksheet) As String
    Select Case ws.visible
        Case xlSheetVisible: VisibilityText = "Visible"
        Case xlSheetHidden: VisibilityText = "Hidden"
        Case xlSheetVeryHidden: VisibilityText = "Very hidden"
        Case Else: VisibilityText = "?"
    End Select
End Function

Private Function ChartSheetCount(ByVal wb As Workbook) As Long
    On Error Resume Next
    ChartSheetCount = wb.Charts.count
    If Err.Number <> 0 Then ChartSheetCount = 0: Err.Clear
    On Error GoTo 0
End Function

Private Function DeadLinkCount(ByVal wb As Workbook) As Long
    Dim sources As Variant
    Dim i As Long
    Dim total As Long

    sources = LinkList(wb)
    If Not IsArray(sources) Then Exit Function

    For i = LBound(sources) To UBound(sources)
        If Not FileExists(CStr(sources(i))) Then total = total + 1
    Next i
    DeadLinkCount = total
End Function

Private Function SafeCount(ByVal wb As Workbook, ByVal collectionName As String) As Variant
    Dim total As Long

    On Error Resume Next
    Select Case LCase$(collectionName)
        Case "pivotcaches": total = wb.PivotCaches.count
        Case "connections": total = wb.Connections.count
        Case "customviews": total = wb.CustomViews.count
    End Select

    If Err.Number <> 0 Then
        Err.Clear
        SafeCount = "n/a"
    Else
        SafeCount = total
    End If
    On Error GoTo 0
End Function

Private Function CalculationText() As String
    On Error Resume Next
    Select Case Application.Calculation
        Case xlCalculationAutomatic: CalculationText = "Automatic"
        Case xlCalculationManual: CalculationText = "Manual"
        Case xlCalculationSemiautomatic: CalculationText = "Automatic except tables"
        Case Else: CalculationText = "?"
    End Select
    If Err.Number <> 0 Then CalculationText = "?": Err.Clear
    On Error GoTo 0
End Function

Private Function SizeText(ByVal bytes As Double) As String
    If bytes <= 0 Then SizeText = "?": Exit Function

    If bytes < 1024# Then
        SizeText = Format$(bytes, "#,##0") & " bytes"
    ElseIf bytes < 1048576# Then
        SizeText = Format$(bytes / 1024#, "#,##0.0") & " KB"
    Else
        SizeText = Format$(bytes / 1048576#, "#,##0.0") & " MB"
    End If
End Function
