Attribute VB_Name = "modDoctorRunner"
'==============================================================================
' modDoctorRunner - the one-click clean
'------------------------------------------------------------------------------
' Runs the low-risk half of every tool in one go, in the order that gets the
' most out of each step, and reports what it did.
'
' What it does NOT do, because each one needs a decision:
'   - delete names that merely look unused
'   - delete custom styles that are not duplicates of built-ins
'   - remove conditional formatting
'   - break external links
'
' Depends on: every other Doctor module, and \src
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' CleanEverything   (menu action)
'------------------------------------------------------------------------------
Public Sub CleanEverything()
    Dim wb As Workbook
    Dim rows As Collection
    Dim names As Collection
    Dim summary As Object
    Dim duplicates As Collection, otherStyles As Collection
    Dim deletedNames As Long, failedNames As Long
    Dim deletedStyles As Long, failedStyles As Long
    Dim removedShapes As Long, deletedSheets As Long
    Dim rowsGone As Long, colsGone As Long, rowsTotal As Long, colsTotal As Long
    Dim ws As Worksheet
    Dim emptyNames As Collection
    Dim item As Variant
    Dim cancelled As Boolean
    Dim ignored As Collection
    Dim done As Long

    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub

    If Not Confirm("Run the safe clean-up on '" & wb.Name & "'?" & vbNewLine & vbNewLine & _
                   "  1  delete broken, external and hidden defined names" & vbNewLine & _
                   "  2  delete duplicate cell styles such as 'Normal 2'" & vbNewLine & _
                   "  3  delete hidden and zero-size drawing objects" & vbNewLine & _
                   "  4  delete sheets that hold nothing at all" & vbNewLine & _
                   "  5  trim every sheet back to its real data" & vbNewLine & vbNewLine & _
                   "Names that merely look unused, other custom styles, conditional " & _
                   "formatting and external links are left alone - those need your judgement, " & _
                   "so run those tools separately.") Then Exit Sub

    If Not OfferBackup(wb) Then Exit Sub

    StartTimer
    Set rows = NewReport()
    ReportRow rows, DOCTOR_NAME & " - Clean-up"
    ReportRow rows, "Workbook", wb.Name
    ReportRow rows, "Run", Format$(Now, "dd mmm yyyy hh:nn")
    ReportHeading rows, "What changed"
    ReportRow rows, "Step", "Result"

    FastMode True

    ' 1. Names
    Set names = AnalyseNames(wb, summary)
    If names.count > 0 Then
        Set ignored = DeleteNames(names, False, deletedNames, failedNames)
    End If
    ReportRow rows, "Defined names deleted", deletedNames & _
              IIf(failedNames > 0, " (" & failedNames & " could not be deleted)", "")

    ' 2. Styles
    ClassifyStyles wb, duplicates, otherStyles
    deletedStyles = DeleteStyles(wb, duplicates, failedStyles, cancelled)
    ReportRow rows, "Duplicate styles deleted", deletedStyles & _
              IIf(failedStyles > 0, " (" & failedStyles & " in use by Excel)", "") & _
              IIf(cancelled, " - stopped early with Esc", "")
    ReportRow rows, "Other custom styles left", otherStyles.count, "Use Clean styles to review these"

    ' 3. Drawing objects
    removedShapes = DeleteJunkShapes(wb)
    ReportRow rows, "Invisible objects removed", removedShapes

    ' 4. Empty sheets
    Set emptyNames = New Collection
    For Each ws In wb.Worksheets
        If SheetIsEmpty(ws) Then emptyNames.Add ws.Name
    Next ws
    For Each item In emptyNames
        If wb.Worksheets.count > 1 Then
            DeleteSheet CStr(item), wb
            deletedSheets = deletedSheets + 1
        End If
    Next item
    ReportRow rows, "Empty sheets deleted", deletedSheets

    ' 5. Used range
    For Each ws In wb.Worksheets
        rowsGone = 0
        colsGone = 0
        done = done + 1
        ShowProgress done, wb.Worksheets.count, "Trimming " & ws.Name
        TrimSheet ws, rowsGone, colsGone
        rowsTotal = rowsTotal + rowsGone
        colsTotal = colsTotal + colsGone
    Next ws
    ReportRow rows, "Empty rows removed", rowsTotal
    ReportRow rows, "Empty columns removed", colsTotal

    FastMode False
    ClearStatusBar

    ReportHeading rows, "Still worth a look"
    ReportRow rows, "Names that look unused", DictGet(summary, KIND_UNUSED, 0), "Clean names, then opt in"
    ReportRow rows, "Broken names still referenced", DictGet(summary, KIND_BROKEN_USED, 0), "Fix the formulas"
    ReportRow rows, "Other custom styles", otherStyles.count, "Clean styles, 'all custom' option"
    ReportRow rows, "External links", LinkCount(wb), "External links tool"

    ShowReport rows, "Clean-up"

    MsgBox "Clean-up finished in " & ElapsedText & "." & vbNewLine & vbNewLine & _
           deletedNames & " name(s), " & Format$(deletedStyles, "#,##0") & " style(s), " & _
           removedShapes & " object(s) and " & deletedSheets & " sheet(s) removed." & vbNewLine & _
           Format$(rowsTotal, "#,##0") & " empty rows trimmed." & vbNewLine & vbNewLine & _
           "Save, close and reopen the workbook to see the file size settle.", _
           vbInformation, DOCTOR_NAME
End Sub
