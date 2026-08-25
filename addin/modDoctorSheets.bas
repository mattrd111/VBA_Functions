Attribute VB_Name = "modDoctorSheets"
'==============================================================================
' modDoctorSheets - sheet level clean-up
'------------------------------------------------------------------------------
' Resetting the used range, removing empty sheets, unhiding what was hidden,
' clearing stray drawing objects and pruning conditional formatting that no
' longer applies to anything.
'
' Depends on: modDoctorCommon, modRange, modWorkbook, modApp (from \src)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' ResetUsedRange   (menu action)
' The classic fix for a workbook whose scroll bar runs to row 1,048,576: delete
' every row below and column right of the real data, so the used range - and
' the file size - collapse back to what is actually there.
'
' Cells occupied by a shape, chart or picture are protected, so nothing anchored
' below the data is deleted with it.
'------------------------------------------------------------------------------
Public Sub ResetUsedRange()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rows As Collection
    Dim rowsGone As Long, colsGone As Long
    Dim totalRows As Long, totalCols As Long
    Dim done As Long
    Dim before As String

    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub

    If Not Confirm("Trim every sheet in '" & wb.Name & "' back to its real data?" & vbNewLine & _
                   vbNewLine & "Formatting, conditional formatting and validation applied to the " & _
                   "empty rows and columns beyond the data will go with them.") Then Exit Sub
    If Not OfferBackup(wb) Then Exit Sub

    Set rows = NewReport()
    ReportRow rows, DOCTOR_NAME & " - Used range reset"
    ReportRow rows, "Workbook", wb.Name
    ReportHeading rows, "Sheets"
    ReportRow rows, "Sheet", "Used range before", "Used range after", "Rows removed", "Columns removed"

    StartTimer
    FastMode True

    For Each ws In wb.Worksheets
        done = done + 1
        ShowProgress done, wb.Worksheets.count, "Trimming " & ws.Name

        rowsGone = 0
        colsGone = 0
        before = UsedRangeAddress(ws)
        TrimSheet ws, rowsGone, colsGone
        totalRows = totalRows + rowsGone
        totalCols = totalCols + colsGone

        ReportRow rows, ws.Name, before, UsedRangeAddress(ws), rowsGone, colsGone
    Next ws

    FastMode False
    ClearStatusBar

    ReportHeading rows, "Total"
    ReportRow rows, "Rows removed", totalRows
    ReportRow rows, "Columns removed", totalCols

    ShowReport rows, "Used range reset"

    MsgBox Format$(totalRows, "#,##0") & " row(s) and " & Format$(totalCols, "#,##0") & _
           " column(s) removed in " & ElapsedText & "." & vbNewLine & vbNewLine & _
           "Save, close and reopen the workbook for the last cell to settle.", _
           vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' DeleteEmptySheets   (menu action)
' Removes sheets with no cells, no shapes, no tables and no pivots.
'------------------------------------------------------------------------------
Public Sub DeleteEmptySheets()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim empties As Collection
    Dim item As Variant
    Dim list As String

    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub

    Set empties = New Collection
    For Each ws In wb.Worksheets
        If SheetIsEmpty(ws) Then empties.Add ws.Name
    Next ws

    If empties.count = 0 Then
        MsgBox "Every sheet in '" & wb.Name & "' has something on it.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    For Each item In empties
        list = list & vbNewLine & "  " & CStr(item)
    Next item

    If Not Confirm("Delete " & empties.count & " empty sheet(s) from '" & wb.Name & "'?" & _
                   vbNewLine & list) Then Exit Sub

    FastMode True
    For Each item In empties
        DeleteSheet CStr(item), wb
    Next item
    FastMode False

    MsgBox "Done. '" & wb.Name & "' now has " & wb.Worksheets.count & " sheet(s)." & vbNewLine & _
           "The last visible sheet is never deleted, so one may remain.", vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' UnhideAllSheets   (menu action)
' Including the very hidden ones, which the Unhide dialog will not show.
'------------------------------------------------------------------------------
Public Sub UnhideAllSheets()
    Dim wb As Workbook
    Dim sh As Object
    Dim hidden As Long, veryHidden As Long
    Dim failures As Long

    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub

    For Each sh In wb.Sheets
        If sh.Visible = xlSheetVeryHidden Then
            veryHidden = veryHidden + 1
        ElseIf sh.Visible = xlSheetHidden Then
            hidden = hidden + 1
        End If
    Next sh

    If hidden + veryHidden = 0 Then
        MsgBox "Nothing is hidden in '" & wb.Name & "'.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    If MsgBox("Unhide " & hidden & " hidden and " & veryHidden & " very hidden sheet(s)?", _
              vbYesNo + vbQuestion, DOCTOR_NAME) <> vbYes Then Exit Sub

    On Error Resume Next
    For Each sh In wb.Sheets
        If sh.Visible <> xlSheetVisible Then
            sh.Visible = xlSheetVisible
            If Err.Number <> 0 Then failures = failures + 1: Err.Clear
        End If
    Next sh
    On Error GoTo 0

    MsgBox (hidden + veryHidden - failures) & " sheet(s) unhidden." & _
           IIf(failures > 0, vbNewLine & failures & " could not be - the workbook structure is protected.", ""), _
           vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' RemoveInvisibleShapes   (menu action)
' Bad pastes and old imports leave behind drawing objects with no size or with
' visibility switched off. A sheet can end up with thousands.
'------------------------------------------------------------------------------
Public Sub RemoveInvisibleShapes()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim doomed As Collection
    Dim rows As Collection
    Dim item As Variant
    Dim i As Long
    Dim removed As Long
    Dim answer As VbMsgBoxResult

    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub

    Set doomed = New Collection
    On Error Resume Next
    For Each ws In wb.Worksheets
        For i = ws.Shapes.count To 1 Step -1
            Set shp = ws.Shapes(i)
            If Err.Number <> 0 Then
                Err.Clear
            ElseIf ShapeIsJunk(shp) Then
                doomed.Add Array(ws.Name, shp.Name, ShapeTypeText(shp), JunkReason(shp), shp)
            End If
        Next i
    Next ws
    On Error GoTo 0

    If doomed.count = 0 Then
        MsgBox "No invisible or zero-size drawing objects found in '" & wb.Name & "'.", _
               vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    Set rows = NewReport()
    ReportRow rows, DOCTOR_NAME & " - Invisible drawing objects"
    ReportRow rows, "Workbook", wb.Name
    ReportHeading rows, "Found (" & doomed.count & ")"
    ReportRow rows, "Sheet", "Shape", "Type", "Why", "Deleted"

    answer = MsgBox(doomed.count & " invisible or zero-size drawing object(s) found in '" & _
                    wb.Name & "'." & vbNewLine & vbNewLine & _
                    "Yes     delete them" & vbNewLine & _
                    "No      list them without deleting" & vbNewLine & _
                    "Cancel  stop" & vbNewLine & vbNewLine & _
                    "A control your macros hide on purpose would also be removed, so list " & _
                    "them first if you are not sure. This cannot be undone with Ctrl+Z.", _
                    vbYesNoCancel + vbExclamation, DOCTOR_NAME)
    If answer = vbCancel Then Exit Sub

    If answer = vbYes Then
        If Not OfferBackup(wb) Then Exit Sub
        FastMode True
    End If

    On Error Resume Next
    For Each item In doomed
        If answer = vbYes Then
            item(4).Delete
            If Err.Number = 0 Then
                removed = removed + 1
                ReportRow rows, item(0), item(1), item(2), item(3), "Yes"
            Else
                Err.Clear
                ReportRow rows, item(0), item(1), item(2), item(3), "Failed"
            End If
        Else
            ReportRow rows, item(0), item(1), item(2), item(3), "No"
        End If
    Next item
    On Error GoTo 0
    If answer = vbYes Then FastMode False

    ShowReport rows, "Drawing objects"

    If answer = vbYes Then
        MsgBox removed & " drawing object(s) removed.", vbInformation, DOCTOR_NAME
    End If
End Sub

'------------------------------------------------------------------------------
' CleanConditionalFormats   (menu action)
' Copy and paste multiplies conditional formatting rules until a sheet carries
' thousands of near-identical ones. This removes the rules that no longer apply
' to any data at all, and reports what is left.
'------------------------------------------------------------------------------
Public Sub CleanConditionalFormats()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rows As Collection
    Dim before As Long, emptyRules As Long
    Dim totalBefore As Long, totalEmpty As Long, removed As Long
    Dim done As Long

    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub

    Set rows = NewReport()
    ReportRow rows, DOCTOR_NAME & " - Conditional formatting"
    ReportRow rows, "Workbook", wb.Name
    ReportHeading rows, "Rules per sheet"
    ReportRow rows, "Sheet", "Rules", "Applying to nothing"

    FastMode True
    For Each ws In wb.Worksheets
        done = done + 1
        ShowProgress done, wb.Worksheets.count, "Reading rules on " & ws.Name
        before = RuleCount(ws)
        emptyRules = EmptyRuleCount(ws)
        totalBefore = totalBefore + before
        totalEmpty = totalEmpty + emptyRules
        If before > 0 Then ReportRow rows, ws.Name, before, emptyRules
    Next ws
    FastMode False
    ClearStatusBar

    If totalBefore = 0 Then
        MsgBox "'" & wb.Name & "' has no conditional formatting.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    If totalEmpty = 0 Then
        ReportHeading rows, "Nothing to remove"
        ReportRow rows, "All " & totalBefore & " rule(s) still apply to cells that hold data."
        ShowReport rows, "Conditional formatting"
        MsgBox "All " & totalBefore & " rule(s) still apply to data, so none were removed." & _
               vbNewLine & "A sheet with hundreds of rules is usually worth rebuilding by hand.", _
               vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    If Not Confirm("'" & wb.Name & "' has " & totalBefore & " conditional formatting rule(s), " & _
                   "of which " & totalEmpty & " apply only to empty cells." & vbNewLine & vbNewLine & _
                   "Remove those " & totalEmpty & "?" & vbNewLine & vbNewLine & _
                   "A rule set up in advance for data that has not arrived yet counts as empty, " & _
                   "so check the report first if that is likely.") Then
        ShowReport rows, "Conditional formatting"
        Exit Sub
    End If
    If Not OfferBackup(wb) Then Exit Sub

    FastMode True
    For Each ws In wb.Worksheets
        removed = removed + DeleteEmptyRules(ws)
    Next ws
    FastMode False

    ReportHeading rows, "Result"
    ReportRow rows, "Rules removed", removed

    ShowReport rows, "Conditional formatting"
    MsgBox removed & " rule(s) removed.", vbInformation, DOCTOR_NAME
End Sub

'==============================================================================
' Shared helpers - also used by the audit and the one-click run
'==============================================================================
' Deletes everything below and to the right of the real data on one sheet.
Public Sub TrimSheet(ByVal ws As Worksheet, ByRef rowsRemoved As Long, ByRef colsRemoved As Long)
    Dim keepRow As Long, keepCol As Long
    Dim usedLastRow As Long, usedLastCol As Long
    Dim ignore As Long

    If ws Is Nothing Then Exit Sub
    On Error GoTo Done
    If ws.ProtectContents Then Exit Sub

    usedLastRow = ws.UsedRange.Row + ws.UsedRange.rows.count - 1
    usedLastCol = ws.UsedRange.column + ws.UsedRange.Columns.count - 1

    keepRow = LastRow(ws)
    keepCol = LastColumn(ws)
    If LastShapeRow(ws) > keepRow Then keepRow = LastShapeRow(ws)
    If LastShapeColumn(ws) > keepCol Then keepCol = LastShapeColumn(ws)

    ' Never trim a sheet down to nothing - keep the first row and column.
    If keepRow < 1 Then keepRow = 1
    If keepCol < 1 Then keepCol = 1

    If usedLastRow > keepRow Then
        ws.Range(ws.rows(keepRow + 1), ws.rows(ws.rows.count)).Delete
        rowsRemoved = usedLastRow - keepRow
    End If

    If usedLastCol > keepCol Then
        ws.Range(ws.Columns(keepCol + 1), ws.Columns(ws.Columns.count)).Delete
        colsRemoved = usedLastCol - keepCol
    End If

    ignore = ws.UsedRange.rows.count            ' forces Excel to rebuild it
Done:
End Sub

Public Function SheetIsEmpty(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    If ws Is Nothing Then Exit Function
    If Application.CountA(ws.Cells) > 0 Then Exit Function
    If ws.Shapes.count > 0 Then Exit Function
    If ws.ListObjects.count > 0 Then Exit Function
    If ws.PivotTables.count > 0 Then Exit Function
    If ws.QueryTables.count > 0 Then Exit Function
    SheetIsEmpty = True
    On Error GoTo 0
End Function

' Deletes every hidden or zero-size drawing object in the workbook and returns
' the count. Used by the one-click run, which does not stop to list them.
Public Function DeleteJunkShapes(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Dim shp As Shape
    Dim i As Long
    Dim removed As Long

    If wb Is Nothing Then Exit Function

    On Error Resume Next
    For Each ws In wb.Worksheets
        For i = ws.Shapes.count To 1 Step -1
            Set shp = ws.Shapes(i)
            If Err.Number <> 0 Then
                Err.Clear
            ElseIf ShapeIsJunk(shp) Then
                shp.Delete
                If Err.Number = 0 Then removed = removed + 1 Else Err.Clear
            End If
        Next i
    Next ws
    On Error GoTo 0

    DeleteJunkShapes = removed
End Function

Public Function RuleCount(ByVal ws As Worksheet) As Long
    On Error Resume Next
    RuleCount = ws.Cells.FormatConditions.count
    If Err.Number <> 0 Then RuleCount = 0: Err.Clear
    On Error GoTo 0
End Function

Public Function CountJunkShapes(ByVal ws As Worksheet) As Long
    Dim shp As Shape
    Dim total As Long

    On Error Resume Next
    For Each shp In ws.Shapes
        If ShapeIsJunk(shp) Then total = total + 1
    Next shp
    On Error GoTo 0
    CountJunkShapes = total
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Function UsedRangeAddress(ByVal ws As Worksheet) As String
    On Error Resume Next
    UsedRangeAddress = ws.UsedRange.Address(False, False)
    If Err.Number <> 0 Then UsedRangeAddress = "?": Err.Clear
    On Error GoTo 0
End Function

Private Function LastShapeRow(ByVal ws As Worksheet) As Long
    Dim shp As Shape
    Dim r As Long
    Dim maxRow As Long

    On Error Resume Next
    For Each shp In ws.Shapes
        r = shp.BottomRightCell.Row
        If Err.Number <> 0 Then
            Err.Clear
        ElseIf r > maxRow Then
            maxRow = r
        End If
    Next shp
    On Error GoTo 0
    LastShapeRow = maxRow
End Function

Private Function LastShapeColumn(ByVal ws As Worksheet) As Long
    Dim shp As Shape
    Dim c As Long
    Dim maxCol As Long

    On Error Resume Next
    For Each shp In ws.Shapes
        c = shp.BottomRightCell.column
        If Err.Number <> 0 Then
            Err.Clear
        ElseIf c > maxCol Then
            maxCol = c
        End If
    Next shp
    On Error GoTo 0
    LastShapeColumn = maxCol
End Function

Private Function ShapeIsJunk(ByVal shp As Shape) As Boolean
    On Error Resume Next
    If shp.Visible = msoFalse Then ShapeIsJunk = True
    If shp.Width < 1 Then ShapeIsJunk = True
    If shp.Height < 1 Then ShapeIsJunk = True
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Private Function JunkReason(ByVal shp As Shape) As String
    On Error Resume Next
    If shp.Visible = msoFalse Then
        JunkReason = "Hidden"
    ElseIf shp.Width < 1 Or shp.Height < 1 Then
        JunkReason = "No size"
    Else
        JunkReason = "Unknown"
    End If
    On Error GoTo 0
End Function

Private Function ShapeTypeText(ByVal shp As Shape) As String
    On Error Resume Next
    Select Case shp.Type
        Case msoPicture: ShapeTypeText = "Picture"
        Case msoTextBox: ShapeTypeText = "Text box"
        Case msoAutoShape: ShapeTypeText = "AutoShape"
        Case msoFormControl: ShapeTypeText = "Form control"
        Case msoOLEControlObject: ShapeTypeText = "ActiveX control"
        Case msoChart: ShapeTypeText = "Chart"
        Case msoLine: ShapeTypeText = "Line"
        Case msoGroup: ShapeTypeText = "Group"
        Case Else: ShapeTypeText = "Type " & shp.Type
    End Select
    If Err.Number <> 0 Then ShapeTypeText = "?": Err.Clear
    On Error GoTo 0
End Function

Private Function EmptyRuleCount(ByVal ws As Worksheet) As Long
    Dim i As Long
    Dim total As Long
    Dim rules As Object

    On Error Resume Next
    Set rules = ws.Cells.FormatConditions
    If Err.Number <> 0 Then Err.Clear: Exit Function

    For i = 1 To rules.count
        If RuleAppliesToNothing(rules(i)) Then total = total + 1
    Next i
    On Error GoTo 0

    EmptyRuleCount = total
End Function

Private Function DeleteEmptyRules(ByVal ws As Worksheet) As Long
    Dim i As Long
    Dim removed As Long
    Dim rules As Object

    On Error Resume Next
    Set rules = ws.Cells.FormatConditions
    If Err.Number <> 0 Then Err.Clear: Exit Function

    For i = rules.count To 1 Step -1
        If RuleAppliesToNothing(rules(i)) Then
            rules(i).Delete
            If Err.Number = 0 Then removed = removed + 1 Else Err.Clear
        End If
    Next i
    On Error GoTo 0

    DeleteEmptyRules = removed
End Function

Private Function RuleAppliesToNothing(ByVal rule As Object) As Boolean
    Dim target As Range

    On Error Resume Next
    Set target = rule.AppliesTo
    If Err.Number <> 0 Then Err.Clear: Exit Function
    If target Is Nothing Then Exit Function

    RuleAppliesToNothing = (Application.CountA(target) = 0)
    If Err.Number <> 0 Then RuleAppliesToNothing = False: Err.Clear
    On Error GoTo 0
End Function
