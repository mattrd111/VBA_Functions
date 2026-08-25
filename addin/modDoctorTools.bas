Attribute VB_Name = "modDoctorTools"
'==============================================================================
' modDoctorTools - everyday tools that act on the selection
'------------------------------------------------------------------------------
' Tidying text, freezing formulas, stripping hyperlinks, and taking a backup.
'
' Depends on: modDoctorCommon, modString, modApp (from \src)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' TrimAndCleanSelection   (menu action)
' Removes the invisible mess that arrives with data pasted from the web, a PDF
' or a mainframe extract: non-breaking spaces, control characters, tabs and
' doubled or trailing spaces. Formulas and numbers are left alone.
'------------------------------------------------------------------------------
Public Sub TrimAndCleanSelection()
    Dim target As Range
    Dim area As Range
    Dim values As Variant
    Dim original As String, cleaned As String
    Dim r As Long, c As Long
    Dim changed As Long, looked As Long
    Dim areaChanged As Long

    Set target = TextCellsInSelection()
    If target Is Nothing Then Exit Sub

    StartTimer
    FastMode True

    For Each area In target.Areas
        If area.Cells.CountLarge = 1 Then
            original = CStr(area.value)
            cleaned = CleanText(original)
            looked = looked + 1
            If StrComp(original, cleaned, vbBinaryCompare) <> 0 Then
                area.value = cleaned
                changed = changed + 1
            End If
        Else
            values = area.value
            areaChanged = 0
            For r = LBound(values, 1) To UBound(values, 1)
                For c = LBound(values, 2) To UBound(values, 2)
                    looked = looked + 1
                    original = CStr(values(r, c))
                    cleaned = CleanText(original)
                    If StrComp(original, cleaned, vbBinaryCompare) <> 0 Then
                        values(r, c) = cleaned
                        areaChanged = areaChanged + 1
                    End If
                Next c
            Next r
            If areaChanged > 0 Then
                area.value = values
                changed = changed + areaChanged
            End If
        End If
    Next area

    FastMode False

    MsgBox Format$(looked, "#,##0") & " text cell(s) checked, " & _
           Format$(changed, "#,##0") & " tidied in " & ElapsedText & ".", _
           vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' SelectionToValues   (menu action)
' Replaces formulas with what they currently show.
'------------------------------------------------------------------------------
Public Sub SelectionToValues()
    Dim target As Range
    Dim formulaCells As Range
    Dim area As Range
    Dim count As Long

    Set target = SelectedRange()
    If target Is Nothing Then Exit Sub

    On Error Resume Next
    Set formulaCells = target.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If formulaCells Is Nothing Then
        MsgBox "There are no formulas in the selection.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    count = formulaCells.Cells.count
    If Not Confirm("Replace " & Format$(count, "#,##0") & " formula(s) in the selection with " & _
                   "their current values?") Then Exit Sub

    FastMode True
    For Each area In formulaCells.Areas
        area.value = area.value
    Next area
    FastMode False

    MsgBox Format$(count, "#,##0") & " formula(s) converted to values.", vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' RemoveHyperlinksFromSelection   (menu action)
' Strips the links but keeps the text - and the underlined blue formatting goes
' with them.
'------------------------------------------------------------------------------
Public Sub RemoveHyperlinksFromSelection()
    Dim target As Range
    Dim count As Long

    Set target = SelectedRange()
    If target Is Nothing Then Exit Sub

    On Error Resume Next
    count = target.Hyperlinks.count
    On Error GoTo 0
    If count = 0 Then
        MsgBox "There are no hyperlinks in the selection.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    If Not Confirm("Remove " & Format$(count, "#,##0") & " hyperlink(s) from the selection?" & _
                   vbNewLine & "The cell text stays.") Then Exit Sub

    FastMode True
    On Error Resume Next
    target.Hyperlinks.Delete
    On Error GoTo 0
    FastMode False

    MsgBox Format$(count, "#,##0") & " hyperlink(s) removed.", vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' BackupActiveWorkbook   (menu action)
' A timestamped copy beside the original. The original stays open and untouched.
'------------------------------------------------------------------------------
Public Sub BackupActiveWorkbook()
    Dim wb As Workbook
    Dim path As String

    Set wb = TargetWorkbook(False)
    If wb Is Nothing Then Exit Sub

    If Len(wb.Path) = 0 Then
        MsgBox "'" & wb.Name & "' has never been saved, so there is nowhere to put a copy." & _
               vbNewLine & "Save it first.", vbExclamation, DOCTOR_NAME
        Exit Sub
    End If

    Application.Cursor = xlWait
    path = BackupWorkbook(wb)
    Application.Cursor = xlDefault

    If Len(path) = 0 Then
        MsgBox "The backup could not be written. Check that the folder is writable.", _
               vbExclamation, DOCTOR_NAME
    Else
        MsgBox "Backup saved as:" & vbNewLine & vbNewLine & path, vbInformation, DOCTOR_NAME
    End If
End Sub

'------------------------------------------------------------------------------
' AboutDoctor   (menu action)
'------------------------------------------------------------------------------
Public Sub AboutDoctor()
    MsgBox DOCTOR_NAME & "  version " & DOCTOR_VERSION & vbNewLine & _
           String$(52, "-") & vbNewLine & vbNewLine & _
           "AUDIT   reads and reports, changes nothing." & vbNewLine & _
           "        Workbook: what is making the file heavy." & vbNewLine & _
           "        Model: rows that stopped calculating, numbers typed over" & vbNewLine & _
           "        formulas, assumptions buried inside them." & vbNewLine & vbNewLine & _
           "CLEAN   unused names, duplicate styles, runaway used ranges," & vbNewLine & _
           "        stray objects, dead links." & vbNewLine & vbNewLine & _
           "DATA    stack many extracts into one table, unpivot a cross-tab," & vbNewLine & _
           "        fill blanks down, match two lists that nearly agree." & vbNewLine & vbNewLine & _
           "Everything destructive asks first and offers a backup. None of it can " & _
           "be undone with Ctrl+Z, so take the backup." & vbNewLine & vbNewLine & _
           "Start with Audit - it costs nothing and tells you what is worth doing." & _
           vbNewLine & vbNewLine & _
           "Source: github.com/mattrd111/VBA_Functions", _
           vbInformation, DOCTOR_NAME
End Sub

'==============================================================================
' Private helpers
'==============================================================================
Private Function SelectedRange() As Range
    Dim target As Range

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open a workbook first.", vbInformation, DOCTOR_NAME
        Exit Function
    End If

    On Error Resume Next
    Set target = Application.Selection
    On Error GoTo 0

    If target Is Nothing Then
        MsgBox "Select some cells first.", vbInformation, DOCTOR_NAME
        Exit Function
    End If

    ' A whole-column selection would otherwise mean a million empty cells.
    If target.Cells.CountLarge > 100000 Then
        On Error Resume Next
        Set target = Application.Intersect(target, target.Worksheet.UsedRange)
        On Error GoTo 0
        If target Is Nothing Then
            MsgBox "The selection holds no data.", vbInformation, DOCTOR_NAME
            Exit Function
        End If
    End If

    Set SelectedRange = target
End Function

Private Function TextCellsInSelection() As Range
    Dim target As Range
    Dim textCells As Range

    Set target = SelectedRange()
    If target Is Nothing Then Exit Function

    On Error Resume Next
    Set textCells = target.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0

    If textCells Is Nothing Then
        MsgBox "There is no text in the selection to tidy.", vbInformation, DOCTOR_NAME
        Exit Function
    End If

    Set TextCellsInSelection = textCells
End Function
