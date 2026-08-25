Attribute VB_Name = "modWrangleShape"
'==============================================================================
' modWrangleShape - getting data into a shape a pivot table will accept
'------------------------------------------------------------------------------
' Two jobs that stand between a management extract and any analysis of it:
' months running across the top instead of down the side, and group names that
' appear only on the first row of each group.
'
' Depends on: modDoctorCommon, modWrangleStack (for the prompts), modRange,
'             modWorkbook, modApp (from \src)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' UnpivotSelection   (menu action)
' Turns a cross-tab into a tall table.
'
'   Client   Jan   Feb   Mar            Client   Period   Value
'   Acme     100   110   120     ->     Acme     Jan      100
'   Beta      90    95   105            Acme     Feb      110
'                                       Acme     Mar      120
'                                       Beta     Jan       90 ...
'------------------------------------------------------------------------------
Public Sub UnpivotSelection()
    Dim source As Range
    Dim data As Variant
    Dim keyColumns As Long
    Dim out As Variant
    Dim r As Long, c As Long, k As Long
    Dim outRow As Long
    Dim rowCount As Long, colCount As Long
    Dim valueColumns As Long
    Dim skipBlanks As Boolean
    Dim book As Workbook
    Dim ws As Worksheet
    Dim attributeName As String
    Dim valueName As String
    Dim cancelled As Boolean

    Set source = AskForBlock("Select the cross-tab, including its header row.")
    If source Is Nothing Then Exit Sub

    If source.rows.count < 2 Or source.Columns.count < 2 Then
        MsgBox "That needs to be at least two rows by two columns - a header row " & _
               "and something under it.", vbExclamation, DOCTOR_NAME
        Exit Sub
    End If

    keyColumns = AskForNumber("How many columns on the left are labels rather than periods?", 1)
    If keyColumns < 1 Then Exit Sub
    If keyColumns >= source.Columns.count Then
        MsgBox "That leaves no columns to unpivot.", vbExclamation, DOCTOR_NAME
        Exit Sub
    End If

    attributeName = AskForText("What should the new column of headings be called?", "Period", cancelled)
    If cancelled Then Exit Sub
    If Len(attributeName) = 0 Then attributeName = "Period"

    valueName = AskForText("And the new column of numbers?", "Value", cancelled)
    If cancelled Then Exit Sub
    If Len(valueName) = 0 Then valueName = "Value"

    skipBlanks = (MsgBox("Leave out rows where the value is blank?", _
                         vbYesNo + vbQuestion, DOCTOR_NAME) = vbYes)

    StartTimer
    FastMode True

    data = RangeToArray(source)
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)
    valueColumns = colCount - keyColumns

    ReDim out(1 To (rowCount - 1) * valueColumns + 1, 1 To keyColumns + 2)
    For k = 1 To keyColumns
        out(1, k) = data(1, k)
    Next k
    out(1, keyColumns + 1) = attributeName
    out(1, keyColumns + 2) = valueName

    outRow = 1
    For r = 2 To rowCount
        For c = keyColumns + 1 To colCount
            If Not (skipBlanks And IsBlank(data(r, c))) Then
                outRow = outRow + 1
                For k = 1 To keyColumns
                    out(outRow, k) = data(r, k)
                Next k
                out(outRow, keyColumns + 1) = data(1, c)
                out(outRow, keyColumns + 2) = data(r, c)
            End If
        Next c
    Next r

    FastMode False

    If outRow < 2 Then
        MsgBox "Nothing to write - every value was blank.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    out = Trimmed(out, outRow)

    Set book = Application.Workbooks.Add
    Set ws = DumpToSheet(out, "Unpivoted", book, True)
    If Not ws Is Nothing Then
        FreezeHeader ws, 1
        AutoFitColumns ws, 40, 9
    End If

    MsgBox Format$(rowCount - 1, "#,##0") & " row(s) by " & valueColumns & " period(s) became " & _
           Format$(outRow - 1, "#,##0") & " row(s), in " & ElapsedText & ".", _
           vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' FillBlanksDown   (menu action)
' The other thing standing between an extract and a pivot table: a label that
' appears once at the top of each group and is blank all the way down.
' This one changes the sheet, so it asks first.
'------------------------------------------------------------------------------
Public Sub FillBlanksDown()
    Dim source As Range
    Dim data As Variant
    Dim r As Long, c As Long
    Dim filled As Long
    Dim lastValue As Variant

    Set source = AskForBlock("Select the column or columns to fill down." & vbNewLine & vbNewLine & _
                             "Do not include the header row.")
    If source Is Nothing Then Exit Sub

    If source.Cells.CountLarge > 500000 Then
        MsgBox "That is a very large selection. Narrow it to the rows that hold data.", _
               vbExclamation, DOCTOR_NAME
        Exit Sub
    End If

    If Not Confirm("Fill every blank cell in " & source.Address(False, False) & " with the " & _
                   "value above it?" & vbNewLine & vbNewLine & _
                   "Formulas in the selection would be overwritten by their own values.") Then Exit Sub

    StartTimer
    FastMode True

    data = RangeToArray(source)
    For c = 1 To UBound(data, 2)
        lastValue = Empty
        For r = 1 To UBound(data, 1)
            If IsBlank(data(r, c)) Then
                If Not IsEmpty(lastValue) Then
                    data(r, c) = lastValue
                    filled = filled + 1
                End If
            Else
                lastValue = data(r, c)
            End If
        Next r
    Next c

    source.value = data
    FastMode False

    MsgBox Format$(filled, "#,##0") & " blank cell(s) filled in " & ElapsedText & ".", _
           vbInformation, DOCTOR_NAME
End Sub

'==============================================================================
' Shared
'==============================================================================
Public Function AskForBlock(ByVal prompt As String) As Range
    Dim chosen As Range
    Dim suggestion As String

    On Error Resume Next
    suggestion = Selection.Address
    Err.Clear
    Set chosen = Application.InputBox(prompt, DOCTOR_NAME, suggestion, Type:=8)
    On Error GoTo 0

    If chosen Is Nothing Then Exit Function

    ' Range.Value only ever returns the first area, so a scattered selection
    ' would quietly lose most of itself.
    If chosen.Areas.count > 1 Then
        MsgBox "Select one continuous block. That selection is in " & chosen.Areas.count & _
               " separate pieces.", vbExclamation, DOCTOR_NAME
        Exit Function
    End If

    Set AskForBlock = chosen
End Function

'==============================================================================
' Private helpers
'==============================================================================
' Cuts an over-allocated output array down to the rows actually used.
Private Function Trimmed(ByVal data As Variant, ByVal usedRows As Long) As Variant
    Dim out As Variant
    Dim r As Long, c As Long

    If usedRows >= UBound(data, 1) Then Trimmed = data: Exit Function

    ReDim out(1 To usedRows, 1 To UBound(data, 2))
    For r = 1 To usedRows
        For c = 1 To UBound(data, 2)
            out(r, c) = data(r, c)
        Next c
    Next r
    Trimmed = out
End Function
