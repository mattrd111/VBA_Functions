Attribute VB_Name = "modRange"
'==============================================================================
' modRange - worksheet range helpers
'------------------------------------------------------------------------------
' Finding the edges of your data, reading and writing whole blocks in one hit,
' locating columns by header text, and the usual tidy-up jobs.
'
' Standalone: no dependency on the other modules in this repository.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' LastRow
' Last used row on a sheet, or in one column when a column is supplied.
' Returns 0 for an empty sheet or column. The column may be a number or a
' letter: LastRow(ws, "B") and LastRow(ws, 2) are the same thing.
'------------------------------------------------------------------------------
Public Function LastRow(ByVal ws As Worksheet, Optional ByVal column As Variant) As Long
    Dim found As Range

    If ws Is Nothing Then Exit Function

    If IsMissing(column) Then
        On Error Resume Next
        Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        On Error GoTo 0
        If Not found Is Nothing Then LastRow = found.Row
    Else
        If Application.CountA(ws.Columns(column)) = 0 Then Exit Function
        LastRow = ws.Cells(ws.Rows.Count, column).End(xlUp).Row
    End If
End Function

'------------------------------------------------------------------------------
' LastColumn
' Last used column on a sheet, or in one row when a row is supplied.
' Returns 0 when there is nothing to find.
'------------------------------------------------------------------------------
Public Function LastColumn(ByVal ws As Worksheet, Optional ByVal row As Variant) As Long
    Dim found As Range

    If ws Is Nothing Then Exit Function

    If IsMissing(row) Then
        On Error Resume Next
        Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                  SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
        On Error GoTo 0
        If Not found Is Nothing Then LastColumn = found.column
    Else
        If Application.CountA(ws.Rows(row)) = 0 Then Exit Function
        LastColumn = ws.Cells(row, ws.Columns.Count).End(xlToLeft).column
    End If
End Function

'------------------------------------------------------------------------------
' DataRange
' The rectangular block of data on a sheet, from the header row down to the
' last used row and across to the last used column. Nothing when the sheet is
' empty. More reliable than UsedRange, which remembers deleted cells.
'------------------------------------------------------------------------------
Public Function DataRange(ByVal ws As Worksheet, Optional ByVal headerRow As Long = 1) As Range
    Dim lastR As Long, lastC As Long

    If ws Is Nothing Then Exit Function
    lastR = LastRow(ws)
    lastC = LastColumn(ws)
    If lastR < headerRow Or lastC = 0 Then Exit Function

    Set DataRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastR, lastC))
End Function

'------------------------------------------------------------------------------
' BodyRange
' Same as DataRange but without the header row - the rows you actually loop over.
'------------------------------------------------------------------------------
Public Function BodyRange(ByVal ws As Worksheet, Optional ByVal headerRow As Long = 1) As Range
    Dim block As Range

    Set block = DataRange(ws, headerRow)
    If block Is Nothing Then Exit Function
    If block.Rows.Count <= 1 Then Exit Function

    Set BodyRange = block.Offset(1).Resize(block.Rows.Count - 1)
End Function

'------------------------------------------------------------------------------
' RangeToArray
' Values of a range as a 1-based 2D array, even for a single cell - so callers
' never have to special case the one cell that would otherwise come back as a
' bare value.
'------------------------------------------------------------------------------
Public Function RangeToArray(ByVal rng As Range) As Variant
    Dim out As Variant

    If rng Is Nothing Then Exit Function
    If rng.Cells.CountLarge = 1 Then
        ReDim out(1 To 1, 1 To 1)
        out(1, 1) = rng.value
        RangeToArray = out
    Else
        RangeToArray = rng.value
    End If
End Function

'------------------------------------------------------------------------------
' WriteArray
' Writes a 1D or 2D array to a sheet in a single operation, starting at
' topLeftCell. Vastly faster than writing cell by cell.
'
'   WriteArray results, ws.Range("A2")
'------------------------------------------------------------------------------
Public Sub WriteArray(ByVal values As Variant, ByVal topLeftCell As Range, _
                      Optional ByVal asRow As Boolean = False)
    Dim block As Variant
    Dim rows As Long, cols As Long
    Dim i As Long, n As Long

    If topLeftCell Is Nothing Then Exit Sub
    If Not IsArray(values) Then
        topLeftCell.value = values
        Exit Sub
    End If

    On Error GoTo Done
    If UBound(values, 1) < LBound(values, 1) Then Exit Sub

    If Dimensions(values) = 1 Then
        n = UBound(values) - LBound(values) + 1
        If asRow Then ReDim block(1 To 1, 1 To n) Else ReDim block(1 To n, 1 To 1)
        For i = 1 To n
            If asRow Then
                block(1, i) = values(LBound(values) + i - 1)
            Else
                block(i, 1) = values(LBound(values) + i - 1)
            End If
        Next i
    Else
        block = values
    End If

    rows = UBound(block, 1) - LBound(block, 1) + 1
    cols = UBound(block, 2) - LBound(block, 2) + 1
    topLeftCell.Resize(rows, cols).value = block
Done:
End Sub

'------------------------------------------------------------------------------
' HeaderMap
' Dictionary of header text -> column number for a header row. Lookups are
' case insensitive and ignore stray spacing, so "Client Name", "client name"
' and " Client  Name " all match.
'
'   Set headers = HeaderMap(ws)
'   amountCol = headers("Amount")
'------------------------------------------------------------------------------
Public Function HeaderMap(ByVal ws As Worksheet, Optional ByVal headerRow As Long = 1) As Object
    Dim map As Object
    Dim lastC As Long, c As Long
    Dim key As String

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = 1                     ' TextCompare
    Set HeaderMap = map
    If ws Is Nothing Then Exit Function

    lastC = LastColumn(ws, headerRow)
    For c = 1 To lastC
        key = NormalizeHeader(ws.Cells(headerRow, c).value)
        If Len(key) > 0 Then
            If Not map.Exists(key) Then map.Add key, c
        End If
    Next c
End Function

'------------------------------------------------------------------------------
' FindColumnByHeader
' Column number of a header, or 0 when it is not there. Same matching rules as
' HeaderMap. Use it instead of hard-coding column letters that move.
'------------------------------------------------------------------------------
Public Function FindColumnByHeader(ByVal ws As Worksheet, ByVal headerText As String, _
                                   Optional ByVal headerRow As Long = 1) As Long
    Dim lastC As Long, c As Long
    Dim target As String

    If ws Is Nothing Then Exit Function
    target = NormalizeHeader(headerText)
    If Len(target) = 0 Then Exit Function

    lastC = LastColumn(ws, headerRow)
    For c = 1 To lastC
        If StrComp(NormalizeHeader(ws.Cells(headerRow, c).value), target, vbTextCompare) = 0 Then
            FindColumnByHeader = c
            Exit Function
        End If
    Next c
End Function

'------------------------------------------------------------------------------
' FindCell
' Wrapper around Range.Find that returns Nothing instead of raising, and always
' searches from the top left so results are repeatable.
'------------------------------------------------------------------------------
Public Function FindCell(ByVal searchRange As Range, ByVal what As Variant, _
                         Optional ByVal wholeCell As Boolean = True, _
                         Optional ByVal searchValues As Boolean = True) As Range
    If searchRange Is Nothing Then Exit Function
    On Error Resume Next
    Set FindCell = searchRange.Find(What:=what, _
                                    After:=searchRange.Cells(searchRange.rows.Count, searchRange.Columns.Count), _
                                    LookIn:=IIf(searchValues, xlValues, xlFormulas), _
                                    LookAt:=IIf(wholeCell, xlWhole, xlPart), _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlNext, _
                                    MatchCase:=False)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' ColumnLetter / ColumnNumber
'
'   ColumnLetter(28)    ->  "AB"
'   ColumnNumber("AB")  ->  28
'------------------------------------------------------------------------------
Public Function ColumnLetter(ByVal columnIndex As Long) As String
    Dim n As Long, r As Long

    If columnIndex < 1 Or columnIndex > 16384 Then Exit Function
    n = columnIndex
    Do While n > 0
        r = (n - 1) Mod 26
        ColumnLetter = Chr$(65 + r) & ColumnLetter
        n = (n - 1) \ 26
    Loop
End Function

Public Function ColumnNumber(ByVal columnLetters As String) As Long
    Dim i As Long, code As Long, total As Long

    columnLetters = UCase$(Trim$(columnLetters))
    If Len(columnLetters) = 0 Or Len(columnLetters) > 3 Then Exit Function

    For i = 1 To Len(columnLetters)
        code = Asc(Mid$(columnLetters, i, 1)) - 64
        If code < 1 Or code > 26 Then Exit Function
        total = total * 26 + code
    Next i
    If total > 16384 Then Exit Function
    ColumnNumber = total
End Function

'------------------------------------------------------------------------------
' IsRangeEmpty
' True when a range holds no values, formulas or errors.
'------------------------------------------------------------------------------
Public Function IsRangeEmpty(ByVal rng As Range) As Boolean
    If rng Is Nothing Then IsRangeEmpty = True: Exit Function
    IsRangeEmpty = (Application.CountA(rng) = 0)
End Function

'------------------------------------------------------------------------------
' CopyValues
' Copies values only from one range to another in one operation - no clipboard,
' no formats, no formulas.
'------------------------------------------------------------------------------
Public Sub CopyValues(ByVal source As Range, ByVal destinationTopLeft As Range)
    If source Is Nothing Or destinationTopLeft Is Nothing Then Exit Sub
    destinationTopLeft.Resize(source.rows.Count, source.Columns.Count).value = source.value
End Sub

'------------------------------------------------------------------------------
' DeleteEmptyRows
' Removes blank rows from a sheet in a single delete operation. Pass a column
' number to treat rows as blank when just that column is empty.
'------------------------------------------------------------------------------
Public Sub DeleteEmptyRows(ByVal ws As Worksheet, Optional ByVal checkColumn As Long = 0, _
                           Optional ByVal firstRow As Long = 1)
    Dim r As Long, lastR As Long
    Dim doomed As Range
    Dim rowIsEmpty As Boolean

    If ws Is Nothing Then Exit Sub
    lastR = LastRow(ws)
    If lastR < firstRow Then Exit Sub

    For r = firstRow To lastR
        If checkColumn > 0 Then
            rowIsEmpty = (Application.CountA(ws.Cells(r, checkColumn)) = 0)
        Else
            rowIsEmpty = (Application.CountA(ws.rows(r)) = 0)
        End If
        If rowIsEmpty Then
            If doomed Is Nothing Then
                Set doomed = ws.rows(r)
            Else
                Set doomed = Union(doomed, ws.rows(r))
            End If
        End If
    Next r

    If Not doomed Is Nothing Then doomed.Delete
End Sub

'------------------------------------------------------------------------------
' AutoFitColumns
' AutoFit with a sensible upper bound, because one long comment should not make
' a column 200 characters wide.
'------------------------------------------------------------------------------
Public Sub AutoFitColumns(ByVal ws As Worksheet, Optional ByVal maxWidth As Double = 60, _
                          Optional ByVal minWidth As Double = 8)
    Dim c As Long, lastC As Long

    If ws Is Nothing Then Exit Sub
    lastC = LastColumn(ws)
    If lastC = 0 Then Exit Sub

    ws.Range(ws.Columns(1), ws.Columns(lastC)).AutoFit
    For c = 1 To lastC
        If ws.Columns(c).ColumnWidth > maxWidth Then ws.Columns(c).ColumnWidth = maxWidth
        If ws.Columns(c).ColumnWidth < minWidth Then ws.Columns(c).ColumnWidth = minWidth
    Next c
End Sub

'------------------------------------------------------------------------------
' FreezeHeader
' Freezes the top rows (and optionally the left columns) of a sheet, restoring
' whichever sheet was active before.
'------------------------------------------------------------------------------
Public Sub FreezeHeader(ByVal ws As Worksheet, Optional ByVal rowsToFreeze As Long = 1, _
                        Optional ByVal columnsToFreeze As Long = 0)
    Dim wb As Workbook
    Dim previousSheet As Object
    Dim previousBook As Workbook
    Dim screenWasOn As Boolean

    If ws Is Nothing Then Exit Sub
    Set wb = ws.Parent
    If wb.Windows.Count = 0 Then Exit Sub

    screenWasOn = Application.ScreenUpdating
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Set previousBook = ActiveWorkbook
    Set previousSheet = wb.ActiveSheet

    wb.Activate
    ws.Activate
    With wb.Windows(1)
        .FreezePanes = False
        .SplitRow = rowsToFreeze
        .SplitColumn = columnsToFreeze
        .FreezePanes = True
    End With

Cleanup:
    On Error Resume Next
    If Not previousSheet Is Nothing Then previousSheet.Activate
    If Not previousBook Is Nothing Then previousBook.Activate
    Application.ScreenUpdating = screenWasOn
    On Error GoTo 0
End Sub

'==============================================================================
' Private helpers
'==============================================================================
' Header text with control characters, non-breaking spaces and repeated spaces
' normalised away.
Private Function NormalizeHeader(ByVal value As Variant) As String
    Dim s As String
    Dim i As Long
    Dim code As Long
    Dim ch As String

    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    s = CStr(value)

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 9, 10, 11, 12, 13, 160, 8194, 8195, 8201, 8239
                Mid$(s, i, 1) = " "
            Case 0 To 31, 127
                Mid$(s, i, 1) = " "
        End Select
    Next i

    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    NormalizeHeader = Trim$(s)
End Function

Private Function Dimensions(ByVal arr As Variant) As Long
    Dim d As Long
    Dim test As Long

    If Not IsArray(arr) Then Exit Function
    On Error GoTo Done
    Do
        d = d + 1
        test = LBound(arr, d)
    Loop While d < 60
Done:
    Dimensions = d - 1
End Function
