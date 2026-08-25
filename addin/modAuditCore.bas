Attribute VB_Name = "modAuditCore"
'==============================================================================
' modAuditCore - the row consistency engine
'------------------------------------------------------------------------------
' The idea every model audit tool is built on: in a working model, a row of a
' calculation block holds one formula copied across. Read the block in R1C1 and
' a copied formula is the same string in every column, so the odd one out falls
' straight out of the comparison. That is where the expensive mistakes live -
' the cell someone typed over in March, three versions ago.
'
' Rules of engagement, chosen to keep the signal worth reading:
'   - a run must be at least 4 cells long before it is judged at all
'   - the dominant formula must appear at least 3 times and hold a majority
'   - a number typed into a row of formulas is High - that is the classic
'     override
'   - a formula that differs at the first or last cell of a run is Medium, not
'     High, because opening balances and closing columns are often meant to
'     differ
'   - array and spilled formulas are left alone; their R1C1 shifts per cell by
'     design and would otherwise light up the whole report
'
' Rows only, not columns. Financial models run periods across the page, and
' scanning both directions doubles the false positives without finding much.
'
' Depends on: modAuditFormula, modDictionary, modRange, modApp (from \src)
'==============================================================================
Option Explicit

Public Const SEV_HIGH As String = "High"
Public Const SEV_MEDIUM As String = "Medium"
Public Const SEV_LOW As String = "Low"
Public Const SEV_INFO As String = "Info"

Private Const MIN_RUN As Long = 4               ' shortest run worth judging
Private Const MIN_PATTERN As Long = 3           ' occurrences before a pattern is "the" pattern
Private Const MAX_PER_CHECK As Long = 150       ' per sheet, per check
Private Const MAX_BAND_CELLS As Long = 250000   ' cells read in one go
Private Const LONG_FORMULA As Long = 250        ' characters

Private Const KIND_EMPTY As Long = 0
Private Const KIND_TEXT As Long = 1
Private Const KIND_NUMBER As Long = 2
Private Const KIND_FORMULA As Long = 3

'==============================================================================
' Findings
'------------------------------------------------------------------------------
' A finding is Array(severity, sheet, address, check, detail, formula).
'==============================================================================
Public Function NewFindings() As Collection
    Set NewFindings = New Collection
End Function

Public Sub AddFinding(ByVal findings As Collection, ByVal severity As String, _
                      ByVal sheetName As String, ByVal address As String, _
                      ByVal check As String, ByVal detail As String, _
                      Optional ByVal formulaText As String = "")
    If findings Is Nothing Then Exit Sub
    findings.Add Array(severity, sheetName, address, check, detail, Left$(formulaText, 300))
End Sub

Public Function NewStats() As Object
    Set NewStats = NewDictionary(True)
End Function

'==============================================================================
' ScanSheet
' One sheet, one pass. Fills findings and stats.
'==============================================================================
Public Sub ScanSheet(ByVal ws As Worksheet, ByVal findings As Collection, ByVal stats As Object)
    Dim block As Range
    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
    Dim colCount As Long, bandRows As Long, bandTop As Long, bandHeight As Long
    Dim a1 As Variant, r1c1 As Variant
    Dim patterns As Object
    Dim caps As Object
    Dim r As Long

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set block = ws.UsedRange
    On Error GoTo 0
    If block Is Nothing Then Exit Sub

    firstRow = block.Row
    firstCol = block.column
    lastRow = firstRow + block.rows.count - 1
    lastCol = firstCol + block.Columns.count - 1
    colCount = lastCol - firstCol + 1
    If colCount < 1 Then Exit Sub

    bandRows = MAX_BAND_CELLS \ colCount
    If bandRows < 1 Then bandRows = 1

    Set patterns = NewDictionary(False)         ' R1C1 is compared exactly
    Set caps = NewDictionary(True)

    For bandTop = firstRow To lastRow Step bandRows
        bandHeight = bandRows
        If bandTop + bandHeight - 1 > lastRow Then bandHeight = lastRow - bandTop + 1

        a1 = ReadBlock(ws.Range(ws.Cells(bandTop, firstCol), ws.Cells(bandTop + bandHeight - 1, lastCol)), False)
        r1c1 = ReadBlock(ws.Range(ws.Cells(bandTop, firstCol), ws.Cells(bandTop + bandHeight - 1, lastCol)), True)
        If Not IsArray(a1) Or Not IsArray(r1c1) Then Exit Sub

        For r = 1 To bandHeight
            CollectRow ws, a1, r1c1, r, bandTop + r - 1, firstCol, patterns, stats
            ScanRowRuns ws, r1c1, r, bandTop + r - 1, firstCol, findings, caps, stats
        Next r
    Next bandTop

    Bump stats, "Unique formulas", patterns.count
    CheckPatterns ws, patterns, findings, caps, stats
End Sub

'==============================================================================
' ScanSheetForFlags
' Just the consistency check, for the "select the suspect cells" action. Returns
' the cells it would flag, or Nothing.
'==============================================================================
Public Function ScanSheetForFlags(ByVal ws As Worksheet) As Range
    Dim findings As Collection
    Dim stats As Object
    Dim item As Variant
    Dim flagged As Range
    Dim cell As Range
    Dim areas As Long

    Set findings = NewFindings()
    Set stats = NewStats()
    ScanSheet ws, findings, stats

    For Each item In findings
        If IsConsistencyCheck(CStr(item(3))) Then
            On Error Resume Next
            Set cell = ws.Range(CStr(item(2)))
            On Error GoTo 0
            If Not cell Is Nothing Then
                If flagged Is Nothing Then
                    Set flagged = cell
                ElseIf areas < 800 Then
                    Set flagged = Union(flagged, cell)
                End If
                areas = areas + 1
            End If
            Set cell = Nothing
        End If
    Next item

    Set ScanSheetForFlags = flagged
End Function

Public Function IsConsistencyCheck(ByVal check As String) As Boolean
    Select Case check
        Case "Formula differs from the row", "Number typed into a calculated row"
            IsConsistencyCheck = True
    End Select
End Function

'==============================================================================
' Row handling
'==============================================================================
' Counts every formula on the row and remembers each distinct R1C1 pattern with
' where it was first seen, so each pattern is parsed once per sheet rather than
' once per cell.
Private Sub CollectRow(ByVal ws As Worksheet, ByRef a1 As Variant, ByRef r1c1 As Variant, _
                       ByVal arrayRow As Long, ByVal sheetRow As Long, ByVal firstCol As Long, _
                       ByVal patterns As Object, ByVal stats As Object)
    Dim c As Long
    Dim kind As Long
    Dim key As String
    Dim entry As Variant

    For c = 1 To UBound(r1c1, 2)
        kind = CellKind(r1c1(arrayRow, c))

        If kind = KIND_FORMULA Then
            Bump stats, "Formulas", 1
            key = CStr(r1c1(arrayRow, c))

            If patterns.Exists(key) Then
                entry = patterns(key)
                entry(0) = entry(0) + 1
                patterns(key) = entry
            Else
                patterns.Add key, Array(1&, CellAddress(sheetRow, firstCol + c - 1), CStr(a1(arrayRow, c)))
            End If

        ElseIf kind = KIND_NUMBER Then
            Bump stats, "Numbers", 1
        ElseIf kind = KIND_TEXT Then
            Bump stats, "Text cells", 1
        End If
    Next c
End Sub

' Splits the row into runs of adjacent non-empty cells and judges each one.
Private Sub ScanRowRuns(ByVal ws As Worksheet, ByRef r1c1 As Variant, _
                        ByVal arrayRow As Long, ByVal sheetRow As Long, ByVal firstCol As Long, _
                        ByVal findings As Collection, ByVal caps As Object, ByVal stats As Object)
    Dim c As Long, n As Long, runStart As Long

    n = UBound(r1c1, 2)
    c = 1

    Do While c <= n
        If CellKind(r1c1(arrayRow, c)) = KIND_EMPTY Then
            c = c + 1
        Else
            runStart = c
            Do While c <= n
                If CellKind(r1c1(arrayRow, c)) = KIND_EMPTY Then Exit Do
                c = c + 1
            Loop
            JudgeRun ws, r1c1, arrayRow, sheetRow, firstCol, runStart, c - 1, findings, caps, stats
        End If
    Loop
End Sub

' The heart of it: find the formula that dominates a run and report what does
' not match it.
Private Sub JudgeRun(ByVal ws As Worksheet, ByRef r1c1 As Variant, _
                     ByVal arrayRow As Long, ByVal sheetRow As Long, ByVal firstCol As Long, _
                     ByVal runStart As Long, ByVal runEnd As Long, _
                     ByVal findings As Collection, ByVal caps As Object, ByVal stats As Object)
    Dim counts As Object
    Dim c As Long
    Dim kind As Long
    Dim key As String
    Dim candidate As Variant
    Dim formulaCells As Long
    Dim dominant As String
    Dim dominantCount As Long
    Dim severity As String
    Dim address As String

    If runEnd - runStart + 1 < MIN_RUN Then Exit Sub

    Set counts = NewDictionary(False)
    For c = runStart To runEnd
        If CellKind(r1c1(arrayRow, c)) = KIND_FORMULA Then
            formulaCells = formulaCells + 1
            key = CStr(r1c1(arrayRow, c))
            If counts.Exists(key) Then counts(key) = counts(key) + 1 Else counts.Add key, 1&
        End If
    Next c

    If formulaCells < MIN_PATTERN Then Exit Sub

    For Each candidate In counts.keys
        If counts(candidate) > dominantCount Then
            dominantCount = counts(candidate)
            dominant = CStr(candidate)
        End If
    Next candidate

    ' No clear winner means this is a row of one-offs, not a copied row. Judging
    ' it would be guessing.
    If dominantCount < MIN_PATTERN Then Exit Sub
    If dominantCount * 2 <= formulaCells Then Exit Sub

    For c = runStart To runEnd
        kind = CellKind(r1c1(arrayRow, c))
        address = CellAddress(sheetRow, firstCol + c - 1)

        If kind = KIND_NUMBER Then
            If Not Capped(caps, "override", ws.Name) Then
                AddFinding findings, SEV_HIGH, ws.Name, address, _
                           "Number typed into a calculated row", _
                           "The rest of this row calculates; this cell holds a typed number.", _
                           CStr(r1c1(arrayRow, c))
            End If

        ElseIf kind = KIND_FORMULA Then
            If StrComp(CStr(r1c1(arrayRow, c)), dominant, vbBinaryCompare) <> 0 Then
                If Not IsArrayOrSpillCell(ws.Cells(sheetRow, firstCol + c - 1)) Then
                    If c = runStart Or c = runEnd Then severity = SEV_MEDIUM Else severity = SEV_HIGH
                    If Not Capped(caps, "consistency", ws.Name) Then
                        AddFinding findings, severity, ws.Name, address, _
                                   "Formula differs from the row", _
                                   "The other " & dominantCount & " cell(s) in this run use " & _
                                   Left$(dominant, 120) & _
                                   IIf(c = runStart Or c = runEnd, " (end of the run, so this may be intentional)", ""), _
                                   CStr(r1c1(arrayRow, c))
                    End If
                End If
            End If
        End If
    Next c
End Sub

'==============================================================================
' Formula content checks - one pass per distinct formula on the sheet
'==============================================================================
Private Sub CheckPatterns(ByVal ws As Worksheet, ByVal patterns As Object, _
                          ByVal findings As Collection, ByVal caps As Object, ByVal stats As Object)
    Dim key As Variant
    Dim entry As Variant
    Dim analysis As Object
    Dim hardcode As Variant
    Dim item As Variant
    Dim a1 As String
    Dim address As String
    Dim occurrences As Long
    Dim suffix As String
    Dim done As Long

    For Each key In patterns.keys
        done = done + 1
        If done Mod 200 = 0 Then ShowProgress done, patterns.count, "Reading formulas on " & ws.Name

        entry = patterns(key)
        occurrences = CLng(entry(0))
        address = CStr(entry(1))
        a1 = CStr(entry(2))
        suffix = IIf(occurrences > 1, "  (this formula appears " & Format$(occurrences, "#,##0") & " times)", "")

        If Len(a1) > LONG_FORMULA Then
            Bump stats, "Long formulas", occurrences
            If Not Capped(caps, "long", ws.Name) Then
                AddFinding findings, SEV_LOW, ws.Name, address, "Very long formula", _
                           Len(a1) & " characters. Long formulas hide their own mistakes." & suffix, a1
            End If
        End If

        Set analysis = AnalyseFormula(a1)

        For Each hardcode In analysis("Hardcodes")
            If IsSuspectConstant(CStr(hardcode(0)), CStr(hardcode(1)), CBool(hardcode(2))) Then
                Bump stats, "Hardcoded numbers", occurrences
                If Not Capped(caps, "hardcode", ws.Name) Then
                    AddFinding findings, ConstantSeverity(CStr(hardcode(0))), ws.Name, address, _
                               "Number inside a formula", _
                               "The value " & CStr(hardcode(0)) & " is typed into the formula" & _
                               IIf(Len(CStr(hardcode(1))) > 0, " inside " & CStr(hardcode(1)) & "()", "") & _
                               ". An assumption belongs in a cell of its own." & suffix, a1
                End If
            End If
        Next hardcode

        For Each item In analysis("Volatiles")
            Bump stats, "Volatile calls", occurrences
            If Not Capped(caps, "volatile", ws.Name) Then
                AddFinding findings, SEV_LOW, ws.Name, address, "Volatile function", _
                           CStr(item) & "() recalculates on every change anywhere in the workbook." & suffix, a1
            End If
        Next item

        For Each item In analysis("WholeRefs")
            Bump stats, "Whole-column references", occurrences
            If Not Capped(caps, "wholeref", ws.Name) Then
                AddFinding findings, SEV_LOW, ws.Name, address, "Whole-column reference", _
                           CStr(item) & " reads every row in the column, used or not." & suffix, a1
            End If
        Next item

        If analysis("External") > 0 Then
            Bump stats, "External references", occurrences
            If Not Capped(caps, "external", ws.Name) Then
                AddFinding findings, SEV_MEDIUM, ws.Name, address, "Reads another workbook", _
                           "This formula depends on a file that may not be there tomorrow." & suffix, a1
            End If
        End If

        If analysis("IfError") Then Bump stats, "IFERROR wrappers", occurrences
    Next key

    ClearStatusBar
End Sub

'==============================================================================
' Shared helpers
'==============================================================================
Public Sub Bump(ByVal stats As Object, ByVal key As String, Optional ByVal amount As Double = 1)
    If stats Is Nothing Then Exit Sub
    DictIncrement stats, key, amount
End Sub

Public Function CellAddress(ByVal rowNumber As Long, ByVal columnNumber As Long) As String
    CellAddress = ColumnLetter(columnNumber) & CStr(rowNumber)
End Function

'==============================================================================
' Private helpers
'==============================================================================
' Always a 2D array, even for a single cell.
Private Function ReadBlock(ByVal rng As Range, ByVal r1c1Style As Boolean) As Variant
    Dim out As Variant

    On Error GoTo Failed
    If rng.Cells.CountLarge = 1 Then
        ReDim out(1 To 1, 1 To 1)
        If r1c1Style Then out(1, 1) = rng.FormulaR1C1 Else out(1, 1) = rng.Formula
        ReadBlock = out
    ElseIf r1c1Style Then
        ReadBlock = rng.FormulaR1C1
    Else
        ReadBlock = rng.Formula
    End If
    Exit Function
Failed:
End Function

Private Function CellKind(ByVal value As Variant) As Long
    Dim text As String

    If IsError(value) Then CellKind = KIND_TEXT: Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function

    If VarType(value) = vbString Then
        text = CStr(value)
        If Len(text) = 0 Then Exit Function
        If Left$(text, 1) = "=" Then CellKind = KIND_FORMULA: Exit Function
        If IsNumeric(text) Then CellKind = KIND_NUMBER: Exit Function
        CellKind = KIND_TEXT
        Exit Function
    End If

    If IsNumeric(value) Then CellKind = KIND_NUMBER Else CellKind = KIND_TEXT
End Function

' Array and spilled formulas repeat across a range by design, and their R1C1
' shifts with each cell, so comparing them to their neighbours is meaningless.
Private Function IsArrayOrSpillCell(ByVal cell As Range) As Boolean
    Dim spilling As Boolean

    On Error Resume Next

    If cell.HasArray Then IsArrayOrSpillCell = True
    If Err.Number <> 0 Then Err.Clear

    spilling = cell.HasSpill
    If Err.Number <> 0 Then
        Err.Clear                                ' older Excel has no spill ranges
    ElseIf spilling Then
        IsArrayOrSpillCell = True
    End If

    On Error GoTo 0
End Function

' One check should not be able to fill the whole report from one bad sheet.
Private Function Capped(ByVal caps As Object, ByVal checkKey As String, ByVal sheetName As String) As Boolean
    Dim key As String

    key = checkKey & "|" & sheetName
    DictIncrement caps, key, 1
    Capped = (caps(key) > MAX_PER_CHECK)
End Function
