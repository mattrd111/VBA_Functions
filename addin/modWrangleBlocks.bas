Attribute VB_Name = "modWrangleBlocks"
'==============================================================================
' modWrangleBlocks - one table out of a tab that holds several
'------------------------------------------------------------------------------
' The other data-room shape: not one table per sheet, but twelve monthly blocks
' down a single tab, or a management pack with four unrelated tables on it.
'
'   StackBlocksOnSheet   finds the separate blocks by looking for the blank
'                        rows and columns between them, then stacks the ones
'                        that share a header row
'   StackChosenRanges    you pick the blocks yourself, Ctrl-clicking as many as
'                        you like, across sheets and across open workbooks
'
' Both hand off to the same assembler as the sheet and file stackers, so columns
' are matched by header NAME rather than position and the report tells you which
' block was missing what.
'
' Depends on: modWrangleStack, modWrangleShape, modDoctorCommon, modString,
'             modDictionary, modRange, modApp (from \src)
'==============================================================================
Option Explicit

Private Const MAX_SCAN_CELLS As Double = 4000000#

'------------------------------------------------------------------------------
' StackBlocksOnSheet   (menu action)
' Finds the blocks automatically and stacks them.
'------------------------------------------------------------------------------
Public Sub StackBlocksOnSheet()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim blocks As Collection
    Dim everySheet As VbMsgBoxResult
    Dim minBlank As Long
    Dim sources As Collection
    Dim block As Variant
    Dim source As Variant
    Dim modalSignature As String
    Dim matching As Long
    Dim answer As VbMsgBoxResult
    Dim origin As String

    Set wb = TargetWorkbook(False)
    If wb Is Nothing Then Exit Sub

    everySheet = MsgBox("Look for blocks on every sheet in '" & wb.Name & "'?" & vbNewLine & vbNewLine & _
                        "Yes     every sheet" & vbNewLine & _
                        "No      just '" & ActiveSheet.Name & "'" & vbNewLine & _
                        "Cancel  stop", vbYesNoCancel + vbQuestion, DOCTOR_NAME)
    If everySheet = vbCancel Then Exit Sub

    minBlank = AskForNumber("How many blank rows separate one block from the next?" & vbNewLine & _
                            vbNewLine & "1 splits on any blank row. 2 keeps a table together when it " & _
                            "has the odd blank line inside it.", 1)
    If minBlank < 1 Then Exit Sub

    StartTimer
    FastMode True

    Set blocks = New Collection
    If everySheet = vbYes Then
        For Each ws In wb.Worksheets
            AppendBlocks blocks, FindBlocks(ws, minBlank, 2, 1)
        Next ws
        origin = wb.Name & " - every sheet"
    Else
        Set ws = ActiveSheet
        AppendBlocks blocks, FindBlocks(ws, minBlank, 2, 1)
        origin = wb.Name & " - " & ws.Name
    End If

    FastMode False
    ClearStatusBar

    If blocks.count = 0 Then
        MsgBox "No blocks found." & vbNewLine & vbNewLine & _
               "A block is at least two rows with blank rows or columns around it. " & _
               "If your tables run together with no gap between them, pick them " & _
               "yourself with 'Stack ranges I pick'.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    modalSignature = ModalSignature(blocks, matching)

    answer = MsgBox(blocks.count & " block(s) found." & vbNewLine & vbNewLine & _
                    matching & " of them share the most common header row:" & vbNewLine & _
                    "  " & Left$(Replace$(modalSignature, "|", " | "), 200) & vbNewLine & vbNewLine & _
                    "Yes     stack only those " & matching & vbNewLine & _
                    "No      stack all " & blocks.count & " (columns are matched by name, so " & _
                    "the odd one out just adds columns)" & vbNewLine & _
                    "Cancel  list what was found and change nothing", _
                    vbYesNoCancel + vbQuestion, DOCTOR_NAME)
    If answer = vbCancel Then
        ShowReport BlockListReport(blocks, origin, modalSignature), "Blocks found"
        Exit Sub
    End If

    FastMode True
    Set sources = New Collection
    For Each block In blocks
        If answer = vbNo Or StrComp(SignatureOf(block, 1), modalSignature, vbBinaryCompare) = 0 Then
            source = BlockToSource(block, True, BlockName(block))
            If IsArray(source) Then sources.Add source
        End If
    Next block
    FastMode False

    If sources.count = 0 Then
        MsgBox "Nothing matched.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    BuildStack sources, "Stacked blocks", origin
End Sub

'------------------------------------------------------------------------------
' StackChosenRanges   (menu action)
' For tables that run together with no gap, or blocks scattered across sheets
' and workbooks. Ctrl-click as many blocks as you like at each prompt.
'------------------------------------------------------------------------------
Public Sub StackChosenRanges()
    Dim sources As Collection
    Dim chosen As Range
    Dim area As Range
    Dim source As Variant
    Dim hasHeaders As Boolean
    Dim more As Boolean
    Dim origin As String

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open a workbook first.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    hasHeaders = (MsgBox("Does each block start with its own header row?" & vbNewLine & vbNewLine & _
                         "Yes  the first row of every block names its columns" & vbNewLine & _
                         "No   no headers - columns are matched by position instead", _
                         vbYesNo + vbQuestion, DOCTOR_NAME) = vbYes)

    Set sources = New Collection
    more = True

    Do While more
        Set chosen = Nothing
        On Error Resume Next
        Set chosen = Application.InputBox( _
            "Select a block. Ctrl-click to pick several at once." & vbNewLine & vbNewLine & _
            "Blocks chosen so far: " & sources.count & vbNewLine & _
            "Press Cancel when you have them all.", DOCTOR_NAME, Type:=8)
        On Error GoTo 0

        If chosen Is Nothing Then
            more = False
        Else
            For Each area In chosen.Areas
                If area.rows.count >= IIf(hasHeaders, 2, 1) Then
                    source = BlockToSource(area, hasHeaders, BlockName(area))
                    If IsArray(source) Then
                        sources.Add source
                        If Len(origin) = 0 Then origin = area.Worksheet.Parent.Name
                    End If
                End If
            Next area
        End If
    Loop

    If sources.count = 0 Then Exit Sub
    If sources.count = 1 Then
        If MsgBox("Only one block was chosen, so there is nothing to stack it with." & _
                  vbNewLine & vbNewLine & "Carry on anyway?", vbYesNo + vbQuestion, _
                  DOCTOR_NAME) <> vbYes Then Exit Sub
    End If

    StartTimer
    BuildStack sources, "Stacked ranges", origin & " - " & sources.count & " block(s) picked by hand"
End Sub

'==============================================================================
' FindBlocks
'------------------------------------------------------------------------------
' The rectangular islands on a sheet: rows are grouped into bands separated by
' runs of blank rows, then each band is split at any blank column. Mirrored in
' build\test_block_finder.py, which is where the edge cases are exercised.
'==============================================================================
Public Function FindBlocks(ByVal ws As Worksheet, Optional ByVal minBlankRows As Long = 1, _
                           Optional ByVal minRows As Long = 2, _
                           Optional ByVal minCols As Long = 1) As Collection
    Dim result As Collection
    Dim used As Range
    Dim values As Variant
    Dim rowHas() As Boolean
    Dim firstRow As Long, firstCol As Long, rowCount As Long, colCount As Long
    Dim r As Long, c As Long
    Dim bandTop As Long, bandLast As Long, blankRun As Long

    Set result = New Collection
    Set FindBlocks = result
    If ws Is Nothing Then Exit Function
    If minBlankRows < 1 Then minBlankRows = 1

    On Error GoTo Done
    Set used = ws.UsedRange
    If used Is Nothing Then Exit Function

    firstRow = used.Row
    firstCol = used.column
    rowCount = used.rows.count
    colCount = used.Columns.count
    If rowCount < minRows Or colCount < 1 Then Exit Function
    If CDbl(rowCount) * CDbl(colCount) > MAX_SCAN_CELLS Then Exit Function

    values = ReadUsed(used, rowCount, colCount)
    If Not IsArray(values) Then Exit Function

    ReDim rowHas(1 To rowCount)
    For r = 1 To rowCount
        For c = 1 To colCount
            If Not IsBlank(values(r, c)) Then rowHas(r) = True: Exit For
        Next c
    Next r

    r = 1
    Do While r <= rowCount
        If Not rowHas(r) Then
            r = r + 1
        Else
            bandTop = r
            bandLast = r
            blankRun = 0

            Do While r <= rowCount
                If rowHas(r) Then
                    blankRun = 0
                    bandLast = r
                Else
                    blankRun = blankRun + 1
                    If blankRun >= minBlankRows Then Exit Do
                End If
                r = r + 1
            Loop

            SplitBand ws, values, colCount, firstRow, firstCol, bandTop, bandLast, _
                      minRows, minCols, result

            r = bandLast + 1 + minBlankRows
        End If
    Loop

Done:
End Function

'==============================================================================
' Private helpers
'==============================================================================
' One band of rows, cut at every blank column.
Private Sub SplitBand(ByVal ws As Worksheet, ByRef values As Variant, ByVal colCount As Long, _
                      ByVal firstRow As Long, ByVal firstCol As Long, _
                      ByVal bandTop As Long, ByVal bandLast As Long, _
                      ByVal minRows As Long, ByVal minCols As Long, ByVal result As Collection)
    Dim colHas() As Boolean
    Dim c As Long, r As Long
    Dim runLeft As Long

    If bandLast - bandTop + 1 < minRows Then Exit Sub

    ReDim colHas(1 To colCount)
    For c = 1 To colCount
        For r = bandTop To bandLast
            If Not IsBlank(values(r, c)) Then colHas(c) = True: Exit For
        Next r
    Next c

    c = 1
    Do While c <= colCount
        If Not colHas(c) Then
            c = c + 1
        Else
            runLeft = c
            Do While c <= colCount
                If Not colHas(c) Then Exit Do
                c = c + 1
            Loop

            If (c - 1) - runLeft + 1 >= minCols Then
                result.Add ws.Range(ws.Cells(firstRow + bandTop - 1, firstCol + runLeft - 1), _
                                    ws.Cells(firstRow + bandLast - 1, firstCol + c - 2))
            End If
        End If
    Loop
End Sub

Private Function ReadUsed(ByVal used As Range, ByVal rowCount As Long, ByVal colCount As Long) As Variant
    Dim out As Variant

    On Error GoTo Done
    If rowCount = 1 And colCount = 1 Then
        ReDim out(1 To 1, 1 To 1)
        out(1, 1) = used.value
        ReadUsed = out
    Else
        ReadUsed = used.value
    End If
Done:
End Function

Private Sub AppendBlocks(ByVal target As Collection, ByVal found As Collection)
    Dim item As Variant

    If found Is Nothing Then Exit Sub
    For Each item In found
        target.Add item
    Next item
End Sub

' Array(sourceName, sheetName, headers(), data()) - the shape BuildStack wants.
Private Function BlockToSource(ByVal block As Range, ByVal hasHeaderRow As Boolean, _
                               ByVal sourceName As String) As Variant
    Dim values As Variant
    Dim headers() As String
    Dim data As Variant
    Dim rowCount As Long, colCount As Long
    Dim r As Long, c As Long

    values = RangeToArray(block)
    If Not IsArray(values) Then Exit Function

    rowCount = UBound(values, 1)
    colCount = UBound(values, 2)

    ReDim headers(1 To colCount)
    For c = 1 To colCount
        If hasHeaderRow Then
            headers(c) = CleanText(values(1, c))
            If Len(headers(c)) = 0 Then headers(c) = "Column " & c
        Else
            headers(c) = "Column " & c
        End If
    Next c

    If hasHeaderRow Then
        If rowCount < 2 Then Exit Function
        ReDim data(1 To rowCount - 1, 1 To colCount)
        For r = 2 To rowCount
            For c = 1 To colCount
                data(r - 1, c) = values(r, c)
            Next c
        Next r
    Else
        data = values
    End If

    BlockToSource = Array(sourceName, block.Worksheet.Name, headers, data)
End Function

Private Function BlockName(ByVal block As Range) As String
    BlockName = block.Worksheet.Name & "!" & block.Address(False, False)
End Function

' The header row as one comparable string.
Private Function SignatureOf(ByVal block As Range, ByVal headerRow As Long) As String
    Dim values As Variant
    Dim c As Long
    Dim parts As String

    On Error GoTo Done
    values = RangeToArray(block.rows(headerRow))
    If Not IsArray(values) Then Exit Function

    For c = 1 To UBound(values, 2)
        parts = parts & IIf(c = 1, "", "|") & UCase$(CleanText(values(1, c)))
    Next c
    SignatureOf = parts
Done:
End Function

' The header row that the most blocks share.
Private Function ModalSignature(ByVal blocks As Collection, ByRef matching As Long) As String
    Dim counts As Object
    Dim block As Variant
    Dim signature As String
    Dim key As Variant

    Set counts = NewDictionary(False)

    For Each block In blocks
        signature = SignatureOf(block, 1)
        If Len(signature) > 0 Then DictIncrement counts, signature, 1
    Next block

    matching = 0
    For Each key In counts.keys
        If counts(key) > matching Then
            matching = counts(key)
            ModalSignature = CStr(key)
        End If
    Next key
End Function

Private Function BlockListReport(ByVal blocks As Collection, ByVal origin As String, _
                                 ByVal modalSignature As String) As Collection
    Dim rows As Collection
    Dim block As Variant
    Dim signature As String

    Set rows = NewReport()
    ReportRow rows, DOCTOR_NAME & " - Blocks found"
    ReportRow rows, "In", origin
    ReportRow rows, "Found", blocks.count

    ReportHeading rows, "Blocks"
    ReportRow rows, "Sheet", "Range", "Rows", "Columns", "Same headers as the rest", "Header row"

    For Each block In blocks
        signature = SignatureOf(block, 1)
        ReportRow rows, block.Worksheet.Name, block.Address(False, False), _
                  block.rows.count, block.Columns.count, _
                  IIf(StrComp(signature, modalSignature, vbBinaryCompare) = 0, "Yes", "No"), _
                  Left$(Replace$(signature, "|", " | "), 250)
    Next block

    ReportHeading rows, "Nothing was changed"
    ReportRow rows, "Note", "This is a list of what is there. Run the tool again and " & _
              "choose Yes or No to stack them."

    Set BlockListReport = rows
End Function
