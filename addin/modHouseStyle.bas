Attribute VB_Name = "modHouseStyle"
'==============================================================================
' modHouseStyle - Alpha FMC house style for Excel
'------------------------------------------------------------------------------
' The palette and type read straight out of the Alpha FMC 2026 PowerPoint theme,
' so a table built here sits next to a slide without anyone reaching for the
' eyedropper.
'
'   ink            2A2723   text, and the near-black everything else sits on
'   violet         503AF5   the primary
'   violet wash    EDEBFD   what their table headers are filled with
'   paper          F0F5EB   panel and banding fill
'   sage           BBC1B2   the quiet neutral
'   Quire Sans     10pt body, 9pt in a header
'
' Their tables carry no borders at all - a pale fill on the header row does the
' work, and the header is not bold. That is followed here rather than corrected.
'
' The number format ladders are NOT from the template, which does not specify
' any. They are the usual consulting conventions and are meant to be edited -
' they are all in one place at the top of the module.
'
' Depends on: modDoctorCommon, modRange, modWorkbook, modApp (from \src)
'==============================================================================
Option Explicit

Public Const HOUSE_FONT As String = "Quire Sans"
Private Const BODY_SIZE As Double = 10
Private Const HEADER_SIZE As Double = 9

' Edit these to taste. "|" separates the steps, and "~" separates each format
' from what to call it in the status bar. They are functions rather than
' constants so the pound sign can be built with Chr$ and the file stays plain
' ASCII, which is what makes it import cleanly on any machine.
Private Function NumberLadder() As String
    NumberLadder = "#,##0;(#,##0)~0 dp|" & _
                   "#,##0.0;(#,##0.0)~1 dp|" & _
                   "#,##0.00;(#,##0.00)~2 dp|" & _
                   "#,##0,;(#,##0,)~thousands|" & _
                   "#,##0.0,,;(#,##0.0,,)~millions|" & _
                   Pound() & "#,##0.0,,;(" & Pound() & "#,##0.0,,)~" & Pound() & "m"
End Function

Private Function PercentLadder() As String
    PercentLadder = "0%;(0%)~0 dp|" & _
                    "0.0%;(0.0%)~1 dp|" & _
                    "0.00%;(0.00%)~2 dp"
End Function

Private Function DateLadder() As String
    DateLadder = "dd mmm yy~dd mmm yy|" & _
                 "mmm-yy~mmm-yy|" & _
                 "dd/mm/yyyy~dd/mm/yyyy|" & _
                 "yyyy-mm-dd~ISO"
End Function

Private Function Pound() As String
    Pound = Chr$(163)
End Function

'==============================================================================
' The palette
'------------------------------------------------------------------------------
' AlphaColour("violet") and so on, so nothing in this repository has a bare hex
' code buried in it.
'==============================================================================
Public Function AlphaColour(ByVal name As String) As Long
    Select Case UCase$(Trim$(name))
        Case "INK":         AlphaColour = RGB(&H2A, &H27, &H23)
        Case "WHITE":       AlphaColour = RGB(&HFF, &HFF, &HFF)
        Case "VIOLET":      AlphaColour = RGB(&H50, &H3A, &HF5)
        Case "VIOLET2":     AlphaColour = RGB(&H6A, &H57, &HF6)
        Case "VIOLET3":     AlphaColour = RGB(&H8D, &H7E, &HF8)
        Case "VIOLET4":     AlphaColour = RGB(&HB8, &HB0, &HFA)
        Case "VIOLET5":     AlphaColour = RGB(&HDC, &HD7, &HFC)
        Case "VIOLETWASH":  AlphaColour = RGB(&HED, &HEB, &HFD)
        Case "SAGE":        AlphaColour = RGB(&HBB, &HC1, &HB2)
        Case "SAGELIGHT":   AlphaColour = RGB(&HE0, &HEB, &HD4)
        Case "PAPER":       AlphaColour = RGB(&HF0, &HF5, &HEB)
        Case "SKY":         AlphaColour = RGB(&H7E, &HA8, &HFF)
        Case "ORANGE":      AlphaColour = RGB(&HF2, &H6B, &H43)
        Case "TEAL":        AlphaColour = RGB(&H0, &HAE, &HCB)
        Case "MINT":        AlphaColour = RGB(&H58, &HF7, &HB0)
        Case "SLATE":       AlphaColour = RGB(&H26, &H33, &H3C)
        Case "CHART1":      AlphaColour = RGB(&H1E, &H29, &H99)
        Case "CHART2":      AlphaColour = RGB(&H60, &H85, &HDC)
        Case "CHART3":      AlphaColour = RGB(&H64, &HB0, &HE2)
        Case "CHART4":      AlphaColour = RGB(&H4B, &HA3, &H79)
        Case "CHART5":      AlphaColour = RGB(&H79, &HC7, &H92)
        Case Else:          AlphaColour = RGB(&H2A, &H27, &H23)
    End Select
End Function

'==============================================================================
' Number format cycling
'------------------------------------------------------------------------------
' The most-pressed key in any formatting add-in. Each call moves the selection
' one step along its ladder and says where it landed in the status bar.
'==============================================================================
Public Sub CycleNumberFormat()
    StepLadder NumberLadder(), "Number"
End Sub

Public Sub CyclePercentFormat()
    StepLadder PercentLadder(), "Percent"
End Sub

Public Sub CycleDateFormat()
    StepLadder DateLadder(), "Date"
End Sub

'==============================================================================
' House style
'==============================================================================
'------------------------------------------------------------------------------
' StyleAlphaTable   (menu action)
' Select the block including its header row.
'------------------------------------------------------------------------------
Public Sub StyleAlphaTable()
    Dim target As Range
    Dim body As Range

    Set target = SelectedBlock()
    If target Is Nothing Then Exit Sub

    FastMode True
    On Error GoTo Done

    With target
        .Font.name = HOUSE_FONT
        .Font.Size = BODY_SIZE
        .Font.Color = AlphaColour("ink")
        .Font.Bold = False
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
        .VerticalAlignment = xlBottom
        .WrapText = False
    End With

    ApplyHeaderStyle target.rows(1)

    If target.rows.count > 1 Then
        Set body = target.Offset(1).Resize(target.rows.count - 1)
        BandRows body
    End If

    FitColumns target, 42, 9

Done:
    FastMode False
End Sub

'------------------------------------------------------------------------------
' StyleAlphaHeader   (menu action)
' Just the header treatment, for a row you have already laid out.
'------------------------------------------------------------------------------
Public Sub StyleAlphaHeader()
    Dim target As Range

    Set target = SelectedBlock()
    If target Is Nothing Then Exit Sub

    FastMode True
    ApplyHeaderStyle target
    FastMode False
End Sub

'------------------------------------------------------------------------------
' StyleAlphaTotal   (menu action)
' A rule above and bold, which is a convention rather than anything the template
' says - it has no tables with totals in it.
'------------------------------------------------------------------------------
Public Sub StyleAlphaTotal()
    Dim target As Range

    Set target = SelectedBlock()
    If target Is Nothing Then Exit Sub

    FastMode True
    With target
        .Font.name = HOUSE_FONT
        .Font.Size = BODY_SIZE
        .Font.Bold = True
        .Font.Color = AlphaColour("ink")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).Color = AlphaColour("ink")
    End With
    FastMode False
End Sub

'------------------------------------------------------------------------------
' StyleAlphaSheet   (menu action)
' Sets the whole sheet on house type and turns the gridlines off, which is what
' makes a sheet screenshot cleanly into a slide.
'------------------------------------------------------------------------------
Public Sub StyleAlphaSheet()
    Dim ws As Worksheet
    Dim previous As Boolean

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open a workbook first.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If
    Set ws = ActiveSheet

    If Not Confirm("Set every cell on '" & ws.Name & "' in " & HOUSE_FONT & " at " & _
                   BODY_SIZE & "pt, and turn the gridlines off?" & vbNewLine & vbNewLine & _
                   "Fills, borders and number formats are left alone.") Then Exit Sub

    FastMode True
    On Error Resume Next

    ws.Cells.Font.name = HOUSE_FONT
    ws.Cells.Font.Size = BODY_SIZE
    ws.Cells.Font.Color = AlphaColour("ink")
    ws.Tab.Color = AlphaColour("violet")

    previous = Application.ScreenUpdating
    Application.ScreenUpdating = False
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    Application.ScreenUpdating = previous

    On Error GoTo 0
    FastMode False
End Sub

'==============================================================================
' Input and formula colouring
'------------------------------------------------------------------------------
' The modelling convention, not the brand: blue means someone typed it, ink
' means it calculates, green means it comes from another sheet and red means it
' comes from another workbook. The point of it is that a reviewer can see where
' the assumptions are without clicking a single cell.
'
' The colours are the conventional ones rather than Alpha's, because their whole
' value is that everyone already reads them the same way.
'==============================================================================
Public Sub ColourInputsAndFormulas()
    Dim target As Range
    Dim ws As Worksheet
    Dim formulas As Variant
    Dim r As Long, c As Long
    Dim firstRow As Long, firstCol As Long
    Dim runStart As Long, runKind As Long, kind As Long
    Dim counts(0 To 3) As Long

    Set target = SelectedBlock()
    If target Is Nothing Then Exit Sub
    Set ws = target.Worksheet

    If target.Cells.CountLarge > 400000 Then
        MsgBox "That is a very large selection. Narrow it to the block you want coloured.", _
               vbExclamation, DOCTOR_NAME
        Exit Sub
    End If

    If Not Confirm("Recolour the text in " & target.Address(False, False) & " by what each " & _
                   "cell is?" & vbNewLine & vbNewLine & _
                   "blue    a number someone typed" & vbNewLine & _
                   "black   a formula on this sheet" & vbNewLine & _
                   "green   a formula reading another sheet" & vbNewLine & _
                   "red     a formula reading another workbook" & vbNewLine & vbNewLine & _
                   "Text labels are left as they are.") Then Exit Sub

    StartTimer
    FastMode True
    On Error GoTo Done

    formulas = target.Formula
    If Not IsArray(formulas) Then
        ReDim formulas(1 To 1, 1 To 1)
        formulas(1, 1) = target.Formula
    End If

    firstRow = target.Row
    firstCol = target.column

    For r = 1 To UBound(formulas, 1)
        runStart = 0
        runKind = -1

        For c = 1 To UBound(formulas, 2)
            kind = CellKindOf(formulas(r, c))

            If kind <> runKind Then
                If runKind >= 0 Then
                    PaintRun ws, firstRow + r - 1, firstCol + runStart - 1, firstCol + c - 2, runKind
                    counts(runKind) = counts(runKind) + (c - runStart)
                End If
                runStart = c
                runKind = kind
            End If
        Next c

        If runKind >= 0 Then
            PaintRun ws, firstRow + r - 1, firstCol + runStart - 1, _
                     firstCol + UBound(formulas, 2) - 1, runKind
            counts(runKind) = counts(runKind) + (UBound(formulas, 2) - runStart + 1)
        End If
    Next r

Done:
    FastMode False

    MsgBox "Coloured in " & ElapsedText & "." & vbNewLine & vbNewLine & _
           Format$(counts(0), "#,##0") & " typed numbers" & vbNewLine & _
           Format$(counts(1), "#,##0") & " formulas on this sheet" & vbNewLine & _
           Format$(counts(2), "#,##0") & " reading another sheet" & vbNewLine & _
           Format$(counts(3), "#,##0") & " reading another workbook", _
           vbInformation, DOCTOR_NAME
End Sub

'==============================================================================
' Charts and reference
'==============================================================================
'------------------------------------------------------------------------------
' ApplyAlphaChartColours   (menu action)
' Select a chart first. Series are recoloured in the order the template uses
' them, which nobody remembers the hex codes for.
'------------------------------------------------------------------------------
Public Sub ApplyAlphaChartColours()
    Dim cht As Chart
    Dim i As Long
    Dim seriesCount As Long
    Dim colours As Variant

    On Error Resume Next
    Set cht = ActiveChart
    On Error GoTo 0

    If cht Is Nothing Then
        MsgBox "Click on a chart first.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    colours = Array("chart1", "chart2", "chart3", "chart4", "chart5", "violet", "orange", "teal")

    On Error Resume Next
    seriesCount = cht.SeriesCollection.count

    For i = 1 To seriesCount
        ' Set the fill and the line to the same colour, so this works on a bar
        ' chart and a line chart without having to ask which it is.
        With cht.SeriesCollection(i).Format
            .Fill.Visible = msoTrue
            .Fill.Solid
            .Fill.ForeColor.RGB = AlphaColour(CStr(colours((i - 1) Mod (UBound(colours) + 1))))
            .Line.ForeColor.RGB = AlphaColour(CStr(colours((i - 1) Mod (UBound(colours) + 1))))
        End With
    Next i

    cht.ChartArea.Font.name = HOUSE_FONT
    cht.ChartArea.Font.Size = HEADER_SIZE
    cht.ChartArea.Font.Color = AlphaColour("ink")
    cht.ChartArea.Format.Line.Visible = msoFalse
    On Error GoTo 0

    MsgBox seriesCount & " series recoloured.", vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' ShowAlphaPalette   (menu action)
' A swatch sheet: every colour, its hex, and what it is for.
'------------------------------------------------------------------------------
Public Sub ShowAlphaPalette()
    Dim names As Variant, roles As Variant
    Dim data As Variant
    Dim book As Workbook
    Dim ws As Worksheet
    Dim i As Long

    names = Array("ink", "violet", "violet2", "violet3", "violet4", "violet5", "violetwash", _
                  "paper", "sagelight", "sage", "slate", "sky", _
                  "orange", "teal", "mint", _
                  "chart1", "chart2", "chart3", "chart4", "chart5")

    roles = Array("Text, and the near-black under everything", _
                  "The primary", "Tint", "Tint", "Tint", "Tint", _
                  "Table header fill - what the template uses", _
                  "Panel and row banding", "Light panel", "The quiet neutral", "Dark panel", "Accent blue", _
                  "Secondary - highlight", "Secondary", "Secondary", _
                  "Chart series 1", "Chart series 2", "Chart series 3", "Chart series 4", "Chart series 5")

    ReDim data(1 To UBound(names) + 3, 1 To 4)
    data(1, 1) = "Alpha FMC 2026 palette"
    data(2, 1) = "Name"
    data(2, 2) = "Swatch"
    data(2, 3) = "Hex"
    data(2, 4) = "What it is for"

    For i = 0 To UBound(names)
        data(i + 3, 1) = CStr(names(i))
        data(i + 3, 3) = HexOf(AlphaColour(CStr(names(i))))
        data(i + 3, 4) = CStr(roles(i))
    Next i

    Set book = Application.Workbooks.Add
    Set ws = DumpToSheet(data, "Alpha palette", book, False)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    For i = 0 To UBound(names)
        ws.Cells(i + 3, 2).Interior.Color = AlphaColour(CStr(names(i)))
    Next i

    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2:D2").Font.Bold = True
    ws.Range("A1:D" & (UBound(names) + 3)).Font.name = HOUSE_FONT
    ws.Range("A1:D" & (UBound(names) + 3)).Font.Color = AlphaColour("ink")
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 46
    ws.Range("C:C").Font.name = "Consolas"
    On Error GoTo 0

    MsgBox "The palette, read out of the Alpha FMC 2026 PowerPoint theme." & vbNewLine & vbNewLine & _
           "Body type is " & HOUSE_FONT & " at " & BODY_SIZE & "pt, headers at " & HEADER_SIZE & _
           "pt and not bold, and their tables carry no borders at all.", _
           vbInformation, DOCTOR_NAME
End Sub

'==============================================================================
' Private helpers
'==============================================================================
' AutoFit on the block's own contents, then cap the width so one long note
' cannot make a column half a screen wide.
Private Sub FitColumns(ByVal target As Range, ByVal maxWidth As Double, ByVal minWidth As Double)
    Dim c As Long

    On Error Resume Next
    target.Columns.AutoFit
    For c = 1 To target.Columns.count
        If target.Columns(c).ColumnWidth > maxWidth Then target.Columns(c).ColumnWidth = maxWidth
        If target.Columns(c).ColumnWidth < minWidth Then target.Columns(c).ColumnWidth = minWidth
    Next c
    On Error GoTo 0
End Sub

Private Sub ApplyHeaderStyle(ByVal header As Range)
    If header Is Nothing Then Exit Sub

    With header
        .Font.name = HOUSE_FONT
        .Font.Size = HEADER_SIZE
        .Font.Bold = False                      ' the template's headers are not bold
        .Font.Color = AlphaColour("ink")
        .Interior.Color = AlphaColour("violetwash")
        .VerticalAlignment = xlBottom
        .WrapText = True
    End With
End Sub

' Banding as a rule rather than a fill, so it survives rows being inserted.
Private Sub BandRows(ByVal body As Range)
    Dim rule As FormatCondition
    Dim expression As String
    Dim i As Long

    expression = "=MOD(ROW()-" & body.Row & ",2)=1"

    On Error Resume Next

    ' Clear our own banding from a previous run, and nothing else - a rule
    ' someone set up deliberately is not ours to delete.
    For i = body.FormatConditions.count To 1 Step -1
        If body.FormatConditions(i).Type = xlExpression Then
            If StrComp(CStr(body.FormatConditions(i).Formula1), expression, vbTextCompare) = 0 Then
                body.FormatConditions(i).Delete
            End If
        End If
        Err.Clear
    Next i

    Set rule = body.FormatConditions.Add(Type:=xlExpression, Formula1:=expression)
    If rule Is Nothing Then Exit Sub
    rule.Interior.Color = AlphaColour("paper")
    On Error GoTo 0
End Sub

Private Sub StepLadder(ByVal ladder As String, ByVal what As String)
    Dim ladderSteps As Variant
    Dim target As Range
    Dim current As Variant
    Dim i As Long, nextStep As Long
    Dim parts As Variant

    Set target = SelectedBlock()
    If target Is Nothing Then Exit Sub

    ladderSteps = Split(ladder, "|")

    On Error Resume Next
    current = target.NumberFormat
    On Error GoTo 0

    nextStep = 0                                 ' mixed or unrecognised starts at the top
    If Not IsNull(current) Then
        For i = 0 To UBound(ladderSteps)
            If StrComp(CStr(current), Split(CStr(ladderSteps(i)), "~")(0), vbTextCompare) = 0 Then
                nextStep = (i + 1) Mod (UBound(ladderSteps) + 1)
                Exit For
            End If
        Next i
    End If

    parts = Split(CStr(ladderSteps(nextStep)), "~")
    On Error Resume Next
    target.NumberFormat = parts(0)
    Application.StatusBar = what & " format: " & parts(1) & "    " & parts(0)
    On Error GoTo 0
End Sub

' 0 typed number, 1 formula here, 2 formula from another sheet,
' 3 formula from another workbook, -1 leave it alone
Private Function CellKindOf(ByVal value As Variant) As Long
    Dim text As String

    CellKindOf = -1
    If IsError(value) Then Exit Function
    If IsEmpty(value) Or IsNull(value) Then Exit Function

    If VarType(value) = vbString Then
        text = CStr(value)
        If Len(text) = 0 Then Exit Function

        If Left$(text, 1) = "=" Then
            If InStr(text, "[") > 0 Then
                CellKindOf = 3
            ElseIf InStr(text, "!") > 0 Then
                CellKindOf = 2
            Else
                CellKindOf = 1
            End If
            Exit Function
        End If

        If IsNumeric(text) Then CellKindOf = 0
        Exit Function                            ' text labels keep whatever they had
    End If

    If IsNumeric(value) Then CellKindOf = 0
End Function

Private Sub PaintRun(ByVal ws As Worksheet, ByVal r As Long, ByVal firstCol As Long, _
                     ByVal lastCol As Long, ByVal kind As Long)
    Dim colour As Long

    If kind < 0 Or lastCol < firstCol Then Exit Sub

    Select Case kind
        Case 0: colour = RGB(0, 0, 255)          ' typed
        Case 1: colour = AlphaColour("ink")      ' calculated here
        Case 2: colour = RGB(0, 128, 0)          ' another sheet
        Case 3: colour = RGB(192, 0, 0)          ' another workbook
        Case Else: Exit Sub
    End Select

    On Error Resume Next
    ws.Range(ws.Cells(r, firstCol), ws.Cells(r, lastCol)).Font.Color = colour
    On Error GoTo 0
End Sub

Private Function SelectedBlock() As Range
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

    If target.Areas.count > 1 Then
        MsgBox "Select one continuous block.", vbExclamation, DOCTOR_NAME
        Exit Function
    End If

    ' A whole-column selection would otherwise mean a million empty cells.
    If target.Cells.CountLarge > 500000 Then
        On Error Resume Next
        Set target = Application.Intersect(target, target.Worksheet.UsedRange)
        On Error GoTo 0
        If target Is Nothing Then Exit Function
    End If

    Set SelectedBlock = target
End Function

Private Function HexOf(ByVal colour As Long) As String
    HexOf = "#" & Right$("0" & Hex$(colour And &HFF), 2) & _
                  Right$("0" & Hex$((colour \ &H100) And &HFF), 2) & _
                  Right$("0" & Hex$((colour \ &H10000) And &HFF), 2)
End Function
