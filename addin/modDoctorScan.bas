Attribute VB_Name = "modDoctorScan"
'==============================================================================
' modDoctorScan - works out what a workbook actually refers to
'------------------------------------------------------------------------------
' Builds a set of every identifier that appears anywhere a defined name could
' be used: cell formulas, conditional formatting, data validation, chart series,
' form controls, pivot sources and the RefersTo of other names.
'
' It is a text scan, so it is deliberately generous - a name is only reported as
' unused when nothing anywhere looks like a reference to it. That errs towards
' keeping names rather than deleting live ones.
'
' Depends on: modDictionary, modApp (from \src)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' BuildUsageIndex
' Dictionary of every upper-cased identifier found in the workbook. A name's own
' RefersTo does not count as a use of itself.
'------------------------------------------------------------------------------
Public Function BuildUsageIndex(ByVal wb As Workbook) As Object
    Dim tokens As Object
    Dim ws As Worksheet
    Dim cht As Chart
    Dim nm As Name
    Dim done As Long

    Set tokens = NewDictionary(True)
    Set BuildUsageIndex = tokens
    If wb Is Nothing Then Exit Function

    For Each ws In wb.Worksheets
        done = done + 1
        ShowProgress done, wb.Worksheets.count, "Scanning " & ws.Name
        ScanFormulas ws, tokens
        ScanConditionalFormats ws, tokens
        ScanValidation ws, tokens
        ScanShapes ws, tokens
        ScanEmbeddedCharts ws, tokens
        ScanPivotTables ws, tokens
    Next ws

    On Error Resume Next
    For Each cht In wb.Charts
        ScanChartSeries cht, tokens
    Next cht
    On Error GoTo 0

    For Each nm In wb.Names
        ScanName nm, tokens
    Next nm

    ClearStatusBar
End Function

'------------------------------------------------------------------------------
' AddTokens
' Pulls every identifier-shaped run of characters out of a formula, ignoring
' anything inside a quoted string.
'------------------------------------------------------------------------------
Public Sub AddTokens(ByVal formulaText As String, ByVal tokens As Object, _
                     Optional ByVal excludeToken As String = "")
    Dim i As Long, n As Long, start As Long
    Dim ch As String
    Dim token As String
    Dim insideString As Boolean

    n = Len(formulaText)
    If n = 0 Then Exit Sub

    i = 1
    Do While i <= n
        ch = Mid$(formulaText, i, 1)

        If insideString Then
            If ch = """" Then insideString = False
        ElseIf ch = """" Then
            insideString = True
        ElseIf IsNameCharacter(ch) Then
            start = i
            Do While i <= n
                If Not IsNameCharacter(Mid$(formulaText, i, 1)) Then Exit Do
                i = i + 1
            Loop
            token = UCase$(Mid$(formulaText, start, i - start))
            If Len(token) > 0 And StrComp(token, excludeToken, vbTextCompare) <> 0 Then
                If Not tokens.Exists(token) Then tokens.Add token, True
            End If
            i = i - 1
        End If

        i = i + 1
    Loop
End Sub

'------------------------------------------------------------------------------
' LocalNameOf / ScopeOf / IsReservedName
'------------------------------------------------------------------------------
Public Function LocalNameOf(ByVal fullName As String) As String
    Dim p As Long

    p = InStrRev(fullName, "!")
    If p > 0 Then
        LocalNameOf = Mid$(fullName, p + 1)
    Else
        LocalNameOf = fullName
    End If
End Function

Public Function ScopeOf(ByVal nm As Name) As String
    On Error GoTo Unknown
    If TypeOf nm.Parent Is Worksheet Then
        ScopeOf = nm.Parent.Name
    Else
        ScopeOf = "Workbook"
    End If
    Exit Function
Unknown:
    ScopeOf = "?"
End Function

' Names Excel creates and relies on. Deleting these breaks printing, filters or
' the advanced filter, so the add-in never touches them.
Public Function IsReservedName(ByVal localName As String) As Boolean
    Select Case UCase$(localName)
        Case "PRINT_AREA", "PRINT_TITLES", "_FILTERDATABASE", "FILTERDATABASE", _
             "CRITERIA", "EXTRACT", "DATABASE", "CONSOLIDATE_AREA", "SHEET_TITLE", _
             "AUTO_OPEN", "AUTO_CLOSE", "AUTO_ACTIVATE", "AUTO_DEACTIVATE", _
             "_XLFN", "_XLNM"
            IsReservedName = True
    End Select

    If Left$(UCase$(localName), 5) = "_XLFN" Then IsReservedName = True
    If Left$(UCase$(localName), 5) = "_XLNM" Then IsReservedName = True
End Function

'==============================================================================
' Private helpers - one per place a formula can hide
'==============================================================================
Private Sub ScanFormulas(ByVal ws As Worksheet, ByVal tokens As Object)
    Dim formulaCells As Range
    Dim area As Range
    Dim values As Variant
    Dim r As Long, c As Long

    On Error Resume Next
    Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If formulaCells Is Nothing Then Exit Sub

    For Each area In formulaCells.Areas
        If area.Cells.CountLarge = 1 Then
            AddTokens CStr(area.Formula), tokens
        Else
            values = area.Formula
            For r = LBound(values, 1) To UBound(values, 1)
                For c = LBound(values, 2) To UBound(values, 2)
                    AddTokens CStr(values(r, c)), tokens
                Next c
            Next r
        End If
    Next area
End Sub

Private Sub ScanConditionalFormats(ByVal ws As Worksheet, ByVal tokens As Object)
    Dim rules As Object
    Dim i As Long
    Dim ruleCount As Long

    On Error Resume Next
    Set rules = ws.Cells.FormatConditions
    ruleCount = rules.count
    If Err.Number <> 0 Then Err.Clear: Exit Sub

    For i = 1 To ruleCount
        AddTokens CStr(rules(i).Formula1), tokens
        Err.Clear
        AddTokens CStr(rules(i).Formula2), tokens
        Err.Clear
    Next i
    On Error GoTo 0
End Sub

Private Sub ScanValidation(ByVal ws As Worksheet, ByVal tokens As Object)
    Dim validationCells As Range
    Dim area As Range

    On Error Resume Next
    Set validationCells = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    If validationCells Is Nothing Then Exit Sub

    For Each area In validationCells.Areas
        AddTokens CStr(area.Cells(1, 1).Validation.Formula1), tokens
        Err.Clear
        AddTokens CStr(area.Cells(1, 1).Validation.Formula2), tokens
        Err.Clear
    Next area
    On Error GoTo 0
End Sub

Private Sub ScanShapes(ByVal ws As Worksheet, ByVal tokens As Object)
    Dim shp As Shape

    On Error Resume Next
    For Each shp In ws.Shapes
        AddTokens CStr(shp.ControlFormat.ListFillRange), tokens
        Err.Clear
        AddTokens CStr(shp.ControlFormat.LinkedCell), tokens
        Err.Clear
        AddTokens CStr(shp.DrawingObject.Formula), tokens
        Err.Clear
    Next shp
    On Error GoTo 0
End Sub

Private Sub ScanEmbeddedCharts(ByVal ws As Worksheet, ByVal tokens As Object)
    Dim chartObject As ChartObject

    On Error Resume Next
    For Each chartObject In ws.ChartObjects
        ScanChartSeries chartObject.Chart, tokens
    Next chartObject
    On Error GoTo 0
End Sub

Private Sub ScanChartSeries(ByVal cht As Chart, ByVal tokens As Object)
    Dim i As Long
    Dim seriesCount As Long

    On Error Resume Next
    seriesCount = cht.SeriesCollection.count
    If Err.Number <> 0 Then Err.Clear: Exit Sub

    For i = 1 To seriesCount
        AddTokens CStr(cht.SeriesCollection(i).Formula), tokens
        Err.Clear
    Next i
    On Error GoTo 0
End Sub

Private Sub ScanPivotTables(ByVal ws As Worksheet, ByVal tokens As Object)
    Dim pt As PivotTable

    On Error Resume Next
    For Each pt In ws.PivotTables
        AddTokens CStr(pt.SourceData), tokens
        Err.Clear
    Next pt
    On Error GoTo 0
End Sub

Private Sub ScanName(ByVal nm As Name, ByVal tokens As Object)
    Dim refersTo As String
    Dim localName As String

    On Error Resume Next
    localName = UCase$(LocalNameOf(nm.Name))
    refersTo = nm.refersTo
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    On Error GoTo 0

    ' A name referring to itself is not a use of itself.
    AddTokens refersTo, tokens, localName
End Sub

Private Function IsNameCharacter(ByVal ch As String) As Boolean
    Select Case ch
        Case "A" To "Z", "a" To "z", "0" To "9", "_", ".", "\", "?"
            IsNameCharacter = True
        Case Else
            IsNameCharacter = (AscW(ch) > 127)
    End Select
End Function
