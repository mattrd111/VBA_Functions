Option Explicit


Sub OverrideYellowCells()
    ' Freeze every yellow cell at its current value so you can type over it.
    Dim wsBank As Worksheet, wsMain As Worksheet
    Dim i As Long, lastRow As Long
    Dim cellRef As String
    Dim wasEnabled As Boolean

    Set wsBank = ThisWorkbook.Sheets("FormulaBank")
    Set wsMain = ThisWorkbook.Sheets("Fund View")

    lastRow = wsBank.Cells(wsBank.Rows.Count, "A").End(xlUp).Row

    wasEnabled = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    For i = 2 To lastRow
        cellRef = wsBank.Cells(i, 1).Value
        With wsMain.Range(cellRef)
            .Value = .Value
        End With
    Next i
    Application.ScreenUpdating = True
    Application.EnableEvents = wasEnabled

    MsgBox "Yellow cells converted to static values." & vbCrLf & _
           "You can now edit any of them by hand.", vbInformation
End Sub

Sub RestoreYellowFormulas()
    ' Formulas restored, overrides discarded, red reset to yellow.

    Const YELLOW_FILL As Long = 13434879   ' RGB(255,255,204)
    Const CHUNK_REFS  As Long = 20

    Dim wsBank As Worksheet, wsMain As Worksheet
    Dim bank As Variant
    Dim i As Long, lastRow As Long
    Dim refs As String, inChunk As Long
    Dim savedCalc As XlCalculation
    Dim savedEvents As Boolean, savedScreen As Boolean

    Set wsBank = ThisWorkbook.Sheets("FormulaBank")
    Set wsMain = ThisWorkbook.Sheets("Fund View")

    lastRow = wsBank.Cells(wsBank.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    bank = wsBank.Range("A2:C" & lastRow).Value

    savedCalc = Application.Calculation
    savedEvents = Application.EnableEvents
    savedScreen = Application.ScreenUpdating

    On Error GoTo CleanUp
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For i = 1 To UBound(bank, 1)
        If Len(bank(i, 1)) > 0 Then
            If bank(i, 3) = "Y" Then
                wsMain.Range(CStr(bank(i, 1))).FormulaArray = CStr(bank(i, 2))
            Else
                wsMain.Range(CStr(bank(i, 1))).Formula = CStr(bank(i, 2))
            End If

         
            If Len(refs) > 0 Then refs = refs & ","
            refs = refs & CStr(bank(i, 1))
            inChunk = inChunk + 1
            If inChunk = CHUNK_REFS Then
                ApplyYellow wsMain, refs, YELLOW_FILL
                refs = "": inChunk = 0
            End If
        End If
    Next i
    If Len(refs) > 0 Then ApplyYellow wsMain, refs, YELLOW_FILL

CleanUp:
    Application.ScreenUpdating = savedScreen
    Application.EnableEvents = savedEvents
    Application.Calculation = savedCalc
    If Err.Number <> 0 Then
        MsgBox "Stopped at FormulaBank row " & i + 1 & ": " & Err.Description, vbExclamation
    Else
        MsgBox "Formulas restored on all yellow cells.", vbInformation
    End If
End Sub

Private Sub ApplyYellow(ws As Worksheet, refs As String, fillColour As Long)
    With ws.Range(refs)
        .Interior.Color = fillColour
        .Font.ColorIndex = xlAutomatic
        .ClearComments
    End With
End Sub



Private Sub BlankChunk(ws As Worksheet, refs As String)
    Dim ar As Range
    For Each ar In ws.Range(refs).Areas
        ar.Value = ""
    Next ar
End Sub

Sub AddNewFund()

    Dim wsMain As Worksheet, wsFund As Worksheet, wsPosition As Worksheet
    Dim newId As String
    Dim existingMatch As Variant
    Dim wasEnabled As Boolean

    Set wsMain = ThisWorkbook.Sheets("Fund View")
    Set wsFund = ThisWorkbook.Sheets("fund")
    Set wsPosition = ThisWorkbook.Sheets("position")

    newId = Trim(InputBox("Enter the new fund's identifier, exactly as it should appear:", "Add New Fund"))

    If newId = "" Then
    
        Exit Sub
    End If

   
    existingMatch = Application.Match(newId, ThisWorkbook.Names("PositionKey").RefersToRange, 0)
    If Not IsError(existingMatch) Then
        MsgBox "'" & newId & "' already exists in the position sheet (row " & _
               ThisWorkbook.Names("PositionKey").RefersToRange.Cells(1, 1).Row + CLng(existingMatch) - 1 & _
               "). Nothing added.", vbExclamation, "Add New Fund"
        Exit Sub
    End If

    If MsgBox("Create a new fund/position row for '" & newId & "'?" & vbCrLf & _
              "This inserts a blank row at the top of the fund and position sheets.", _
              vbQuestion + vbYesNo, "Add New Fund") = vbNo Then
        Exit Sub
    End If

    wasEnabled = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False


    wsFund.Rows(3).Insert Shift:=xlDown
    wsFund.Range("A3").Value = newId
    wsFund.Range("B3").Value = newId

    wsPosition.Rows(3).Insert Shift:=xlDown
    wsPosition.Range("A3").Value = newId
    wsPosition.Range("B3").Value = newId

    wsMain.Range("C6").Value = newId

    Application.ScreenUpdating = True
    Application.EnableEvents = wasEnabled

    Application.Calculate

    

    MsgBox "New fund '" & newId & "' created on the fund and position sheets.", _
           vbInformation, "Add New Fund"
End Sub




Sub CheckDuplicateFundIDs()

    Dim report As String

    report = FindDuplicatesInRange("position", ThisWorkbook.Names("PositionKey").RefersToRange)
    report = report & FindDuplicatesInRange("fund", ThisWorkbook.Names("FundKey").RefersToRange)

    If report = "" Then
        MsgBox "No duplicate identifiers found in position!A or fund!A.", vbInformation, "Duplicate check"
    Else
        MsgBox "Duplicate identifiers found:" & vbCrLf & vbCrLf & report, vbExclamation, "Duplicate check"
    End If
End Sub

Private Function FindDuplicatesInRange(ByVal sheetLabel As String, ByVal rng As Range) As String
    Dim dict As Object
    Dim c As Range
    Dim v As String
    Dim result As String
    Dim k As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    result = ""

    For Each c In rng.Cells
        v = Trim(CStr(c.Value))
        If v <> "" Then
            If dict.Exists(v) Then
                dict(v) = dict(v) + 1
            Else
                dict.Add v, 1
            End If
        End If
    Next c

    For Each k In dict.Keys
        If dict(k) > 1 Then
            result = result & "  " & sheetLabel & "!A: '" & k & "' appears " & dict(k) & " times" & vbCrLf
        End If
    Next k

    FindDuplicatesInRange = result
End Function

Sub ExportAllFundSnapshots()
    ' Every fund into a new plain .xlsx, one tab each
    Dim wsMain As Worksheet, wsBank As Worksheet
    Dim keyRange As Range, usedRange As Range, c As Range
    Dim newWb As Workbook, newWs As Worksheet
    Dim fundId As String, savedC6 As String, outPath As String
    Dim originalCalcMode As XlCalculation
    Dim wasEnabled As Boolean
    Dim usedNames As Object
    Dim i As Long, total As Long, colIdx As Long
    Dim overriddenCount As Long
    Dim resp As VbMsgBoxResult

    Set wsMain = ThisWorkbook.Sheets("Fund View")
    Set wsBank = ThisWorkbook.Sheets("FormulaBank")
    Set keyRange = ThisWorkbook.Names("PositionKey").RefersToRange

   
    overriddenCount = CountOverriddenTrackedCells(wsBank, wsMain)
    If overriddenCount > 0 Then
        resp = MsgBox(overriddenCount & " tracked cell(s) on Fund View currently show a manually-" & _
                       "overridden value instead of a live formula. Every fund's snapshot would " & _
                       "inherit whatever is frozen there right now, not that fund's real number." & _
                       vbCrLf & vbCrLf & "Restore all formulas first (recommended), or cancel and " & _
                       "check by hand?", vbExclamation + vbYesNoCancel, "Overridden cells found")
        If resp = vbCancel Then Exit Sub
        If resp = vbYes Then RestoreYellowFormulas
    End If

    total = keyRange.Cells.Count
    If MsgBox("This will snapshot all " & total & " funds into a new workbook, one tab each. " & _
              "With this many funds it can take a while - continue?", vbQuestion + vbYesNo, _
              "Export All Fund Snapshots") = vbNo Then
        Exit Sub
    End If

    savedC6 = wsMain.Range("C6").Value
    originalCalcMode = Application.Calculation
    wasEnabled = Application.EnableEvents
    Set usedNames = CreateObject("Scripting.Dictionary")

    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    Set newWb = Workbooks.Add(xlWBATWorksheet)

    i = 0
    For Each c In keyRange.Cells
        fundId = Trim(CStr(c.Value))
        If fundId <> "" Then
            i = i + 1
            Application.StatusBar = "Exporting fund " & i & " of " & total & ": " & fundId

            wsMain.Range("C6").Value = fundId
            Application.Calculate

            Set usedRange = wsMain.usedRange

            Set newWs = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.Count))
            newWs.Name = SafeSheetName(fundId, usedNames)

            usedRange.Copy
            newWs.Range("A1").PasteSpecial Paste:=xlPasteValues
            newWs.Range("A1").PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False

            For colIdx = 1 To usedRange.Columns.Count
                newWs.Columns(colIdx).ColumnWidth = wsMain.Columns(colIdx).ColumnWidth
            Next colIdx
        End If
    Next c

  
    Application.DisplayAlerts = False
    Do While newWb.Sheets.Count > total
        newWb.Sheets(1).Delete
    Loop
    Application.DisplayAlerts = True

    wsMain.Range("C6").Value = savedC6
    Application.Calculate

    Application.ScreenUpdating = True
    Application.Calculation = originalCalcMode
    Application.EnableEvents = wasEnabled
    Application.StatusBar = False

    If ThisWorkbook.Path <> "" Then
        outPath = ThisWorkbook.Path & Application.PathSeparator
    Else
        outPath = Application.DefaultFilePath & Application.PathSeparator
    End If
    outPath = outPath & "Fund Snapshots " & Format(Now, "yyyy-mm-dd HHmm") & ".xlsx"

    newWb.SaveAs Filename:=outPath, FileFormat:=xlOpenXMLWorkbook

    MsgBox "Exported " & total & " fund snapshots to:" & vbCrLf & outPath, _
           vbInformation, "Export All Fund Snapshots"
End Sub

Private Function CountOverriddenTrackedCells(ByVal wsBank As Worksheet, ByVal wsMain As Worksheet) As Long
    Dim i As Long, lastRow As Long, cnt As Long
    lastRow = wsBank.Cells(wsBank.Rows.Count, "A").End(xlUp).Row
    cnt = 0
    For i = 2 To lastRow
        If Not wsMain.Range(wsBank.Cells(i, 1).Value).HasFormula Then
            cnt = cnt + 1
        End If
    Next i
    CountOverriddenTrackedCells = cnt
End Function

Private Function SafeSheetName(ByVal rawName As String, ByRef usedNames As Object) As String
 
    Dim nm As String, badChars As String
    Dim i As Long
    Dim baseName As String, candidate As String, suffixText As String
    Dim suffix As Long

    badChars = ":\/?*[]"
    nm = rawName
    For i = 1 To Len(badChars)
        nm = Replace(nm, Mid(badChars, i, 1), " ")
    Next i
    nm = Trim(nm)
    If nm = "" Then nm = "Fund"
    If Len(nm) > 31 Then nm = Left(nm, 31)

    baseName = nm
    candidate = nm
    suffix = 1
    Do While usedNames.Exists(LCase(candidate))
        suffix = suffix + 1
        suffixText = " (" & suffix & ")"
        candidate = Left(baseName, 31 - Len(suffixText)) & suffixText
    Loop

    usedNames.Add LCase(candidate), True
    SafeSheetName = candidate
End Function
Sub RestoreOverriddenCellsOnly()
    ' Formulas back on the overridden cells only
    Const YELLOW_FILL As Long = 13434879
    Const CHUNK_REFS  As Long = 20
    Dim wsBank As Worksheet, wsMain As Worksheet
    Dim bank As Variant
    Dim i As Long, lastRow As Long, fixed As Long, pushedBack As Long
    Dim refs As String, inChunk As Long, fixedList As String
    Dim savedCalc As XlCalculation
    Dim savedEvents As Boolean, savedScreen As Boolean

    Set wsBank = ThisWorkbook.Sheets("FormulaBank")
    Set wsMain = ThisWorkbook.Sheets("Fund View")
    lastRow = wsBank.Cells(wsBank.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

  
    bank = wsBank.Range("A2:D" & lastRow).Value

    savedCalc = Application.Calculation
    savedEvents = Application.EnableEvents
    savedScreen = Application.ScreenUpdating
    On Error GoTo CleanUp
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For i = 1 To UBound(bank, 1)
        If Len(bank(i, 1)) > 0 Then
            If Not wsMain.Range(CStr(bank(i, 1))).HasFormula Then
                If bank(i, 3) = "Y" Then
                    wsMain.Range(CStr(bank(i, 1))).FormulaArray = CStr(bank(i, 2))
                Else
                    wsMain.Range(CStr(bank(i, 1))).Formula = CStr(bank(i, 2))
                End If
                fixed = fixed + 1
                If bank(i, 4) = "Y" Then pushedBack = pushedBack + 1
                If Len(fixedList) < 200 Then
                    If Len(fixedList) > 0 Then fixedList = fixedList & ", "
                    fixedList = fixedList & CStr(bank(i, 1))
                End If

                If Len(refs) > 0 Then refs = refs & ","
                refs = refs & CStr(bank(i, 1))
                inChunk = inChunk + 1
                If inChunk = CHUNK_REFS Then
                    ApplyYellow wsMain, refs, YELLOW_FILL
                    refs = "": inChunk = 0
                End If
            End If
        End If
    Next i
    If Len(refs) > 0 Then ApplyYellow wsMain, refs, YELLOW_FILL

CleanUp:
    Application.ScreenUpdating = savedScreen
    Application.EnableEvents = savedEvents
    Application.Calculation = savedCalc

    If Err.Number <> 0 Then
        MsgBox "Stopped at FormulaBank row " & i + 1 & ": " & Err.Description, vbExclamation
    ElseIf fixed = 0 Then
        MsgBox "No overridden cells found - every tracked cell already holds a live formula.", _
               vbInformation, "Restore overridden cells"
    Else
        Dim msg As String
        msg = fixed & " overridden cell" & IIf(fixed = 1, "", "s") & " restored (" & _
              Left(fixedList, IIf(Len(fixedList) > 200, 200, Len(fixedList))) & _
              IIf(Len(fixedList) >= 200, "...", "") & ")." & vbCrLf & _
              "All other tracked cells were left untouched."
        ' Empty block - pushedBack counted but never shown.
        If pushedBack > 0 Then
        End If
        MsgBox msg, vbInformation, "Restore overridden cells"
    End If
End Sub
