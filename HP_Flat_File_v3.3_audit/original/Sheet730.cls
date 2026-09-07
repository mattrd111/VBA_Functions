Option Explicit

Private Const OVERRIDE_FILL As Long = 13551615   ' RGB(255,199,206) light red
Private Const OVERRIDE_FONT As Long = 6693        ' RGB(156,0,6) dark red
Private Const YELLOW_FILL As Long = 13434879      ' RGB(255,255,204) pale yellow


Private Const COMPANY_QUARTER_RANGE As String = "AT117:AZ159"
Private lastSelAddress As String
Private lastSelValue As Variant

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

End Sub

Private Sub Worksheet_PivotTableBeforeDiscardChanges(ByVal TargetPivotTable As PivotTable, ByVal ValueChangeStart As Long, ByVal ValueChangeEnd As Long)

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count = 1 Then
        lastSelAddress = Target.Address(False, False)
        lastSelValue = Target.Value
    Else
        lastSelAddress = ""
        lastSelValue = Empty
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range

    If Not Intersect(Target, Me.Range("C6")) Is Nothing Then
        WarnIfDuplicateKey
        WarnIfOverridesStale

    End If


   

    If Target.Cells.Count > 1 Then
        For Each cell In Target.Cells
            HandlePossibleOverride cell
        Next cell
    Else
        HandlePossibleOverride Target
    End If
End Sub

Private Sub WarnIfOverridesStale()
    ' Typed-over values get reset on c6 change
    Dim wsBank As Worksheet
    Dim i As Long, lastRow As Long
    Dim cellRef As String
    Dim overriddenCount As Long
    Dim overriddenList As String
    Dim msg As String

    Set wsBank = ThisWorkbook.Sheets("FormulaBank")
    lastRow = wsBank.Cells(wsBank.Rows.Count, "A").End(xlUp).Row

    overriddenCount = 0
    overriddenList = ""
    For i = 2 To lastRow
        cellRef = wsBank.Cells(i, 1).Value
        If Not Me.Range(cellRef).HasFormula Then
            overriddenCount = overriddenCount + 1
            If Len(overriddenList) < 200 Then
                If Len(overriddenList) > 0 Then overriddenList = overriddenList & ", "
                overriddenList = overriddenList & cellRef
            End If
        End If
    Next i

    If overriddenCount = 0 Then Exit Sub

    msg = overriddenCount & " cell" & IIf(overriddenCount = 1, "", "s") & _
          " on this sheet " & IIf(overriddenCount = 1, "is", "are") & _
          " still showing a manually-overridden value from before you changed C6 these will be restored to formula"
 

  
        RestoreYellowFormulas
    
End Sub

Private Sub WarnIfDuplicateKey()
    ' Duplicates in position!A
    Dim currentId As String
    Dim cnt As Long

    currentId = Trim(Me.Range("C6").Value)
    If currentId = "" Then Exit Sub

    cnt = Application.WorksheetFunction.CountIf(ThisWorkbook.Names("PositionKey").RefersToRange, currentId)
    If cnt > 1 Then
        MsgBox "'" & currentId & "' appears " & cnt & " times in the position sheet's identifier column." & _
               vbCrLf & vbCrLf & _
               "Every INDEX/MATCH lookup on this sheet will use whichever of those rows " & _
               "comes first - Fix the duplicate on the position " & _
               "", _
               vbCritical, "Duplicate fund identifier"
    End If
End Sub

Private Sub HandlePossibleOverride(ByVal chg As Range)
    Dim wsBank As Worksheet
    Dim matchRow As Variant
    Dim isCompanyQuarterCell As Boolean

    On Error GoTo CleanExit

    ' Tracked cells only
    Set wsBank = ThisWorkbook.Sheets("FormulaBank")
    isCompanyQuarterCell = Not Intersect(chg, Me.Range(COMPANY_QUARTER_RANGE)) Is Nothing

    matchRow = Application.Match(chg.Address(False, False), wsBank.Columns("A"), 0)
    If IsError(matchRow) And Not isCompanyQuarterCell Then GoTo CleanExit

    If chg.HasFormula Then
      
        chg.Interior.Color = YELLOW_FILL
        chg.Font.ColorIndex = xlAutomatic
        If Not chg.Comment Is Nothing Then chg.Comment.Delete
        GoTo CleanExit
    End If

    ' Tracked cell holding a constant = an override. Flag it red.
    chg.Interior.Color = OVERRIDE_FILL
    chg.Font.Color = OVERRIDE_FONT

    RecordOverrideComment chg, GetPreEditDescription(chg)


    StampFundUpdated

    If isCompanyQuarterCell Then
        PushCompanyQuarterOverride chg
    Else
 
        If wsBank.Cells(CLng(matchRow), 4).Value <> "Y" Then GoTo CleanExit
        PushOverrideToSource chg, wsBank, CLng(matchRow)
    End If

CleanExit:
End Sub

Private Function GetPreEditDescription(ByVal chg As Range) As String
  
    If chg.Cells.Count = 1 And chg.Address(False, False) = lastSelAddress Then
        GetPreEditDescription = DescribeOldValue(lastSelValue)
    Else
        GetPreEditDescription = "unknown (multi-cell edit/paste)"
    End If
End Function

Private Function DescribeOldValue(ByVal v As Variant) As String
    
    On Error GoTo Fallback

    If IsEmpty(v) Then
        DescribeOldValue = "(blank)"
    ElseIf IsError(v) Then
        DescribeOldValue = "(error value)"
    ElseIf VarType(v) = vbString And Len(v) = 0 Then
        DescribeOldValue = "(empty text)"
    ElseIf VarType(v) = vbString And Left(v, 1) = "=" Then
        
        DescribeOldValue = v
    ElseIf VarType(v) = vbDate Then
        DescribeOldValue = Format(v, "dd/mm/yyyy")
    Else
        DescribeOldValue = CStr(v)
    End If
    Exit Function

Fallback:
    DescribeOldValue = "(unreadable)"
End Function

Private Sub RecordOverrideComment(ByVal chg As Range, ByVal oldValueDescription As String)
  
    Dim noteText As String

    noteText = "Overridden by " & Environ("USERNAME") & " on " & Format(Now, "dd/mm/yyyy hh:nn") & _
               vbCrLf & "Previous value: " & oldValueDescription

    On Error Resume Next
    If chg.Comment Is Nothing Then
        chg.AddComment noteText
    Else
        chg.Comment.Text noteText
    End If
    On Error GoTo 0
End Sub

Private Sub StampFundUpdated()
    ' Username + date into position!AB / position!S for the fund in C6.

    Dim wsPosition As Worksheet
    Dim currentId As String
    Dim matchPos As Variant
    Dim targetRow As Long
    Dim wasEnabled As Boolean

    On Error GoTo CleanExit

    currentId = Trim(Me.Range("C6").Value)
    If currentId = "" Then Exit Sub

    Set wsPosition = ThisWorkbook.Sheets("position")

    matchPos = Application.Match(currentId, ThisWorkbook.Names("PositionKey").RefersToRange, 0)
    If IsError(matchPos) Then Exit Sub

    targetRow = ThisWorkbook.Names("PositionKey").RefersToRange.Cells(1, 1).Row + CLng(matchPos) - 1

    wasEnabled = Application.EnableEvents
    Application.EnableEvents = False
    wsPosition.Range("AB" & targetRow).Value = Environ("USERNAME")
    wsPosition.Range("S" & targetRow).Value = Date
    Application.EnableEvents = wasEnabled

CleanExit:
End Sub

Private Sub PushOverrideToSource(ByVal chg As Range, ByVal wsBank As Worksheet, ByVal bankRow As Long)
    Dim sourceSheetName As String, keyExpr As String, keyRange As String, targetCol As String
    Dim wsSource As Worksheet
    Dim keyExprSubbed As String
    Dim keyValue As Variant
    Dim matchPos As Variant
    Dim targetRow As Long
    Dim wasEnabled As Boolean

    ' Columns: E=SourceSheet, F=KeyExpr, G=KeyRange, H=TargetCol
    sourceSheetName = wsBank.Cells(bankRow, 5).Value
    keyExpr = wsBank.Cells(bankRow, 6).Value
    keyRange = wsBank.Cells(bankRow, 7).Value
    targetCol = wsBank.Cells(bankRow, 8).Value

    On Error GoTo Fail
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)


    keyExprSubbed = Replace(keyExpr, "ROW()", CStr(chg.Row))
    keyValue = Me.Evaluate(keyExprSubbed)

    matchPos = Application.Match(keyValue, wsSource.Range(keyRange), 0)
    If IsError(matchPos) Then
        If sourceSheetName = "position" Or sourceSheetName = "fund" Then
     
       
            MsgBox "Could not find a matching row in '" & sourceSheetName & "' for " & _
                   chg.Address(False, False) & " (key: " & keyValue & ")." & vbCrLf & _
                   "Your override was NOT written back " & _
                   "to the source data. If '" & keyValue & "' is a new fund, run " & _
                   "AddNewFund first.", vbExclamation, "Pushback failed"
            Exit Sub
        End If


        targetRow = AppendOverrideRow(wsSource, keyRange, keyValue, targetCol, chg, sourceSheetName)
        Exit Sub
    End If

    targetRow = wsSource.Range(keyRange).Cells(1, 1).Row + CLng(matchPos) - 1

    wasEnabled = Application.EnableEvents
    Application.EnableEvents = False
    wsSource.Range(targetCol & targetRow).Value = chg.Value
    Application.EnableEvents = wasEnabled


    Exit Sub

Fail:
    MsgBox "Error pushing your override for " & chg.Address(False, False) & " back to '" & _
           sourceSheetName & "': " & Err.Description, vbCritical, "Pushback error"
End Sub

Private Function AppendOverrideRow(ByVal wsSource As Worksheet, ByVal keyRangeName As String, _
                                    ByVal keyValue As Variant, ByVal targetCol As String, _
                                    ByVal chg As Range, ByVal sourceSheetName As String) As Long

    '

    Dim keyRng As Range
    Dim insertRow As Long
    Dim keyCol As Long
    Dim colSourceTab As Variant, colPositionId As Variant, colSourceRow As Variant
    Dim wasEnabled As Boolean, wasScreen As Boolean

    Set keyRng = wsSource.Range(keyRangeName)
    keyCol = keyRng.Column
    insertRow = keyRng.Row + 1

    wasEnabled = Application.EnableEvents
    wasScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    wsSource.Rows(insertRow).Insert Shift:=xlDown

    wsSource.Cells(insertRow, keyCol).Value = keyValue

    colSourceTab = Application.Match("source_tab", wsSource.Rows(1), 0)
    If Not IsError(colSourceTab) Then wsSource.Cells(insertRow, CLng(colSourceTab)).Value = Me.Range("C6").Value

    colPositionId = Application.Match("position_id", wsSource.Rows(1), 0)
    If Not IsError(colPositionId) Then wsSource.Cells(insertRow, CLng(colPositionId)).Value = Me.Range("C6").Value

    colSourceRow = Application.Match("source_row", wsSource.Rows(1), 0)
    If Not IsError(colSourceRow) Then wsSource.Cells(insertRow, CLng(colSourceRow)).Value = chg.Row

    Dim colCompanyId As Variant
    colCompanyId = Application.Match("company_id", wsSource.Rows(1), 0)
    If Not IsError(colCompanyId) Then
        If Len(Trim(CStr(Me.Cells(chg.Row, "G").Value))) > 0 Then
            wsSource.Cells(insertRow, CLng(colCompanyId)).Value = Me.Cells(chg.Row, "G").Value
        End If
    End If

    wsSource.Range(targetCol & insertRow).Value = chg.Value

    Application.ScreenUpdating = wasScreen
    Application.EnableEvents = wasEnabled

    AppendOverrideRow = insertRow

    MsgBox "No existing row for this fund/period in '" & sourceSheetName & "' - created row " & _
           insertRow & " and wrote your value into column " & targetCol & "." & vbCrLf & vbCrLf & _
           "Check the other columns on that new row (dates, company, etc.) - only the identifier " & _
           "and your overridden value were filled in automatically.", vbInformation, "New row created"
End Function

Private Sub PushCompanyQuarterOverride(ByVal chg As Range)

    Dim wsSource As Worksheet
    Dim dateText As String
    Dim keyValue As String
    Dim searchRange As Range
    Dim matchPos As Variant
    Dim otherMatch As Variant
    Dim targetRow As Long
    Dim insertRow As Long
    Dim fundId As String
    Dim companyId As String
    Dim valueCol As Long
    Dim companyNameVal As Variant, currencyVal As Variant
    Dim rawHeaderVal As Variant
    Dim wasEnabled As Boolean, wasScreen As Boolean

    On Error GoTo Fail
    Set wsSource = ThisWorkbook.Sheets("company_nav")

    rawHeaderVal = Me.Cells(114, chg.Column).Value
    If Not IsDate(rawHeaderVal) Then

        If Not PromptForQuarterHeader(chg.Column) Then
            MsgBox "No quarter-end date for column " & chg.Address(False, False) & _
                   ", so it couldn't be matched back to 'company_nav'." & vbCrLf & vbCrLf & _
                   "Your override was NOT written back.", _
                   vbExclamation, "Pushback failed"
            Exit Sub
        End If
        rawHeaderVal = Me.Cells(114, chg.Column).Value
    End If

    dateText = FormatQuarterDate(rawHeaderVal)
    If dateText = "" Then
        MsgBox "This column's header in row 114 isn't a recognisable date, so " & chg.Address(False, False) & _
               " couldn't be matched back to 'company_nav'." & vbCrLf & vbCrLf & _
               "Your override was NOT written back.", _
               vbExclamation, "Pushback failed"
        Exit Sub
    End If

    fundId = Trim(Me.Range("C6").Value)
    companyId = CStr(Me.Cells(chg.Row, "G").Value)
    keyValue = fundId & "|" & chg.Row & "|" & dateText

    ' Column D = lookup_key.
    matchPos = Application.Match("lookup_key", wsSource.Rows(1), 0)
    If IsError(matchPos) Then
        MsgBox "'company_nav' has no 'lookup_key' column header in row 1 - can't match or create rows.", _
               vbCritical, "Pushback failed"
        Exit Sub
    End If
    Set searchRange = wsSource.Range(wsSource.Cells(2, CLng(matchPos)), wsSource.Cells(15131, CLng(matchPos)))

    valueCol = 0
    otherMatch = Application.Match("nav_value", wsSource.Rows(1), 0)
    If Not IsError(otherMatch) Then valueCol = CLng(otherMatch)
    If valueCol = 0 Then
        MsgBox "'company_nav' has no 'nav_value' column header in row 1 - can't write your override back.", _
               vbCritical, "Pushback failed"
        Exit Sub
    End If

    matchPos = Application.Match(keyValue, searchRange, 0)

    wasEnabled = Application.EnableEvents
    wasScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If IsError(matchPos) Then

        insertRow = 3
        wsSource.Rows(insertRow).Insert Shift:=xlDown


        companyNameVal = Empty
        currencyVal = Empty
        Dim companyIdCol As Variant, companyIdRange As Range, cMatch As Variant
        Dim nameColMatch As Variant, currColMatch As Variant
        companyIdCol = Application.Match("company_id", wsSource.Rows(1), 0)
        nameColMatch = Application.Match("company_name", wsSource.Rows(1), 0)
        currColMatch = Application.Match("currency", wsSource.Rows(1), 0)
        If Not IsError(companyIdCol) Then
            Set companyIdRange = wsSource.Range(wsSource.Cells(2, CLng(companyIdCol)), wsSource.Cells(15131, CLng(companyIdCol)))
            cMatch = Application.Match(companyId, companyIdRange, 0)
            If Not IsError(cMatch) Then
                Dim existingRow As Long
                existingRow = companyIdRange.Cells(1, 1).Row + CLng(cMatch) - 1
                If Not IsError(nameColMatch) Then companyNameVal = wsSource.Cells(existingRow, CLng(nameColMatch)).Value
                If Not IsError(currColMatch) Then currencyVal = wsSource.Cells(existingRow, CLng(currColMatch)).Value
            End If
        End If

        wsSource.Cells(insertRow, "A").Value = fundId       ' source_tab
        wsSource.Cells(insertRow, "B").Value = fundId       ' position_id
        wsSource.Cells(insertRow, "C").Value = chg.Row      ' source_row
        wsSource.Cells(insertRow, "D").Value = keyValue     ' lookup_key
        wsSource.Cells(insertRow, "E").Value = companyId    ' company_id
        If IsDate(rawHeaderVal) Then
            wsSource.Cells(insertRow, "H").Value = CDate(rawHeaderVal)  ' Nav_date, real date
        Else
            wsSource.Cells(insertRow, "H").Value = dateText
        End If
        wsSource.Cells(insertRow, valueCol).Value = chg.Value  ' nav_value
        If Not IsEmpty(companyNameVal) Then wsSource.Cells(insertRow, "F").Value = companyNameVal
        If Not IsEmpty(currencyVal) Then wsSource.Cells(insertRow, "G").Value = currencyVal

        Application.ScreenUpdating = wasScreen
        Application.EnableEvents = wasEnabled

        Dim missingNote As String
        missingNote = ""
        If IsEmpty(companyNameVal) Or IsEmpty(currencyVal) Then
            missingNote = vbCrLf & vbCrLf & "Couldn't find an existing 'company_nav' row for company " & companyId & _
                           " to copy company_name/currency from - check row " & insertRow & " and fill those in by hand."
        End If

        MsgBox "No existing row for " & fundId & " / row " & chg.Row & " / " & dateText & _
               " in 'company_nav' - created a new row and wrote the value in." & missingNote & vbCrLf & vbCrLf & _
               "It will populate back into this grid next time the sheet recalculates.", _
               vbInformation, "New row created"
    Else
        targetRow = searchRange.Cells(1, 1).Row + CLng(matchPos) - 1
        wsSource.Cells(targetRow, valueCol).Value = chg.Value

        Application.ScreenUpdating = wasScreen
        Application.EnableEvents = wasEnabled
    End If

    Exit Sub

Fail:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error pushing your override for " & chg.Address(False, False) & " back to 'company_nav': " & _
           Err.Description, vbCritical, "Pushback error"
End Sub

Private Function PromptForQuarterHeader(ByVal colNum As Long) As Boolean

    Dim colLetter As String
    Dim entry As String
    Dim wasEnabled As Boolean

    colLetter = Split(Me.Cells(1, colNum).Address(True, False), "$")(0)

    Do
        entry = Trim(InputBox( _
            "Column " & colLetter & " has no usable quarter-end date in row 114, so this " & _
            "override can't be matched back to 'company_nav'." & vbCrLf & vbCrLf & _
            "Enter the quarter-end date for column " & colLetter & " (e.g. 31/03/2026):", _
            "Quarter-end date needed"))

        If entry = "" Then Exit Function

        If IsDate(entry) Then
            wasEnabled = Application.EnableEvents
            Application.EnableEvents = False
            Me.Cells(114, colNum).Value = CDate(entry)
            Application.EnableEvents = wasEnabled
            PromptForQuarterHeader = True
            Exit Function
        End If

        MsgBox "'" & entry & "' isn't a date. Try again, or Cancel to leave the override " & _
               "on the sheet without writing it back.", vbExclamation, "Quarter-end date needed"
    Loop
End Function

Private Function FormatQuarterDate(ByVal v As Variant) As String

    On Error GoTo Fallback
    If IsEmpty(v) Or v = "" Then
        FormatQuarterDate = ""
    ElseIf IsDate(v) Then
        FormatQuarterDate = Format(CDate(v), "yyyy-mm-dd")
    Else
        FormatQuarterDate = CStr(v)
    End If
    Exit Function

Fallback:
    FormatQuarterDate = ""
End Function
