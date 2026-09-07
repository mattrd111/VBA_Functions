Option Explicit

Private Const OVERRIDE_FILL As Long = 13551615   ' RGB(255,199,206) light red
Private Const OVERRIDE_FONT As Long = 6693        ' RGB(156,0,6) dark red
Private Const YELLOW_FILL As Long = 13434879      ' RGB(255,255,204) the original pale yellow

' Remembers the single cell you were last sitting on, and its value at
' that point, so that if you type over it Worksheet_Change can still say
' what it used to hold. Only works for single-cell hand edits - a
' paste over several cells at once can only recover the old value of
' whichever cell was the anchor of the selection beforehand; the rest are
' recorded as "unknown" rather than guessed at. See RecordOverrideComment.
Private lastSelAddress As String
Private lastSelValue As Variant

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
    ' Fires whenever C6 changes. Scans every tracked cell and, if any of
    ' them are still static overrides (rather than live formulas), warns
    ' that they belong to whatever fund was in C6 before this edit and
    ' offers to restore all formulas so the sheet is consistent again.
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
          " still showing a manually-overridden value from before you changed C6 (e.g. " & _
          overriddenList & IIf(Len(overriddenList) >= 200, "...", "") & ")." & vbCrLf & vbCrLf & _
          "Those values belong to the PREVIOUS fund and will NOT update for the fund now in C6." & _
          vbCrLf & vbCrLf & "Restore all formulas now?"

    If MsgBox(msg, vbExclamation + vbYesNo, "Stale overrides after fund change") = vbYes Then
        RestoreYellowFormulas
    End If
End Sub

Private Sub WarnIfDuplicateKey()
    ' Fires whenever C6 changes. If the identifier just entered exists
    ' more than once in position!A, every lookup formula on this sheet
    ' will silently use whichever matching row comes first.
    Dim currentId As String
    Dim cnt As Long

    currentId = Trim(Me.Range("C6").Value)
    If currentId = "" Then Exit Sub

    cnt = Application.WorksheetFunction.CountIf(ThisWorkbook.Names("PositionKey").RefersToRange, currentId)
    If cnt > 1 Then
        MsgBox "'" & currentId & "' appears " & cnt & " times in the position sheet's identifier column." & _
               vbCrLf & vbCrLf & _
               "Every INDEX/MATCH lookup on this sheet will silently use whichever of those rows " & _
               "comes first - not necessarily the one you mean. Fix the duplicate on the position " & _
               "sheet before trusting any number shown here.", _
               vbCritical, "Duplicate fund identifier"
    End If
End Sub

Private Sub HandlePossibleOverride(ByVal chg As Range)
    Dim wsBank As Worksheet
    Dim matchRow As Variant

    On Error GoTo CleanExit

    ' Ignore edits outside the tracked yellow-cell range entirely
    Set wsBank = ThisWorkbook.Sheets("FormulaBank")

    matchRow = Application.Match(chg.Address(False, False), wsBank.Columns("A"), 0)
    If IsError(matchRow) Then GoTo CleanExit

    If chg.HasFormula Then
        ' The formula has been typed back in by hand - clear the override flag
        chg.Interior.Color = YELLOW_FILL
        chg.Font.ColorIndex = xlAutomatic
        ' ...and drop the "old value" note left behind from the override,
        ' if any - it no longer describes anything relevant now the formula
        ' is back.
        If Not chg.Comment Is Nothing Then chg.Comment.Delete
        GoTo CleanExit
    End If

    ' Cell is now a static value while this sheet is tracking it as a
    ' formula cell - someone has genuinely overridden it. Flag it red.
    chg.Interior.Color = OVERRIDE_FILL
    chg.Font.Color = OVERRIDE_FONT

    ' Note what the cell held immediately before this edit, so a reviewer
    ' can see what was overwritten without digging through FormulaBank.
    RecordOverrideComment chg, GetPreEditDescription(chg)

    ' Log who corrected this fund's dashboard and when, on every manual
    ' override - regardless of which source table the specific cell
    ' happens to be tracked back to. See StampFundUpdated for why.
    StampFundUpdated

    ' Column D = CanPushback
    If wsBank.Cells(CLng(matchRow), 4).Value <> "Y" Then GoTo CleanExit

    PushOverrideToSource chg, wsBank, CLng(matchRow)

CleanExit:
End Sub

Private Function GetPreEditDescription(ByVal chg As Range) As String
    ' Returns a human-readable description of what chg held immediately
    ' before this edit. Only Worksheet_SelectionChange (single-cell only)
    ' gives us that - a multi-cell paste or fill can't be attributed to a
    ' single "old value" per cell without guessing, so we say so rather
    ' than fabricate a number.
    If chg.Cells.Count = 1 And chg.Address(False, False) = lastSelAddress Then
        GetPreEditDescription = DescribeOldValue(lastSelValue)
    Else
        GetPreEditDescription = "unknown (multi-cell edit/paste)"
    End If
End Function

Private Function DescribeOldValue(ByVal v As Variant) As String
    ' Renders the previous cell value/formula for the comment text.
    On Error GoTo Fallback

    If IsEmpty(v) Then
        DescribeOldValue = "(blank)"
    ElseIf IsError(v) Then
        DescribeOldValue = "(error value)"
    ElseIf VarType(v) = vbString And Len(v) = 0 Then
        DescribeOldValue = "(empty text)"
    ElseIf VarType(v) = vbString And Left(v, 1) = "=" Then
        ' A formula was captured (SelectionChange stores .Value, which for a
        ' formula cell is the *displayed result*, not the formula text - but
        ' guard here anyway in case that ever changes).
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
    ' Leaves a note on the cell recording what it held before this override,
    ' plus who made the change and when, so a reviewer doesn't have to dig
    ' through FormulaBank or ask around to see what was overwritten.
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
    ' Logs who made a manual correction and when, in the position sheet's
    ' updated_by (AB) / date_of_update (S) columns for whatever fund is
    ' currently selected in C6.
    '
    ' Runs on every manual override, not just ones whose own source
    ' happens to be the position sheet - "Date of Update" / "Updated By"
    ' are meant to answer "was this fund's dashboard hand-corrected, and
    ' by whom", not "which specific source table changed". None of the
    ' other 7 source sheets have an equally unambiguous "corrected by /
    ' on" pair (only as_of_date or entered_by, which mean something
    ' else), so position stays the one place this gets stamped.
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

    ' The key expressions use ROW() meaning "the row of the yellow cell
    ' on this dashboard sheet" - substitute the literal row number before
    ' evaluating, since Evaluate has no cell context of its own.
    keyExprSubbed = Replace(keyExpr, "ROW()", CStr(chg.Row))
    keyValue = Me.Evaluate(keyExprSubbed)

    matchPos = Application.Match(keyValue, wsSource.Range(keyRange), 0)
    If IsError(matchPos) Then
        MsgBox "Could not find a matching row in '" & sourceSheetName & "' for " & _
               chg.Address(False, False) & " (key: " & keyValue & ")." & vbCrLf & _
               "Your override is kept on the dashboard, but it was NOT written back " & _
               "to the source data.", vbExclamation, "Pushback failed"
        Exit Sub
    End If

    targetRow = wsSource.Range(keyRange).Cells(1, 1).Row + CLng(matchPos) - 1

    wasEnabled = Application.EnableEvents
    Application.EnableEvents = False
    wsSource.Range(targetCol & targetRow).Value = chg.Value
    Application.EnableEvents = wasEnabled

    ' Updated By / Date of Update on the position sheet are stamped by
    ' StampFundUpdated (called from HandlePossibleOverride) for every
    ' override, not just this one - see there for why.
    Exit Sub

Fail:
    MsgBox "Error pushing your override for " & chg.Address(False, False) & " back to '" & _
           sourceSheetName & "': " & Err.Description, vbCritical, "Pushback error"
End Sub
