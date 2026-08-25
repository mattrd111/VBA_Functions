Attribute VB_Name = "modWrangleMatch"
'==============================================================================
' modWrangleMatch - reconciling two lists that nearly agree
'------------------------------------------------------------------------------
' Customer names from the CRM against the ones on the invoices. Entity names in
' the model against the ones in the trial balance. GL accounts across two
' charts. The lists are the same list, and none of the strings match.
'
' Three passes, in order of confidence:
'   1  exact, once both sides are tidied up
'   2  exact, once company decoration is stripped - THE ACME GROUP LTD. and
'      Acme Group Limited are the same company
'   3  closest match by edit distance, with a score you can sort on
'
' Nothing is decided for you. The output is a table with a score and a verdict
' per row, and a column left empty for the human answer.
'
' Depends on: modDoctorCommon, modWrangleStack, modWrangleShape (prompts),
'             modString, modDictionary, modWorkbook, modRange, modApp (from \src)
'==============================================================================
Option Explicit

Private Const PAIR_WARNING As Double = 400000#

'------------------------------------------------------------------------------
' FuzzyMatchLists   (menu action)
'------------------------------------------------------------------------------
Public Sub FuzzyMatchLists()
    Dim fromRange As Range, toRange As Range
    Dim fromList As Variant, toList As Variant
    Dim threshold As Long
    Dim stripCompany As Boolean
    Dim rows As Collection
    Dim book As Workbook
    Dim ws As Worksheet
    Dim out As Variant

    Set fromRange = AskForBlock("Select the list to match FROM - one column.")
    If fromRange Is Nothing Then Exit Sub

    Set toRange = AskForBlock("Now the list to match TO - one column.")
    If toRange Is Nothing Then Exit Sub

    threshold = AskForNumber("Below what score should a match be called weak? (0 to 100)", 85)
    If threshold < 0 Or threshold > 100 Then Exit Sub

    stripCompany = (MsgBox("Are these company names?" & vbNewLine & vbNewLine & _
                           "If so, Ltd, Limited, PLC, Inc, Holdings, Group and the like are " & _
                           "ignored when comparing, along with punctuation and a leading 'The'.", _
                           vbYesNo + vbQuestion, DOCTOR_NAME) = vbYes)

    fromList = ColumnValues(fromRange)
    toList = ColumnValues(toRange)

    If Not IsArray(fromList) Or Not IsArray(toList) Then
        MsgBox "One of those lists is empty.", vbExclamation, DOCTOR_NAME
        Exit Sub
    End If

    If CDbl(UBound(fromList)) * CDbl(UBound(toList)) > PAIR_WARNING Then
        If MsgBox(Format$(UBound(fromList), "#,##0") & " x " & Format$(UBound(toList), "#,##0") & _
                  " is a lot of comparisons and may take a few minutes." & vbNewLine & vbNewLine & _
                  "Carry on?", vbYesNo + vbExclamation, DOCTOR_NAME) <> vbYes Then Exit Sub
    End If

    StartTimer
    FastMode True
    out = MatchLists(fromList, toList, threshold, stripCompany)
    FastMode False
    ClearStatusBar

    Set book = Application.Workbooks.Add
    Set ws = DumpToSheet(out, "Matches", book, True)
    If Not ws Is Nothing Then
        FreezeHeader ws, 1
        AutoFitColumns ws, 45, 9
    End If

    Set rows = BuildMatchReport(out, fromRange, toRange, threshold, stripCompany)
    ShowReport rows, "Match summary"

    MsgBox Format$(UBound(fromList), "#,##0") & " name(s) matched against " & _
           Format$(UBound(toList), "#,##0") & " in " & ElapsedText & "." & vbNewLine & vbNewLine & _
           "The 'Your call' column is empty on purpose - the scores are a starting point, " & _
           "not an answer.", vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' MatchLists
' The output table, ready to write: one row per item in the FROM list.
'------------------------------------------------------------------------------
Public Function MatchLists(ByVal fromList As Variant, ByVal toList As Variant, _
                           ByVal threshold As Long, ByVal stripCompany As Boolean) As Variant
    Dim out As Variant
    Dim exactIndex As Object, strippedIndex As Object
    Dim toTidy() As String, toStripped() As String
    Dim i As Long, j As Long
    Dim tidy As String, stripped As String
    Dim best As String, runnerUp As String
    Dim bestScore As Double, runnerScore As Double
    Dim score As Double
    Dim verdict As String

    ReDim toTidy(1 To UBound(toList))
    ReDim toStripped(1 To UBound(toList))
    Set exactIndex = NewDictionary(True)
    Set strippedIndex = NewDictionary(True)

    For j = 1 To UBound(toList)
        toTidy(j) = CleanText(toList(j))
        toStripped(j) = StripCompanyWords(toTidy(j))
        If Not exactIndex.Exists(toTidy(j)) Then exactIndex.Add toTidy(j), j
        If Len(toStripped(j)) > 0 Then
            If Not strippedIndex.Exists(toStripped(j)) Then strippedIndex.Add toStripped(j), j
        End If
    Next j

    ReDim out(1 To UBound(fromList) + 1, 1 To 7)
    out(1, 1) = "From"
    out(1, 2) = "Best match"
    out(1, 3) = "Score"
    out(1, 4) = "Verdict"
    out(1, 5) = "Next best"
    out(1, 6) = "Its score"
    out(1, 7) = "Your call"

    For i = 1 To UBound(fromList)
        If i Mod 25 = 0 Then ShowProgress i, UBound(fromList), "Matching"

        tidy = CleanText(fromList(i))
        stripped = StripCompanyWords(tidy)
        best = ""
        runnerUp = ""
        bestScore = 0
        runnerScore = 0
        verdict = "No match"

        If Len(tidy) = 0 Then
            verdict = "Blank"

        ElseIf exactIndex.Exists(tidy) Then
            best = toTidy(exactIndex(tidy))
            bestScore = 100
            verdict = "Exact"

        ElseIf stripCompany And Len(stripped) > 0 And strippedIndex.Exists(stripped) Then
            best = toTidy(strippedIndex(stripped))
            bestScore = 100
            verdict = "Same company"

        Else
            For j = 1 To UBound(toList)
                If stripCompany Then
                    score = ScoreOf(stripped, toStripped(j))
                Else
                    score = ScoreOf(tidy, toTidy(j))
                End If

                If score > bestScore Then
                    runnerScore = bestScore
                    runnerUp = best
                    bestScore = score
                    best = toTidy(j)
                ElseIf score > runnerScore Then
                    runnerScore = score
                    runnerUp = toTidy(j)
                End If
            Next j

            If bestScore >= threshold Then
                verdict = "Likely"
            ElseIf bestScore >= threshold - 15 Then
                verdict = "Weak - check"
            Else
                verdict = "No match"
                best = IIf(bestScore > 0, best, "")
            End If
        End If

        out(i + 1, 1) = fromList(i)
        out(i + 1, 2) = best
        out(i + 1, 3) = IIf(bestScore > 0, Round(bestScore, 1), "")
        out(i + 1, 4) = verdict
        out(i + 1, 5) = runnerUp
        out(i + 1, 6) = IIf(runnerScore > 0, Round(runnerScore, 1), "")
    Next i

    MatchLists = out
End Function

'------------------------------------------------------------------------------
' StripCompanyWords
' Reduces a company name to the bit that identifies it.
'
'   "THE Acme Group Ltd."  ->  "ACME"
'   "Acme Holdings PLC"    ->  "ACME"
'------------------------------------------------------------------------------
Public Function StripCompanyWords(ByVal text As String) As String
    Dim words As Variant
    Dim result As String
    Dim i As Long
    Dim word As String

    text = UCase$(CleanText(text))
    text = RemovePunctuation(text)
    If Len(text) = 0 Then Exit Function

    words = Split(text, " ")
    For i = LBound(words) To UBound(words)
        word = CStr(words(i))
        If Len(word) > 0 And Not IsDecorationWord(word) Then
            result = result & IIf(Len(result) = 0, "", " ") & word
        End If
    Next i

    ' If the name was nothing but decoration, keep what we started with rather
    ' than matching every empty string to every other one.
    If Len(result) = 0 Then result = text

    StripCompanyWords = result
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Function ScoreOf(ByVal a As String, ByVal b As String) As Double
    Dim longest As Long
    Dim allowed As Long
    Dim distance As Long

    If Len(a) = 0 Or Len(b) = 0 Then Exit Function

    longest = IIf(Len(a) > Len(b), Len(a), Len(b))

    ' Edit distance is at least the difference in length, so a pair that cannot
    ' possibly score well is not worth measuring.
    allowed = longest \ 2
    If Abs(Len(a) - Len(b)) > allowed Then Exit Function

    distance = Levenshtein(a, b, True)
    ScoreOf = (1 - distance / longest) * 100
End Function

Private Function ColumnValues(ByVal rng As Range) As Variant
    Dim data As Variant
    Dim out() As Variant
    Dim r As Long, n As Long

    If rng Is Nothing Then Exit Function
    data = RangeToArray(rng)
    If Not IsArray(data) Then Exit Function

    ReDim out(1 To UBound(data, 1))
    For r = 1 To UBound(data, 1)
        If Not IsBlank(data(r, 1)) Then
            n = n + 1
            out(n) = data(r, 1)
        End If
    Next r

    If n = 0 Then Exit Function
    ReDim Preserve out(1 To n)
    ColumnValues = out
End Function

Private Function RemovePunctuation(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        Select Case ch
            Case "A" To "Z", "a" To "z", "0" To "9", " "
                out = out & ch
            Case Else
                out = out & " "
        End Select
    Next i

    Do While InStr(out, "  ") > 0
        out = Replace$(out, "  ", " ")
    Loop
    RemovePunctuation = Trim$(out)
End Function

Private Function IsDecorationWord(ByVal word As String) As Boolean
    Select Case UCase$(word)
        Case "THE", "LTD", "LIMITED", "PLC", "LLP", "LLC", "LP", "INC", "INCORPORATED", _
             "CO", "COMPANY", "CORP", "CORPORATION", "HOLDING", "HOLDINGS", "GROUP", _
             "INTERNATIONAL", "INTL", "GMBH", "AG", "SA", "SARL", "SAS", "BV", "NV", _
             "AB", "AS", "OY", "PTY", "SPA", "SRL", "KFT", "ZOO", "AND", "&"
            IsDecorationWord = True
    End Select
End Function

Private Function BuildMatchReport(ByVal out As Variant, ByVal fromRange As Range, _
                                  ByVal toRange As Range, ByVal threshold As Long, _
                                  ByVal stripCompany As Boolean) As Collection
    Dim rows As Collection
    Dim counts As Object
    Dim r As Long
    Dim key As Variant

    Set rows = NewReport()
    Set counts = NewDictionary(True)

    For r = 2 To UBound(out, 1)
        DictIncrement counts, CStr(out(r, 4)), 1
    Next r

    ReportRow rows, DOCTOR_NAME & " - Match summary"
    ReportRow rows, "From", fromRange.Parent.Name & "!" & fromRange.Address(False, False)
    ReportRow rows, "To", toRange.Parent.Name & "!" & toRange.Address(False, False)
    ReportRow rows, "Weak below", threshold & "%"
    ReportRow rows, "Company names", IIf(stripCompany, "Yes - decoration ignored", "No")

    ReportHeading rows, "Verdicts"
    ReportRow rows, "Verdict", "Count", "What it means"
    ReportRow rows, "Exact", DictGet(counts, "Exact", 0), "The same string once tidied up"
    ReportRow rows, "Same company", DictGet(counts, "Same company", 0), "The same once Ltd, Group and punctuation are set aside"
    ReportRow rows, "Likely", DictGet(counts, "Likely", 0), "Scored at or above " & threshold & "%"
    ReportRow rows, "Weak - check", DictGet(counts, "Weak - check", 0), "Within 15 points of the line. Read these"
    ReportRow rows, "No match", DictGet(counts, "No match", 0), "Nothing came close"
    ReportRow rows, "Blank", DictGet(counts, "Blank", 0), "Nothing there to match"

    ReportHeading rows, "How to use this"
    ReportRow rows, "Sort", "Sort the Matches sheet by Verdict, then by Score. Everything " & _
              "needing a decision is then together."
    ReportRow rows, "Next best", "When the best and next best scores are close, the match is a " & _
              "coin toss however high the score."
    ReportRow rows, "Your call", "Put the answer in that column. Nothing overwrites it."

    Set BuildMatchReport = rows
End Function
