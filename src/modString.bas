Attribute VB_Name = "modString"
'==============================================================================
' modString - text helpers
'------------------------------------------------------------------------------
' Cleaning, testing, slicing, padding, regular expressions and fuzzy matching.
'
' Standalone: no dependency on the other modules in this repository.
' RegEx helpers are late bound (VBScript.RegExp) so no project reference is
' required.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' CleanText
' Converts non-breaking spaces to normal spaces, strips control characters and
' collapses runs of whitespace to a single space. Safe with Empty, Null and
' cell error values.
'
'   CleanText("  Total" & Chr(160) & "  assets ")   ->  "Total assets"
'------------------------------------------------------------------------------
Public Function CleanText(ByVal value As Variant) As String
    Dim s As String
    Dim buf As String
    Dim i As Long, n As Long
    Dim ch As String
    Dim code As Long

    If IsError(value) Then Exit Function
    If IsArray(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    If IsObject(value) Then Exit Function

    s = CStr(value)
    If Len(s) = 0 Then Exit Function

    buf = Space$(Len(s))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 9, 10, 11, 12, 13, 32, 160, 8194, 8195, 8201, 8239
                n = n + 1
                Mid$(buf, n, 1) = " "
            Case 0 To 31, 127
                ' control character - dropped
            Case Else
                n = n + 1
                Mid$(buf, n, 1) = ch
        End Select
    Next i

    s = Left$(buf, n)
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop

    CleanText = Trim$(s)
End Function

'------------------------------------------------------------------------------
' IsBlank
' True when the value is Empty, Null, an error, or contains nothing but
' whitespace. Useful for "is this cell really empty" tests.
'------------------------------------------------------------------------------
Public Function IsBlank(ByVal value As Variant) As Boolean
    If IsError(value) Then IsBlank = True: Exit Function
    If IsNull(value) Or IsEmpty(value) Then IsBlank = True: Exit Function
    If IsObject(value) Then IsBlank = (value Is Nothing): Exit Function
    If IsArray(value) Then IsBlank = False: Exit Function
    IsBlank = (Len(CleanText(value)) = 0)
End Function

'------------------------------------------------------------------------------
' StartsWith / EndsWith / ContainsText
' Case insensitive by default.
'------------------------------------------------------------------------------
Public Function StartsWith(ByVal text As String, ByVal prefix As String, _
                           Optional ByVal caseSensitive As Boolean = False) As Boolean
    If Len(prefix) = 0 Then StartsWith = True: Exit Function
    If Len(prefix) > Len(text) Then Exit Function
    StartsWith = (StrComp(Left$(text, Len(prefix)), prefix, CompareMode(caseSensitive)) = 0)
End Function

Public Function EndsWith(ByVal text As String, ByVal suffix As String, _
                         Optional ByVal caseSensitive As Boolean = False) As Boolean
    If Len(suffix) = 0 Then EndsWith = True: Exit Function
    If Len(suffix) > Len(text) Then Exit Function
    EndsWith = (StrComp(Right$(text, Len(suffix)), suffix, CompareMode(caseSensitive)) = 0)
End Function

Public Function ContainsText(ByVal text As String, ByVal needle As String, _
                             Optional ByVal caseSensitive As Boolean = False) As Boolean
    If Len(needle) = 0 Then ContainsText = True: Exit Function
    ContainsText = (InStr(1, text, needle, CompareMode(caseSensitive)) > 0)
End Function

'------------------------------------------------------------------------------
' TextBefore / TextAfter / TextBeforeLast / TextAfterLast
' Slice a string around a delimiter. When the delimiter is not found the whole
' string is returned unless ifNotFound is supplied.
'
'   TextAfter("client_2024_Q3.xlsx", "_")        ->  "2024_Q3.xlsx"
'   TextAfterLast("C:\data\report.xlsx", "\")    ->  "report.xlsx"
'------------------------------------------------------------------------------
Public Function TextBefore(ByVal text As String, ByVal delimiter As String, _
                           Optional ByVal ifNotFound As Variant) As String
    Dim p As Long
    p = InStr(1, text, delimiter, vbTextCompare)
    If p = 0 Then
        TextBefore = IIf(IsMissing(ifNotFound), text, CStr(ifNotFound))
    Else
        TextBefore = Left$(text, p - 1)
    End If
End Function

Public Function TextAfter(ByVal text As String, ByVal delimiter As String, _
                          Optional ByVal ifNotFound As Variant) As String
    Dim p As Long
    p = InStr(1, text, delimiter, vbTextCompare)
    If p = 0 Then
        TextAfter = IIf(IsMissing(ifNotFound), text, CStr(ifNotFound))
    Else
        TextAfter = Mid$(text, p + Len(delimiter))
    End If
End Function

Public Function TextBeforeLast(ByVal text As String, ByVal delimiter As String, _
                               Optional ByVal ifNotFound As Variant) As String
    Dim p As Long
    p = InStrRev(text, delimiter, -1, vbTextCompare)
    If p = 0 Then
        TextBeforeLast = IIf(IsMissing(ifNotFound), text, CStr(ifNotFound))
    Else
        TextBeforeLast = Left$(text, p - 1)
    End If
End Function

Public Function TextAfterLast(ByVal text As String, ByVal delimiter As String, _
                              Optional ByVal ifNotFound As Variant) As String
    Dim p As Long
    p = InStrRev(text, delimiter, -1, vbTextCompare)
    If p = 0 Then
        TextAfterLast = IIf(IsMissing(ifNotFound), text, CStr(ifNotFound))
    Else
        TextAfterLast = Mid$(text, p + Len(delimiter))
    End If
End Function

'------------------------------------------------------------------------------
' PadLeft / PadRight / Repeat
'------------------------------------------------------------------------------
Public Function PadLeft(ByVal text As String, ByVal totalLength As Long, _
                        Optional ByVal padChar As String = " ") As String
    If Len(padChar) = 0 Then padChar = " "
    If Len(text) >= totalLength Then PadLeft = text: Exit Function
    PadLeft = Right$(String$(totalLength, Left$(padChar, 1)) & text, totalLength)
End Function

Public Function PadRight(ByVal text As String, ByVal totalLength As Long, _
                         Optional ByVal padChar As String = " ") As String
    If Len(padChar) = 0 Then padChar = " "
    If Len(text) >= totalLength Then PadRight = text: Exit Function
    PadRight = Left$(text & String$(totalLength, Left$(padChar, 1)), totalLength)
End Function

Public Function Repeat(ByVal text As String, ByVal count As Long) As String
    Dim i As Long
    For i = 1 To count
        Repeat = Repeat & text
    Next i
End Function

'------------------------------------------------------------------------------
' JoinNonBlank
' Joins the non-blank elements of an array or Collection with a delimiter.
'
'   JoinNonBlank(Array("Smith", "", "London"), ", ")   ->  "Smith, London"
'------------------------------------------------------------------------------
Public Function JoinNonBlank(ByVal values As Variant, Optional ByVal delimiter As String = ", ") As String
    Dim item As Variant
    Dim parts As String
    Dim first As Boolean

    first = True
    If IsObject(values) Or IsArray(values) Then
        For Each item In values
            If Not IsBlank(item) Then
                If first Then
                    parts = CleanText(item)
                    first = False
                Else
                    parts = parts & delimiter & CleanText(item)
                End If
            End If
        Next item
    Else
        parts = CleanText(values)
    End If
    JoinNonBlank = parts
End Function

'------------------------------------------------------------------------------
' KeepDigits / ExtractNumber
'
'   KeepDigits("GB-12 345 678")     ->  "12345678"
'   ExtractNumber("Fee: 1,250.75 GBP") ->  1250.75
'------------------------------------------------------------------------------
Public Function KeepDigits(ByVal text As String) As String
    Dim i As Long, ch As String, buf As String, n As Long
    buf = Space$(Len(text))
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then
            n = n + 1
            Mid$(buf, n, 1) = ch
        End If
    Next i
    KeepDigits = Left$(buf, n)
End Function

Public Function ExtractNumber(ByVal text As String, Optional ByVal defaultValue As Double = 0) As Double
    Dim s As String
    s = RegExExtract(Replace$(text, ",", ""), "-?\d+(\.\d+)?")
    If Len(s) = 0 Then
        ExtractNumber = defaultValue
    ElseIf IsNumeric(s) Then
        ExtractNumber = CDbl(s)
    Else
        ExtractNumber = defaultValue
    End If
End Function

'------------------------------------------------------------------------------
' TitleCase
' Like StrConv(..., vbProperCase) but leaves the rest of a word alone after an
' apostrophe and capitalises after hyphens.
'
'   TitleCase("o'neill smith-jones")  ->  "O'Neill Smith-Jones"
'------------------------------------------------------------------------------
Public Function TitleCase(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    Dim capNext As Boolean
    Dim out As String

    text = LCase$(CleanText(text))
    capNext = True
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If capNext And ch Like "[a-z]" Then
            out = out & UCase$(ch)
            capNext = False
        Else
            out = out & ch
            If InStr(" -'/(", ch) > 0 Then capNext = True
        End If
    Next i
    TitleCase = out
End Function

'------------------------------------------------------------------------------
' Regular expression helpers (late bound VBScript.RegExp)
'
'   RegExTest("INV-2024-001", "^INV-\d{4}-\d{3}$")             ->  True
'   RegExExtract("Invoice INV-2024-001", "INV-(\d{4})", 1)      ->  "2024"
'   RegExReplace("a1b2c3", "\d", "")                            ->  "abc"
'------------------------------------------------------------------------------
Public Function RegExTest(ByVal text As String, ByVal pattern As String, _
                          Optional ByVal ignoreCase As Boolean = True) As Boolean
    Dim re As Object
    Set re = NewRegExp(pattern, ignoreCase, True)
    If re Is Nothing Then Exit Function
    RegExTest = re.Test(text)
End Function

Public Function RegExReplace(ByVal text As String, ByVal pattern As String, ByVal replacement As String, _
                             Optional ByVal ignoreCase As Boolean = True, _
                             Optional ByVal replaceAll As Boolean = True) As String
    Dim re As Object
    Set re = NewRegExp(pattern, ignoreCase, replaceAll)
    If re Is Nothing Then RegExReplace = text: Exit Function
    RegExReplace = re.Replace(text, replacement)
End Function

' groupIndex 0 returns the whole match, 1 the first capture group, and so on.
Public Function RegExExtract(ByVal text As String, ByVal pattern As String, _
                             Optional ByVal groupIndex As Long = 0, _
                             Optional ByVal ignoreCase As Boolean = True) As String
    Dim re As Object
    Dim matches As Object
    Dim m As Object

    Set re = NewRegExp(pattern, ignoreCase, True)
    If re Is Nothing Then Exit Function
    Set matches = re.Execute(text)
    If matches.count = 0 Then Exit Function

    Set m = matches(0)
    If groupIndex <= 0 Then
        RegExExtract = m.value
    ElseIf groupIndex <= m.SubMatches.count Then
        RegExExtract = CStr(m.SubMatches(groupIndex - 1) & "")
    End If
End Function

' Returns every match as a 1-based 1D array. Empty array when there is no match.
Public Function RegExMatches(ByVal text As String, ByVal pattern As String, _
                             Optional ByVal ignoreCase As Boolean = True) As Variant
    Dim re As Object
    Dim matches As Object
    Dim out() As String
    Dim i As Long

    Set re = NewRegExp(pattern, ignoreCase, True)
    If re Is Nothing Then RegExMatches = Array(): Exit Function
    Set matches = re.Execute(text)
    If matches.count = 0 Then RegExMatches = Array(): Exit Function

    ReDim out(1 To matches.count)
    For i = 0 To matches.count - 1
        out(i + 1) = matches(i).value
    Next i
    RegExMatches = out
End Function

'------------------------------------------------------------------------------
' Levenshtein / SimilarityPercent
' Fuzzy matching - handy for reconciling names that almost agree.
'
'   Levenshtein("Barclays", "Barclay's")        ->  1
'   SimilarityPercent("Barclays", "Barclay's")  ->  88.9
'------------------------------------------------------------------------------
Public Function Levenshtein(ByVal text1 As String, ByVal text2 As String, _
                            Optional ByVal ignoreCase As Boolean = True) As Long
    Dim prev() As Long, curr() As Long
    Dim i As Long, j As Long, cost As Long
    Dim len1 As Long, len2 As Long

    If ignoreCase Then
        text1 = LCase$(text1)
        text2 = LCase$(text2)
    End If
    len1 = Len(text1)
    len2 = Len(text2)
    If len1 = 0 Then Levenshtein = len2: Exit Function
    If len2 = 0 Then Levenshtein = len1: Exit Function

    ReDim prev(0 To len2)
    ReDim curr(0 To len2)
    For j = 0 To len2
        prev(j) = j
    Next j

    For i = 1 To len1
        curr(0) = i
        For j = 1 To len2
            If Mid$(text1, i, 1) = Mid$(text2, j, 1) Then cost = 0 Else cost = 1
            curr(j) = MinOf3(curr(j - 1) + 1, prev(j) + 1, prev(j - 1) + cost)
        Next j
        For j = 0 To len2
            prev(j) = curr(j)
        Next j
    Next i

    Levenshtein = prev(len2)
End Function

' 100 = identical, 0 = nothing in common.
Public Function SimilarityPercent(ByVal text1 As String, ByVal text2 As String, _
                                  Optional ByVal ignoreCase As Boolean = True) As Double
    Dim longest As Long
    longest = IIf(Len(text1) > Len(text2), Len(text1), Len(text2))
    If longest = 0 Then SimilarityPercent = 100: Exit Function
    SimilarityPercent = (1 - Levenshtein(text1, text2, ignoreCase) / longest) * 100
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Function CompareMode(ByVal caseSensitive As Boolean) As VbCompareMethod
    CompareMode = IIf(caseSensitive, vbBinaryCompare, vbTextCompare)
End Function

Private Function NewRegExp(ByVal pattern As String, ByVal ignoreCase As Boolean, _
                           ByVal globalMatch As Boolean) As Object
    Dim engine As Object

    On Error GoTo Failed
    Set engine = CreateObject("VBScript.RegExp")
    engine.pattern = pattern
    engine.IgnoreCase = ignoreCase
    engine.Global = globalMatch
    Set NewRegExp = engine
    Exit Function
Failed:
    Set NewRegExp = Nothing
End Function

Private Function MinOf3(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    MinOf3 = a
    If b < MinOf3 Then MinOf3 = b
    If c < MinOf3 Then MinOf3 = c
End Function
