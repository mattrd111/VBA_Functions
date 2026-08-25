Attribute VB_Name = "modAuditFormula"
'==============================================================================
' modAuditFormula - reading a formula the way an auditor reads it
'------------------------------------------------------------------------------
' Walks formula text once and reports what matters: numbers typed into the
' middle of a calculation, volatile functions, whole-column references and
' links to other workbooks.
'
' The walker keeps a stack of the functions it is inside, which is what makes
' the hardcode check usable: the 3 in VLOOKUP(x, y, 3, FALSE) is a column index
' and nobody wants to see it flagged, while the 1.03 in B5*1.03 is an
' assumption someone buried in a formula and is exactly what you are looking
' for.
'
' Depends on: modDictionary (from \src)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' AnalyseFormula
' One pass over one formula. Returns a dictionary:
'   "Hardcodes"    Collection of Array(literalText, containingFunction, isDirectArgument)
'   "Volatiles"    Collection of volatile function names used
'   "WholeRefs"    Collection of whole-column references such as A:A
'   "External"     count of references to another workbook
'   "IfError"      True when the formula wraps something in IFERROR
'   "Functions"    count of function calls, as a rough complexity measure
'------------------------------------------------------------------------------
Public Function AnalyseFormula(ByVal formulaText As String) As Object
    Dim result As Object
    Dim hardcodes As Collection, volatiles As Collection, wholeRefs As Collection
    Dim stack() As String
    Dim depth As Long
    Dim i As Long, n As Long, start As Long
    Dim ch As String, token As String, quoted As String
    Dim externalRefs As Long, functionCount As Long
    Dim hasIfError As Boolean
    Dim pendingName As String

    Set result = NewDictionary(True)
    Set hardcodes = New Collection
    Set volatiles = New Collection
    Set wholeRefs = New Collection

    result.Add "Hardcodes", hardcodes
    result.Add "Volatiles", volatiles
    result.Add "WholeRefs", wholeRefs
    result.Add "External", 0
    result.Add "IfError", False
    result.Add "Functions", 0
    Set AnalyseFormula = result

    n = Len(formulaText)
    If n = 0 Then Exit Function
    If Left$(formulaText, 1) <> "=" Then Exit Function

    ReDim stack(0 To 128)
    i = 2                                       ' step over the leading =

    Do While i <= n
        ch = Mid$(formulaText, i, 1)

        If ch = """" Then
            ' A text literal. Doubled quotes inside simply end one and start
            ' another, so stopping at the next quote is correct either way.
            i = i + 1
            Do While i <= n
                If Mid$(formulaText, i, 1) = """" Then Exit Do
                i = i + 1
            Loop
            i = i + 1

        ElseIf ch = "'" Then
            ' A quoted sheet or workbook name. If it holds [something] then the
            ' formula is reaching into another file.
            start = i + 1
            i = i + 1
            Do While i <= n
                If Mid$(formulaText, i, 1) = "'" Then Exit Do
                i = i + 1
            Loop
            quoted = Mid$(formulaText, start, i - start)
            If InStr(quoted, "[") > 0 Then externalRefs = externalRefs + 1
            i = i + 1

        ElseIf ch = "[" Then
            ' [Book1.xlsx]Sheet1!A1 when it opens a term, Table1[Column] when it
            ' follows a name.
            If Not FollowsIdentifier(formulaText, i) Then externalRefs = externalRefs + 1
            Do While i <= n
                If Mid$(formulaText, i, 1) = "]" Then Exit Do
                i = i + 1
            Loop
            i = i + 1

        ElseIf IsIdentifierStart(ch) Then
            start = i
            Do While i <= n
                If Not IsIdentifierChar(Mid$(formulaText, i, 1)) Then Exit Do
                i = i + 1
            Loop
            token = Mid$(formulaText, start, i - start)

            If NextRealCharacter(formulaText, i) = "(" Then
                functionCount = functionCount + 1
                pendingName = UCase$(token)
                If IsVolatileFunction(token) Then volatiles.Add token
                If StrComp(token, "IFERROR", vbTextCompare) = 0 Then hasIfError = True
                If StrComp(token, "IFNA", vbTextCompare) = 0 Then hasIfError = True
            Else
                If IsWholeColumnReference(token) Then wholeRefs.Add token
                If InStr(token, "!") > 0 And InStr(token, "[") > 0 Then
                    externalRefs = externalRefs + 1
                End If
            End If

        ElseIf IsDigit(ch) Or (ch = "." And IsDigit(NextCharacter(formulaText, i + 1))) Then
            start = i
            i = ConsumeNumber(formulaText, i)
            token = Mid$(formulaText, start, i - start)
            hardcodes.Add Array(token, CurrentFunction(stack, depth), _
                                IsDirectArgument(formulaText, start, i))

        ElseIf ch = "(" Then
            depth = depth + 1
            If depth <= UBound(stack) Then stack(depth) = pendingName
            pendingName = ""
            i = i + 1

        ElseIf ch = ")" Then
            If depth > 0 Then
                If depth <= UBound(stack) Then stack(depth) = ""
                depth = depth - 1
            End If
            i = i + 1

        Else
            i = i + 1
        End If
    Loop

    result("External") = externalRefs
    result("IfError") = hasIfError
    result("Functions") = functionCount
End Function

'------------------------------------------------------------------------------
' IsSuspectConstant
' True when a literal found in a formula is worth a look: not a round number
' everyone uses, and not sitting where a function wants an index or a count.
'------------------------------------------------------------------------------
Public Function IsSuspectConstant(ByVal literalText As String, ByVal containingFunction As String, _
                                  ByVal isDirectArgument As Boolean) As Boolean
    If Len(literalText) = 0 Then Exit Function
    If Not IsNumeric(literalText) Then Exit Function
    If IsCommonConstant(literalText) Then Exit Function

    ' The 3 in VLOOKUP(x, y, 3, FALSE) is a column index and nobody wants to see
    ' it. The 1.2 in ROUND(B5*1.2, 2) is an assumption and everybody does - so
    ' only a number standing alone as an argument gets the benefit of the doubt.
    If IsStructuralFunction(containingFunction) And isDirectArgument Then Exit Function

    IsSuspectConstant = True
End Function

'------------------------------------------------------------------------------
' ConstantSeverity
' A number with a decimal point, or one big enough to be money, is an
' assumption. A bare small integer is more often a counter.
'------------------------------------------------------------------------------
Public Function ConstantSeverity(ByVal literalText As String) As String
    Dim value As Double

    If InStr(literalText, ".") > 0 Then ConstantSeverity = "High": Exit Function
    If Not IsNumeric(literalText) Then ConstantSeverity = "Medium": Exit Function

    value = Abs(CDbl(literalText))
    If value >= 1000 Then ConstantSeverity = "High" Else ConstantSeverity = "Medium"
End Function

'------------------------------------------------------------------------------
' Function classification
'------------------------------------------------------------------------------
' Recalculate on every change, whatever they depend on. A model full of these
' is a model that takes ten seconds to breathe.
Public Function IsVolatileFunction(ByVal functionName As String) As Boolean
    Select Case UCase$(functionName)
        Case "NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", _
             "OFFSET", "INDIRECT", "CELL", "INFO"
            IsVolatileFunction = True
    End Select
End Function

' Functions whose numeric arguments are positions, counts or flags rather than
' assumptions, so a number sitting in one is not a hardcode.
Public Function IsStructuralFunction(ByVal functionName As String) As Boolean
    Select Case UCase$(functionName)
        Case "INDEX", "MATCH", "XMATCH", "VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP", _
             "OFFSET", "INDIRECT", "CHOOSE", "COLUMN", "COLUMNS", "ROW", "ROWS", _
             "ROUND", "ROUNDUP", "ROUNDDOWN", "MROUND", "CEILING", "FLOOR", "TRUNC", _
             "LEFT", "RIGHT", "MID", "FIND", "SEARCH", "SUBSTITUTE", "REPT", "TEXT", _
             "SUBTOTAL", "AGGREGATE", "LARGE", "SMALL", "RANK", "PERCENTILE", "QUARTILE", _
             "DATE", "TIME", "WEEKDAY", "WORKDAY", "NETWORKDAYS", "EOMONTH", "EDATE", _
             "YEARFRAC", "DAYS360", "CELL", "INFO", "ERROR.TYPE", _
             "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS", _
             "MAXIFS", "MINIFS", "TEXTJOIN", "SORT", "SORTBY", "TAKE", "DROP", "WRAPROWS"
            IsStructuralFunction = True
    End Select
End Function

' Numbers that mean something to everyone and are not worth reporting.
Public Function IsCommonConstant(ByVal literalText As String) As Boolean
    Dim value As Double

    If Not IsNumeric(literalText) Then Exit Function
    value = CDbl(literalText)

    Select Case value
        Case 0, 1, 2, 12, 100, 360, 365, 366, 1000, 10000, 1000000, 1000000000#
            IsCommonConstant = True
    End Select
End Function

'------------------------------------------------------------------------------
' IsWholeColumnReference
' A:A, $C:$F, Sheet1!A:A - the usual reason a SUMIF takes a second per cell.
'------------------------------------------------------------------------------
Public Function IsWholeColumnReference(ByVal token As String) As Boolean
    Dim parts() As String
    Dim bare As String
    Dim p As Long

    If InStr(token, ":") = 0 Then Exit Function

    bare = token
    p = InStrRev(bare, "!")
    If p > 0 Then bare = Mid$(bare, p + 1)

    parts = Split(bare, ":")
    If UBound(parts) <> 1 Then Exit Function

    IsWholeColumnReference = IsColumnOnly(parts(0)) And IsColumnOnly(parts(1))
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Function IsColumnOnly(ByVal text As String) As Boolean
    Dim i As Long
    Dim ch As String

    text = Replace$(text, "$", "")
    If Len(text) = 0 Or Len(text) > 3 Then Exit Function

    For i = 1 To Len(text)
        ch = UCase$(Mid$(text, i, 1))
        If ch < "A" Or ch > "Z" Then Exit Function
    Next i
    IsColumnOnly = True
End Function

Private Function CurrentFunction(ByRef stack() As String, ByVal depth As Long) As String
    Dim level As Long

    For level = depth To 1 Step -1
        If level <= UBound(stack) Then
            If Len(stack(level)) > 0 Then CurrentFunction = stack(level): Exit Function
        End If
    Next level
End Function

Private Function ConsumeNumber(ByVal text As String, ByVal start As Long) As Long
    Dim i As Long, n As Long
    Dim ch As String
    Dim seenDot As Boolean

    n = Len(text)
    i = start

    Do While i <= n
        ch = Mid$(text, i, 1)
        If IsDigit(ch) Then
            i = i + 1
        ElseIf ch = "." And Not seenDot Then
            seenDot = True
            i = i + 1
        ElseIf (ch = "E" Or ch = "e") And IsExponent(text, i) Then
            i = i + 2
            If IsDigit(Mid$(text, i, 1)) Then i = i + 1
            Do While i <= n
                If Not IsDigit(Mid$(text, i, 1)) Then Exit Do
                i = i + 1
            Loop
            Exit Do
        Else
            Exit Do
        End If
    Loop

    ConsumeNumber = i
End Function

' 2E5 and 2E-5 are numbers; 2*E5 is arithmetic on a cell, but the * gets there
' first, so only a digit or sign directly after the E counts.
Private Function IsExponent(ByVal text As String, ByVal position As Long) As Boolean
    Dim nextCh As String

    nextCh = NextCharacter(text, position + 1)
    If IsDigit(nextCh) Then IsExponent = True: Exit Function
    If nextCh = "+" Or nextCh = "-" Then
        IsExponent = IsDigit(NextCharacter(text, position + 2))
    End If
End Function

Private Function NextCharacter(ByVal text As String, ByVal position As Long) As String
    If position < 1 Or position > Len(text) Then Exit Function
    NextCharacter = Mid$(text, position, 1)
End Function

' True when the number is the whole argument rather than part of a sum. A
' leading sign still counts as part of the argument.
Private Function IsDirectArgument(ByVal text As String, ByVal start As Long, _
                                  ByVal afterEnd As Long) As Boolean
    Dim position As Long
    Dim ch As String

    position = start
    ch = PreviousRealCharacter(text, position)
    If ch = "+" Or ch = "-" Then ch = PreviousRealCharacter(text, position)
    If ch <> "(" And ch <> "," Then Exit Function

    ch = NextRealCharacter(text, afterEnd)
    IsDirectArgument = (ch = "," Or ch = ")")
End Function

' Walks back past spaces and reports where it stopped, so the caller can step
' back again over a sign.
Private Function PreviousRealCharacter(ByVal text As String, ByRef position As Long) As String
    Dim i As Long

    For i = position - 1 To 1 Step -1
        If Mid$(text, i, 1) <> " " Then
            position = i
            PreviousRealCharacter = Mid$(text, i, 1)
            Exit Function
        End If
    Next i
    position = 0
End Function

Private Function NextRealCharacter(ByVal text As String, ByVal position As Long) As String
    Dim i As Long

    For i = position To Len(text)
        If Mid$(text, i, 1) <> " " Then
            NextRealCharacter = Mid$(text, i, 1)
            Exit Function
        End If
    Next i
End Function

Private Function FollowsIdentifier(ByVal text As String, ByVal position As Long) As Boolean
    If position <= 1 Then Exit Function
    FollowsIdentifier = IsIdentifierChar(Mid$(text, position - 1, 1))
End Function

Private Function IsDigit(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsDigit = (ch >= "0" And ch <= "9")
End Function

Private Function IsIdentifierStart(ByVal ch As String) As Boolean
    Select Case ch
        Case "A" To "Z", "a" To "z", "_", "$", "\"
            IsIdentifierStart = True
        Case Else
            IsIdentifierStart = (AscW(ch) > 127)
    End Select
End Function

' Deliberately generous: ! and : are pulled in so that Sheet1!A1 and A:A arrive
' as single tokens.
Private Function IsIdentifierChar(ByVal ch As String) As Boolean
    Select Case ch
        Case "A" To "Z", "a" To "z", "0" To "9", "_", ".", "$", "!", ":", "\", "?"
            IsIdentifierChar = True
        Case Else
            IsIdentifierChar = (AscW(ch) > 127)
    End Select
End Function
