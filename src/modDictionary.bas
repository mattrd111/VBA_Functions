Attribute VB_Name = "modDictionary"
'==============================================================================
' modDictionary - fast lookups, grouping and counting
'------------------------------------------------------------------------------
' Dictionary based replacements for VLOOKUP, SUMIF and COUNTIF loops. Building
' an index once and reading it back is orders of magnitude faster than a lookup
' per row on a large sheet.
'
' Standalone: no dependency on the other modules in this repository.
' Uses late-bound Scripting.Dictionary, so no project reference is needed.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' NewDictionary
' A dictionary that ignores case by default.
'------------------------------------------------------------------------------
Public Function NewDictionary(Optional ByVal ignoreCase As Boolean = True) As Object
    Dim result As Object

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = IIf(ignoreCase, 1, 0)      ' 1 = TextCompare, 0 = BinaryCompare
    Set NewDictionary = result
End Function

'------------------------------------------------------------------------------
' DictGet
' The item for a key, or a default when the key is absent - without the
' side effect of dict(key) on a missing key, which silently adds it.
'------------------------------------------------------------------------------
Public Function DictGet(ByVal source As Object, ByVal key As Variant, _
                        Optional ByVal defaultValue As Variant = "") As Variant
    If source Is Nothing Then DictGet = defaultValue: Exit Function

    If source.Exists(key) Then
        If IsObject(source(key)) Then
            Set DictGet = source(key)
        Else
            DictGet = source(key)
        End If
    Else
        DictGet = defaultValue
    End If
End Function

'------------------------------------------------------------------------------
' DictIncrement
' Adds to a running total, creating the key when it is new. The building block
' for counting and totalling in a loop.
'
'   DictIncrement totals, region, amount
'------------------------------------------------------------------------------
Public Sub DictIncrement(ByVal target As Object, ByVal key As Variant, _
                         Optional ByVal amount As Double = 1)
    If target Is Nothing Then Exit Sub

    If target.Exists(key) Then
        target(key) = target(key) + amount
    Else
        target.Add key, amount
    End If
End Sub

'------------------------------------------------------------------------------
' DictAppend
' Collects values under a key, so one key holds a Collection of everything seen
' for it - a group-by that keeps the detail rather than a total.
'
'   DictAppend byRegion, "North", invoiceNumber
'   For Each inv In byRegion("North") ...
'------------------------------------------------------------------------------
Public Sub DictAppend(ByVal target As Object, ByVal key As Variant, ByVal value As Variant)
    Dim bucket As Collection

    If target Is Nothing Then Exit Sub

    If target.Exists(key) Then
        Set bucket = target(key)
    Else
        Set bucket = New Collection
        target.Add key, bucket
    End If
    bucket.Add value
End Sub

'------------------------------------------------------------------------------
' SortedKeys
' Dictionary keys as a sorted 1-based array - dictionaries keep insertion
' order, which is rarely the order you want to report in.
'------------------------------------------------------------------------------
Public Function SortedKeys(ByVal source As Object, Optional ByVal descending As Boolean = False) As Variant
    Dim keys As Variant
    Dim out() As Variant
    Dim i As Long, n As Long

    If source Is Nothing Then SortedKeys = Array(): Exit Function
    If source.Count = 0 Then SortedKeys = Array(): Exit Function

    keys = source.keys
    n = UBound(keys) - LBound(keys) + 1
    ReDim out(1 To n)
    For i = 1 To n
        out(i) = keys(LBound(keys) + i - 1)
    Next i

    SortKeys out, 1, n
    If descending Then out = Reversed(out)
    SortedKeys = out
End Function

'------------------------------------------------------------------------------
' DictToArray
' Dictionary as a two column 1-based array, ready to write straight to a sheet.
' Object items (for example the Collections built by DictAppend) are reported
' as their item count.
'------------------------------------------------------------------------------
Public Function DictToArray(ByVal source As Object, Optional ByVal sortByKey As Boolean = True) As Variant
    Dim out As Variant
    Dim keys As Variant
    Dim i As Long, n As Long
    Dim item As Variant

    If source Is Nothing Then Exit Function
    If source.Count = 0 Then Exit Function

    If sortByKey Then
        keys = SortedKeys(source)
    Else
        keys = source.keys
    End If

    n = UBound(keys) - LBound(keys) + 1
    ReDim out(1 To n, 1 To 2)
    For i = 1 To n
        out(i, 1) = keys(LBound(keys) + i - 1)
        If source.Exists(out(i, 1)) Then
            If IsObject(source(out(i, 1))) Then
                Set item = source(out(i, 1))
                On Error Resume Next
                out(i, 2) = item.Count
                On Error GoTo 0
            Else
                out(i, 2) = source(out(i, 1))
            End If
        End If
    Next i

    DictToArray = out
End Function

'------------------------------------------------------------------------------
' BuildLookup
' Indexes a range or 2D array by one column so you can look values up in
' constant time - a VLOOKUP over 100,000 rows done once instead of per row.
' Read it back with LookupValue, which matches keys the same way.
'
'   Set prices = BuildLookup(ws.Range("A1:D5000"), 1, 3)
'   price = LookupValue(prices, code, 0)
'
' Numbers stored as text match real numbers, but text that merely looks numeric
' keeps its leading zeros, so "007" and 7 stay apart.
'------------------------------------------------------------------------------
Public Function BuildLookup(ByVal source As Variant, _
                            Optional ByVal keyColumn As Long = 1, _
                            Optional ByVal returnColumn As Long = 2, _
                            Optional ByVal skipHeaderRow As Boolean = True, _
                            Optional ByVal keepFirstDuplicate As Boolean = True) As Object
    Dim data As Variant
    Dim index As Object
    Dim r As Long, firstRow As Long
    Dim key As String

    Set index = NewDictionary(True)
    Set BuildLookup = index

    data = AsArray(source)
    If IsEmpty(data) Then Exit Function
    If keyColumn < LBound(data, 2) Or keyColumn > UBound(data, 2) Then Exit Function
    If returnColumn < LBound(data, 2) Or returnColumn > UBound(data, 2) Then Exit Function

    firstRow = LBound(data, 1)
    If skipHeaderRow Then firstRow = firstRow + 1

    For r = firstRow To UBound(data, 1)
        key = KeyOf(data(r, keyColumn))
        If Len(key) > 0 Then
            If index.Exists(key) Then
                If Not keepFirstDuplicate Then index(key) = data(r, returnColumn)
            Else
                index.Add key, data(r, returnColumn)
            End If
        End If
    Next r
End Function

'------------------------------------------------------------------------------
' LookupValue
' Reads back an index built by BuildLookup, normalising the key the same way.
'------------------------------------------------------------------------------
Public Function LookupValue(ByVal index As Object, ByVal key As Variant, _
                            Optional ByVal defaultValue As Variant = "") As Variant
    Dim k As String

    If index Is Nothing Then LookupValue = defaultValue: Exit Function
    k = KeyOf(key)
    If Len(k) = 0 Then LookupValue = defaultValue: Exit Function

    If index.Exists(k) Then
        LookupValue = index(k)
    Else
        LookupValue = defaultValue
    End If
End Function

'------------------------------------------------------------------------------
' GroupSum
' Totals one column by another - SUMIF for every key in a single pass over the
' data. Keys keep the value they have in the sheet, so the result can be
' written straight out with DictToArray.
'
'   Set byRegion = GroupSum(ws.Range("A1:D5000"), 2, 4)
'------------------------------------------------------------------------------
Public Function GroupSum(ByVal source As Variant, ByVal keyColumn As Long, ByVal valueColumn As Long, _
                         Optional ByVal skipHeaderRow As Boolean = True) As Object
    Dim data As Variant
    Dim totals As Object
    Dim r As Long, firstRow As Long
    Dim key As Variant

    Set totals = NewDictionary(True)
    Set GroupSum = totals

    data = AsArray(source)
    If IsEmpty(data) Then Exit Function
    If keyColumn < LBound(data, 2) Or keyColumn > UBound(data, 2) Then Exit Function
    If valueColumn < LBound(data, 2) Or valueColumn > UBound(data, 2) Then Exit Function

    firstRow = LBound(data, 1)
    If skipHeaderRow Then firstRow = firstRow + 1

    For r = firstRow To UBound(data, 1)
        key = data(r, keyColumn)
        If Not IsBlankValue(key) Then
            If IsNumericValue(data(r, valueColumn)) Then
                DictIncrement totals, key, CDbl(data(r, valueColumn))
            ElseIf Not totals.Exists(key) Then
                totals.Add key, 0
            End If
        End If
    Next r
End Function

'------------------------------------------------------------------------------
' CountBy
' Row count per distinct value of a column - COUNTIF for every key at once.
'------------------------------------------------------------------------------
Public Function CountBy(ByVal source As Variant, ByVal keyColumn As Long, _
                        Optional ByVal skipHeaderRow As Boolean = True) As Object
    Dim data As Variant
    Dim counts As Object
    Dim r As Long, firstRow As Long
    Dim key As Variant

    Set counts = NewDictionary(True)
    Set CountBy = counts

    data = AsArray(source)
    If IsEmpty(data) Then Exit Function
    If keyColumn < LBound(data, 2) Or keyColumn > UBound(data, 2) Then Exit Function

    firstRow = LBound(data, 1)
    If skipHeaderRow Then firstRow = firstRow + 1

    For r = firstRow To UBound(data, 1)
        key = data(r, keyColumn)
        If Not IsBlankValue(key) Then DictIncrement counts, key, 1
    Next r
End Function

'------------------------------------------------------------------------------
' UniqueValues
' Distinct non-blank values of a range or array, in first seen order, as a
' 1-based array.
'------------------------------------------------------------------------------
Public Function UniqueValues(ByVal source As Variant, Optional ByVal ignoreCase As Boolean = True) As Variant
    Dim seen As Object
    Dim data As Variant
    Dim item As Variant
    Dim out() As Variant
    Dim i As Long

    Set seen = NewDictionary(ignoreCase)

    If IsObject(source) Then
        data = source.value
    Else
        data = source
    End If

    If IsArray(data) Then
        For Each item In data
            If Not IsBlankValue(item) Then
                If Not seen.Exists(item) Then seen.Add item, True
            End If
        Next item
    ElseIf Not IsBlankValue(data) Then
        seen.Add data, True
    End If

    If seen.Count = 0 Then UniqueValues = Array(): Exit Function

    ReDim out(1 To seen.Count)
    For Each item In seen.keys
        i = i + 1
        out(i) = item
    Next item
    UniqueValues = out
End Function

'------------------------------------------------------------------------------
' DuplicateKeys
' Values of a column that appear more than once - the fastest way to answer
' "are these references unique?".
'------------------------------------------------------------------------------
Public Function DuplicateKeys(ByVal source As Variant, ByVal keyColumn As Long, _
                              Optional ByVal skipHeaderRow As Boolean = True) As Variant
    Dim counts As Object
    Dim out() As Variant
    Dim key As Variant
    Dim n As Long

    Set counts = CountBy(source, keyColumn, skipHeaderRow)
    If counts.Count = 0 Then DuplicateKeys = Array(): Exit Function

    ReDim out(1 To counts.Count)
    For Each key In counts.keys
        If counts(key) > 1 Then
            n = n + 1
            out(n) = key
        End If
    Next key

    If n = 0 Then DuplicateKeys = Array(): Exit Function
    ReDim Preserve out(1 To n)
    DuplicateKeys = out
End Function

'==============================================================================
' Private helpers
'==============================================================================
' Accepts a Range or an array and always hands back a 2D array.
Private Function AsArray(ByVal source As Variant) As Variant
    Dim out As Variant

    If IsObject(source) Then
        If source Is Nothing Then Exit Function
        If source.Cells.CountLarge = 1 Then
            ReDim out(1 To 1, 1 To 1)
            out(1, 1) = source.value
            AsArray = out
        Else
            AsArray = source.value
        End If
        Exit Function
    End If

    If Not IsArray(source) Then Exit Function
    If Dimensions(source) <> 2 Then Exit Function
    AsArray = source
End Function

' Canonical string form of a lookup key. Real numbers and dates collapse to a
' single form so that 1, 1.0 and the text "1" all match, while text such as
' "007" is left alone.
Private Function KeyOf(ByVal value As Variant) As String
    Dim inner As Variant
    Dim text As String

    If IsObject(value) Then
        On Error Resume Next
        inner = value.value
        On Error GoTo 0
        KeyOf = KeyOf(inner)
        Exit Function
    End If

    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    If IsArray(value) Then Exit Function

    Select Case VarType(value)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            KeyOf = CStr(CDbl(value))
        Case vbDate
            KeyOf = CStr(CDbl(CDate(value)))
        Case vbBoolean
            KeyOf = CStr(value)
        Case Else
            text = Trim$(CStr(value))
            If IsNumeric(text) Then
                If Left$(text, 1) <> "0" Or text = "0" Or Left$(text, 2) = "0." Then
                    text = CStr(CDbl(text))
                End If
            End If
            KeyOf = text
    End Select
End Function

Private Function IsBlankValue(ByVal value As Variant) As Boolean
    If IsObject(value) Then IsBlankValue = (value Is Nothing): Exit Function
    If IsError(value) Then IsBlankValue = True: Exit Function
    If IsNull(value) Or IsEmpty(value) Then IsBlankValue = True: Exit Function
    If VarType(value) = vbString Then IsBlankValue = (Len(Trim$(value)) = 0)
End Function

Private Function IsNumericValue(ByVal value As Variant) As Boolean
    If IsObject(value) Or IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    IsNumericValue = IsNumeric(value)
End Function

Private Sub SortKeys(ByRef arr As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant, temp As Variant

    If lo >= hi Then Exit Sub
    i = lo: j = hi
    pivot = arr((lo + hi) \ 2)
    Do While i <= j
        Do While CompareKeys(arr(i), pivot) < 0
            i = i + 1
        Loop
        Do While CompareKeys(arr(j), pivot) > 0
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then SortKeys arr, lo, j
    If i < hi Then SortKeys arr, i, hi
End Sub

Private Function CompareKeys(ByVal a As Variant, ByVal b As Variant) As Long
    If IsNumericValue(a) And IsNumericValue(b) Then
        If CDbl(a) < CDbl(b) Then
            CompareKeys = -1
        ElseIf CDbl(a) > CDbl(b) Then
            CompareKeys = 1
        End If
        Exit Function
    End If
    CompareKeys = StrComp(CStr(a), CStr(b), vbTextCompare)
End Function

Private Function Reversed(ByVal arr As Variant) As Variant
    Dim out As Variant
    Dim i As Long, lo As Long, hi As Long

    lo = LBound(arr): hi = UBound(arr)
    ReDim out(lo To hi)
    For i = lo To hi
        out(i) = arr(hi - (i - lo))
    Next i
    Reversed = out
End Function

Private Function Dimensions(ByVal arr As Variant) As Long
    Dim d As Long
    Dim test As Long

    If Not IsArray(arr) Then Exit Function
    On Error GoTo Done
    Do
        d = d + 1
        test = LBound(arr, d)
    Loop While d < 60
Done:
    Dimensions = d - 1
End Function
