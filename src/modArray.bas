Attribute VB_Name = "modArray"
'==============================================================================
' modArray - array helpers
'------------------------------------------------------------------------------
' Testing, searching, de-duplicating, sorting and reshaping VBA arrays,
' including the 2D arrays you get back from Range.Value.
'
' Standalone: no dependency on the other modules in this repository.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' IsArrayAllocated
' True only when the variable is an array that actually has elements. An
' unallocated dynamic array raises an error on LBound, which this absorbs.
'------------------------------------------------------------------------------
Public Function IsArrayAllocated(ByVal arr As Variant) As Boolean
    On Error Resume Next
    If Not IsArray(arr) Then Exit Function
    IsArrayAllocated = (UBound(arr, 1) >= LBound(arr, 1))
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' ArrayDimensions
' Number of dimensions of an array (0 when unallocated or not an array).
'------------------------------------------------------------------------------
Public Function ArrayDimensions(ByVal arr As Variant) As Long
    Dim d As Long
    Dim test As Long

    If Not IsArray(arr) Then Exit Function
    On Error GoTo Done
    Do
        d = d + 1
        test = LBound(arr, d)
    Loop While d < 60
Done:
    ArrayDimensions = d - 1
End Function

'------------------------------------------------------------------------------
' ArrayLength
' Element count along a dimension (0 when unallocated).
'------------------------------------------------------------------------------
Public Function ArrayLength(ByVal arr As Variant, Optional ByVal dimension As Long = 1) As Long
    On Error Resume Next
    ArrayLength = UBound(arr, dimension) - LBound(arr, dimension) + 1
    If Err.Number <> 0 Then ArrayLength = 0
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' ArrayIndexOf / ArrayContains
' Searches a 1D array. ArrayIndexOf returns LBound - 1 when not found.
'------------------------------------------------------------------------------
Public Function ArrayIndexOf(ByVal arr As Variant, ByVal value As Variant, _
                             Optional ByVal caseSensitive As Boolean = False) As Long
    Dim i As Long

    If Not IsArrayAllocated(arr) Then ArrayIndexOf = -1: Exit Function
    ArrayIndexOf = LBound(arr) - 1
    For i = LBound(arr) To UBound(arr)
        If ValuesMatch(arr(i), value, caseSensitive) Then
            ArrayIndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function ArrayContains(ByVal arr As Variant, ByVal value As Variant, _
                              Optional ByVal caseSensitive As Boolean = False) As Boolean
    If Not IsArrayAllocated(arr) Then Exit Function
    ArrayContains = (ArrayIndexOf(arr, value, caseSensitive) >= LBound(arr))
End Function

'------------------------------------------------------------------------------
' ArrayUnique
' Distinct values of a 1D or 2D array, returned as a 1-based 1D array in first
' seen order. Blanks are skipped unless keepBlanks is True.
'------------------------------------------------------------------------------
Public Function ArrayUnique(ByVal arr As Variant, _
                            Optional ByVal caseSensitive As Boolean = False, _
                            Optional ByVal keepBlanks As Boolean = False) As Variant
    Dim seen As Object
    Dim item As Variant
    Dim key As Variant
    Dim out() As Variant
    Dim i As Long

    If Not IsArrayAllocated(arr) Then ArrayUnique = Array(): Exit Function

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = IIf(caseSensitive, 0, 1)

    For Each item In arr
        If keepBlanks Or Not IsEmptyValue(item) Then
            key = NormalizeKey(item)
            If Not seen.Exists(key) Then seen.Add key, item
        End If
    Next item

    If seen.count = 0 Then ArrayUnique = Array(): Exit Function

    ReDim out(1 To seen.count)
    For Each key In seen.Keys
        i = i + 1
        out(i) = seen(key)
    Next key
    ArrayUnique = out
End Function

'------------------------------------------------------------------------------
' ArrayReverse
' Returns a reversed copy of a 1D array.
'------------------------------------------------------------------------------
Public Function ArrayReverse(ByVal arr As Variant) As Variant
    Dim out As Variant
    Dim i As Long, lo As Long, hi As Long

    If Not IsArrayAllocated(arr) Then ArrayReverse = Array(): Exit Function
    lo = LBound(arr): hi = UBound(arr)
    ReDim out(lo To hi)
    For i = lo To hi
        out(i) = arr(hi - (i - lo))
    Next i
    ArrayReverse = out
End Function

'------------------------------------------------------------------------------
' SortArray
' Returns a sorted copy of a 1D array. Numbers sort numerically, everything
' else alphabetically (case insensitive), numbers before text.
'------------------------------------------------------------------------------
Public Function SortArray(ByVal arr As Variant, Optional ByVal descending As Boolean = False) As Variant
    Dim out As Variant

    If Not IsArrayAllocated(arr) Then SortArray = Array(): Exit Function
    out = arr
    QuickSort out, LBound(out), UBound(out)
    If descending Then out = ArrayReverse(out)
    SortArray = out
End Function

'------------------------------------------------------------------------------
' QuickSort
' Sorts a 1D array in place. Call without the index arguments to sort it all.
'------------------------------------------------------------------------------
Public Sub QuickSort(ByRef arr As Variant, Optional ByVal lowIndex As Variant, _
                     Optional ByVal highIndex As Variant)
    Dim i As Long, j As Long, lo As Long, hi As Long
    Dim pivot As Variant, temp As Variant

    If Not IsArrayAllocated(arr) Then Exit Sub
    If IsMissing(lowIndex) Then lowIndex = LBound(arr)
    If IsMissing(highIndex) Then highIndex = UBound(arr)

    lo = CLng(lowIndex): hi = CLng(highIndex)
    If lo >= hi Then Exit Sub

    i = lo: j = hi
    pivot = arr((lo + hi) \ 2)
    Do While i <= j
        Do While CompareValues(arr(i), pivot) < 0
            i = i + 1
        Loop
        Do While CompareValues(arr(j), pivot) > 0
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If lo < j Then QuickSort arr, lo, j
    If i < hi Then QuickSort arr, i, hi
End Sub

'------------------------------------------------------------------------------
' Sort2D
' Returns a copy of a 2D array with its rows sorted by one column - exactly
' what you want after reading a block of cells with Range.Value.
'
'   data = ws.Range("A2:D500").Value
'   data = Sort2D(data, 3, True)      ' by column 3, largest first
'------------------------------------------------------------------------------
Public Function Sort2D(ByVal arr As Variant, ByVal keyColumn As Long, _
                       Optional ByVal descending As Boolean = False) As Variant
    Dim keys() As Variant
    Dim idx() As Long
    Dim out As Variant
    Dim rowLo As Long, rowHi As Long, colLo As Long, colHi As Long
    Dim i As Long, c As Long, n As Long, source As Long

    If Not IsArrayAllocated(arr) Then Exit Function
    If ArrayDimensions(arr) <> 2 Then Exit Function

    rowLo = LBound(arr, 1): rowHi = UBound(arr, 1)
    colLo = LBound(arr, 2): colHi = UBound(arr, 2)
    If keyColumn < colLo Or keyColumn > colHi Then Exit Function

    n = rowHi - rowLo + 1
    ReDim keys(0 To n - 1)
    ReDim idx(0 To n - 1)
    For i = 0 To n - 1
        keys(i) = arr(rowLo + i, keyColumn)
        idx(i) = rowLo + i
    Next i

    SortPairs keys, idx, 0, n - 1

    ReDim out(rowLo To rowHi, colLo To colHi)
    For i = 0 To n - 1
        If descending Then source = idx(n - 1 - i) Else source = idx(i)
        For c = colLo To colHi
            out(rowLo + i, c) = arr(source, c)
        Next c
    Next i

    Sort2D = out
End Function

'------------------------------------------------------------------------------
' GetColumn / GetRow
' Pulls one column or row out of a 2D array as a 1-based 1D array.
'------------------------------------------------------------------------------
Public Function GetColumn(ByVal arr As Variant, ByVal columnIndex As Long) As Variant
    Dim out() As Variant
    Dim i As Long, n As Long

    If ArrayDimensions(arr) <> 2 Then GetColumn = Array(): Exit Function
    If columnIndex < LBound(arr, 2) Or columnIndex > UBound(arr, 2) Then GetColumn = Array(): Exit Function

    n = UBound(arr, 1) - LBound(arr, 1) + 1
    ReDim out(1 To n)
    For i = 1 To n
        out(i) = arr(LBound(arr, 1) + i - 1, columnIndex)
    Next i
    GetColumn = out
End Function

Public Function GetRow(ByVal arr As Variant, ByVal rowIndex As Long) As Variant
    Dim out() As Variant
    Dim i As Long, n As Long

    If ArrayDimensions(arr) <> 2 Then GetRow = Array(): Exit Function
    If rowIndex < LBound(arr, 1) Or rowIndex > UBound(arr, 1) Then GetRow = Array(): Exit Function

    n = UBound(arr, 2) - LBound(arr, 2) + 1
    ReDim out(1 To n)
    For i = 1 To n
        out(i) = arr(rowIndex, LBound(arr, 2) + i - 1)
    Next i
    GetRow = out
End Function

'------------------------------------------------------------------------------
' TransposeArray
' Transposes a 2D array without Application.Transpose, so it is not limited to
' 65,536 rows and does not trip over long strings.
'------------------------------------------------------------------------------
Public Function TransposeArray(ByVal arr As Variant) As Variant
    Dim out As Variant
    Dim r As Long, c As Long

    If ArrayDimensions(arr) <> 2 Then Exit Function
    ReDim out(LBound(arr, 2) To UBound(arr, 2), LBound(arr, 1) To UBound(arr, 1))
    For r = LBound(arr, 1) To UBound(arr, 1)
        For c = LBound(arr, 2) To UBound(arr, 2)
            out(c, r) = arr(r, c)
        Next c
    Next r
    TransposeArray = out
End Function

'------------------------------------------------------------------------------
' To2D
' Turns a 1D array into a 1-based 2D array so it can be written straight to a
' range. asRow = False gives a column (n x 1), True gives a row (1 x n).
'------------------------------------------------------------------------------
Public Function To2D(ByVal arr As Variant, Optional ByVal asRow As Boolean = False) As Variant
    Dim out As Variant
    Dim i As Long, n As Long

    If Not IsArrayAllocated(arr) Then Exit Function
    If ArrayDimensions(arr) = 2 Then To2D = arr: Exit Function

    n = UBound(arr) - LBound(arr) + 1
    If asRow Then
        ReDim out(1 To 1, 1 To n)
    Else
        ReDim out(1 To n, 1 To 1)
    End If

    For i = 1 To n
        If asRow Then
            out(1, i) = arr(LBound(arr) + i - 1)
        Else
            out(i, 1) = arr(LBound(arr) + i - 1)
        End If
    Next i
    To2D = out
End Function

'------------------------------------------------------------------------------
' RemoveBlanks
' Copy of a 1D array with Empty, Null, error and whitespace-only values dropped.
'------------------------------------------------------------------------------
Public Function RemoveBlanks(ByVal arr As Variant) As Variant
    Dim out() As Variant
    Dim item As Variant
    Dim n As Long

    If Not IsArrayAllocated(arr) Then RemoveBlanks = Array(): Exit Function
    ReDim out(1 To TotalElements(arr))
    For Each item In arr
        If Not IsEmptyValue(item) Then
            n = n + 1
            out(n) = item
        End If
    Next item

    If n = 0 Then RemoveBlanks = Array(): Exit Function
    ReDim Preserve out(1 To n)
    RemoveBlanks = out
End Function

'------------------------------------------------------------------------------
' ConcatArrays
' Appends two 1D arrays into one 1-based array.
'------------------------------------------------------------------------------
Public Function ConcatArrays(ByVal first As Variant, ByVal second As Variant) As Variant
    Dim out() As Variant
    Dim item As Variant
    Dim n As Long

    n = TotalElements(first) + TotalElements(second)
    If n = 0 Then ConcatArrays = Array(): Exit Function

    ReDim out(1 To n)
    n = 0
    If IsArrayAllocated(first) Then
        For Each item In first
            n = n + 1
            out(n) = item
        Next item
    End If
    If IsArrayAllocated(second) Then
        For Each item In second
            n = n + 1
            out(n) = item
        Next item
    End If
    ConcatArrays = out
End Function

'------------------------------------------------------------------------------
' ArrayToString
' Readable dump of a 1D or 2D array - handy in the Immediate window.
'------------------------------------------------------------------------------
Public Function ArrayToString(ByVal arr As Variant, Optional ByVal delimiter As String = ", ") As String
    Dim r As Long, c As Long
    Dim lines As String
    Dim rowText As String

    If Not IsArrayAllocated(arr) Then ArrayToString = "<empty>": Exit Function

    If ArrayDimensions(arr) = 1 Then
        For r = LBound(arr) To UBound(arr)
            lines = lines & IIf(r = LBound(arr), "", delimiter) & SafeString(arr(r))
        Next r
    Else
        For r = LBound(arr, 1) To UBound(arr, 1)
            rowText = ""
            For c = LBound(arr, 2) To UBound(arr, 2)
                rowText = rowText & IIf(c = LBound(arr, 2), "", delimiter) & SafeString(arr(r, c))
            Next c
            lines = lines & IIf(r = LBound(arr, 1), "", vbNewLine) & rowText
        Next r
    End If
    ArrayToString = lines
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Sub SortPairs(ByRef keys As Variant, ByRef idx() As Long, _
                      ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant, tempKey As Variant
    Dim tempIdx As Long

    If lo >= hi Then Exit Sub
    i = lo: j = hi
    pivot = keys((lo + hi) \ 2)
    Do While i <= j
        Do While CompareValues(keys(i), pivot) < 0
            i = i + 1
        Loop
        Do While CompareValues(keys(j), pivot) > 0
            j = j - 1
        Loop
        If i <= j Then
            tempKey = keys(i): keys(i) = keys(j): keys(j) = tempKey
            tempIdx = idx(i): idx(i) = idx(j): idx(j) = tempIdx
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then SortPairs keys, idx, lo, j
    If i < hi Then SortPairs keys, idx, i, hi
End Sub

' -1 when a sorts before b, 0 when equal, 1 when a sorts after b.
' Blanks last, then numbers, then text.
Private Function CompareValues(ByVal a As Variant, ByVal b As Variant) As Long
    Dim aBlank As Boolean, bBlank As Boolean
    Dim aNum As Boolean, bNum As Boolean

    aBlank = IsEmptyValue(a)
    bBlank = IsEmptyValue(b)
    If aBlank And bBlank Then Exit Function
    If aBlank Then CompareValues = 1: Exit Function
    If bBlank Then CompareValues = -1: Exit Function

    aNum = IsNumericValue(a)
    bNum = IsNumericValue(b)
    If aNum And bNum Then
        If CDbl(a) < CDbl(b) Then
            CompareValues = -1
        ElseIf CDbl(a) > CDbl(b) Then
            CompareValues = 1
        End If
        Exit Function
    End If
    If aNum Then CompareValues = -1: Exit Function
    If bNum Then CompareValues = 1: Exit Function

    CompareValues = StrComp(SafeString(a), SafeString(b), vbTextCompare)
End Function

Private Function ValuesMatch(ByVal a As Variant, ByVal b As Variant, _
                             ByVal caseSensitive As Boolean) As Boolean
    If IsEmptyValue(a) Or IsEmptyValue(b) Then
        ValuesMatch = (IsEmptyValue(a) And IsEmptyValue(b))
        Exit Function
    End If
    If IsNumericValue(a) And IsNumericValue(b) Then
        ValuesMatch = (CDbl(a) = CDbl(b))
    Else
        ValuesMatch = (StrComp(SafeString(a), SafeString(b), _
                       IIf(caseSensitive, vbBinaryCompare, vbTextCompare)) = 0)
    End If
End Function

' Total element count across every dimension.
Private Function TotalElements(ByVal arr As Variant) As Long
    Dim d As Long, dims As Long, total As Long

    dims = ArrayDimensions(arr)
    If dims = 0 Then Exit Function
    total = 1
    For d = 1 To dims
        total = total * ArrayLength(arr, d)
    Next d
    TotalElements = total
End Function

Private Function IsEmptyValue(ByVal value As Variant) As Boolean
    If IsObject(value) Then IsEmptyValue = (value Is Nothing): Exit Function
    If IsError(value) Then IsEmptyValue = True: Exit Function
    If IsNull(value) Or IsEmpty(value) Then IsEmptyValue = True: Exit Function
    If VarType(value) = vbString Then IsEmptyValue = (Len(Trim$(value)) = 0)
End Function

Private Function IsNumericValue(ByVal value As Variant) As Boolean
    If IsObject(value) Or IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    If VarType(value) = vbString Then Exit Function
    IsNumericValue = IsNumeric(value)
End Function

Private Function SafeString(ByVal value As Variant) As String
    If IsObject(value) Then SafeString = TypeName(value): Exit Function
    If IsError(value) Then SafeString = "#ERROR": Exit Function
    If IsNull(value) Then SafeString = "": Exit Function
    If IsEmpty(value) Then SafeString = "": Exit Function
    If IsArray(value) Then SafeString = "<array>": Exit Function
    SafeString = CStr(value)
End Function

' Dictionary keys must be scalars, and 1 must not collide with "1".
Private Function NormalizeKey(ByVal value As Variant) As Variant
    If IsObject(value) Then NormalizeKey = "obj:" & TypeName(value): Exit Function
    If IsError(value) Then NormalizeKey = "e:#error": Exit Function
    If IsNumericValue(value) Then
        NormalizeKey = "n:" & CStr(CDbl(value))
    Else
        NormalizeKey = "s:" & SafeString(value)
    End If
End Function
