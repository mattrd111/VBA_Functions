Attribute VB_Name = "modExample"
Option Explicit

' Seed module so the harness has something to lint and test on day one.
' Replace with real functions; keep the shape (typed args, explicit errors).

Public Function SumTwo(ByVal a As Double, ByVal b As Double) As Double
    SumTwo = a + b
End Function

Public Function SumRange(ByVal target As Range) As Double
    ' Sums the numeric cells of a range. Raises 91 if target is Nothing, 13 if
    ' the range holds an error value - callers get a real failure, not a zero.
    Dim values As Variant
    Dim total As Double
    Dim r As Long, c As Long

    If target Is Nothing Then Err.Raise 91, "SumRange", "target range is not set"

    values = target.Value2
    If Not IsArray(values) Then
        If IsNumeric(values) Then total = CDbl(values)
    Else
        For r = LBound(values, 1) To UBound(values, 1)
            For c = LBound(values, 2) To UBound(values, 2)
                If IsNumeric(values(r, c)) Then total = total + CDbl(values(r, c))
            Next c
        Next r
    End If

    SumRange = total
End Function
