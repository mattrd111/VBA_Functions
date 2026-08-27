Attribute VB_Name = "Test_Example"
Option Explicit

' Example tests. Delete these once real ones exist, and keep the registry in
' modTestRunner.RunAllTests in step with what lives here.

Public Sub Test_Example_AddsNumbers()
    modAssert.AssertEqual 3, SumTwo(1, 2), "SumTwo(1,2)"
    modAssert.AssertEqual -1, SumTwo(2, -3), "SumTwo(2,-3)"
End Sub

Public Sub Test_Example_RejectsEmptyRange()
    Dim result As Variant
    On Error Resume Next
    result = SumRange(Nothing)
    modAssert.AssertRaises 91, "SumRange(Nothing) should raise"
    On Error GoTo 0
End Sub
