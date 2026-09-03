Attribute VB_Name = "modAssert"
Option Explicit

' Minimal assertion library for the workbook test harness.
' Results accumulate in module state so modTestRunner can write them out.

Private mFailures As Collection
Private mAsserts As Long
Private mCurrentTest As String
Private mCurrentFailed As Boolean

Public Sub BeginTest(ByVal testName As String)
    If mFailures Is Nothing Then Set mFailures = New Collection
    mCurrentTest = testName
    mCurrentFailed = False
End Sub

Public Function TestFailed() As Boolean
    TestFailed = mCurrentFailed
End Function

Public Property Get AssertCount() As Long
    AssertCount = mAsserts
End Property

Public Property Get Failures() As Collection
    If mFailures Is Nothing Then Set mFailures = New Collection
    Set Failures = mFailures
End Property

Public Sub Fail(ByVal message As String)
    If mFailures Is Nothing Then Set mFailures = New Collection
    mCurrentFailed = True
    mFailures.Add mCurrentTest & ": " & message
End Sub

Public Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    mAsserts = mAsserts + 1
    If Not condition Then Fail message & " (expected True)"
End Sub

Public Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    mAsserts = mAsserts + 1
    If CStr(expected) <> CStr(actual) Then
        Fail message & " (expected '" & CStr(expected) & "', got '" & CStr(actual) & "')"
    End If
End Sub

Public Sub AssertClose(ByVal expected As Double, ByVal actual As Double, _
                       ByVal tolerance As Double, ByVal message As String)
    mAsserts = mAsserts + 1
    If Abs(expected - actual) > tolerance Then
        Fail message & " (expected " & expected & " +/- " & tolerance & ", got " & actual & ")"
    End If
End Sub

Public Sub AssertRaises(ByVal expectedErrorNumber As Long, ByVal message As String)
    ' Call after an On Error Resume Next block: checks Err.Number then clears it.
    mAsserts = mAsserts + 1
    If Err.Number <> expectedErrorNumber Then
        Fail message & " (expected error " & expectedErrorNumber & ", got " & Err.Number & ")"
    End If
    Err.Clear
End Sub
