Attribute VB_Name = "modTestRunner"
Option Explicit

' Test registry and runner. VBA cannot enumerate procedures, so every test is
' registered by name here; RunAllTests invokes them via Application.Run and
' writes a JSON result file the CI/agent runner reads.

Private Const RESULT_FILE As String = "vba-test-results.json"

Private mPassed As Long
Private mFailed As Long
Private mLog As String

Public Sub RunAllTests()
    mPassed = 0
    mFailed = 0
    mLog = ""

    ' --- register tests here -------------------------------------------------
    ' One RunTest line per test procedure, e.g.:
    '   RunTest "Test_SumRange_RaisesOnNothing"
    ' VBA cannot enumerate procedures, so a test that is not listed here never
    ' runs - and a suite that silently skips is worse than no suite.
    ' -------------------------------------------------------------------------

    WriteResults
End Sub

Private Sub RunTest(ByVal testName As String)
    modAssert.BeginTest testName
    On Error GoTo Failed
    Application.Run testName
    If modAssert.TestFailed Then
        mFailed = mFailed + 1
        mLog = mLog & "FAIL " & testName & vbLf
    Else
        mPassed = mPassed + 1
        mLog = mLog & "PASS " & testName & vbLf
    End If
    Exit Sub
Failed:
    mFailed = mFailed + 1
    modAssert.Fail "unhandled error " & Err.Number & ": " & Err.Description
    mLog = mLog & "ERROR " & testName & " - " & Err.Description & vbLf
End Sub

Private Sub WriteResults()
    Dim path As String, fileNum As Integer, i As Long, failures As String
    Dim f As Variant

    For Each f In modAssert.Failures
        If Len(failures) > 0 Then failures = failures & ","
        failures = failures & """" & JsonEscape(CStr(f)) & """"
    Next f

    path = ThisWorkbook.path & Application.PathSeparator & RESULT_FILE
    fileNum = FreeFile
    Open path For Output As #fileNum
    Print #fileNum, "{"
    Print #fileNum, "  ""passed"": " & mPassed & ","
    Print #fileNum, "  ""failed"": " & mFailed & ","
    Print #fileNum, "  ""assertions"": " & modAssert.AssertCount & ","
    Print #fileNum, "  ""failures"": [" & failures & "],"
    Print #fileNum, "  ""log"": """ & JsonEscape(mLog) & """"
    Print #fileNum, "}"
    Close #fileNum

    Debug.Print mLog
    Debug.Print "passed=" & mPassed & " failed=" & mFailed
End Sub

Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function
