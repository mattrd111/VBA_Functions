Attribute VB_Name = "modFX"
Option Explicit

' =====================================================================================
' modFX - refresh the FX Input Power Query connection from VBA.
'
' Companion to flatfile/FX_Input.pq. Once the query is loaded to the "FX Input" sheet
' (see FX-PowerQuery-Setup.md), this replaces the manual copy-and-paste of the ECB CSV.
' Wire RefreshFXInput to a button on the FX tab if you want one.
' =====================================================================================

Private Const FX_QUERY_NAME As String = "FX_Input"
Private Const FX_INPUT_SHEET As String = "FX Input"

Public Sub RefreshFXInput()
    Dim conn As WorkbookConnection
    Set conn = FindFXConnection()
    If conn Is Nothing Then
        MsgBox "No Power Query connection called '" & FX_QUERY_NAME & "' was found in this workbook." & vbCrLf & _
               "Set it up first - see FX-PowerQuery-Setup.md.", vbExclamation, "FX Input"
        Exit Sub
    End If

    Dim topBefore As Variant, topAfter As Variant
    topBefore = TopDate()

    On Error GoTo RefreshFailed
    ' Synchronous refresh so the calculation below sees the new rows.
    If conn.Type = xlConnectionTypeOLEDB Then conn.OLEDBConnection.BackgroundQuery = False
    conn.Refresh
    On Error GoTo 0

    Application.Calculate
    topAfter = TopDate()

    MsgBox "FX Input refreshed from the ECB." & vbCrLf & vbCrLf & _
           "Newest date before: " & FormatDate(topBefore) & vbCrLf & _
           "Newest date now:    " & FormatDate(topAfter) & vbCrLf & vbCrLf & _
           "Remember to roll the FX tab's row-1 dates forward if the month-end has moved.", _
           vbInformation, "FX Input"
    Exit Sub

RefreshFailed:
    MsgBox "The refresh failed: " & Err.Description & vbCrLf & vbCrLf & _
           "If this is a proxy or network error, download eurofxref-hist.csv from the ECB page " & _
           "and point the query's Raw step at the file (see FX-PowerQuery-Setup.md).", _
           vbCritical, "FX Input"
End Sub

' Power Query connections are named "Query - <query name>"; fall back to a name match.
Private Function FindFXConnection() As WorkbookConnection
    Dim conn As WorkbookConnection
    On Error Resume Next
    Set FindFXConnection = ThisWorkbook.Connections("Query - " & FX_QUERY_NAME)
    On Error GoTo 0
    If Not FindFXConnection Is Nothing Then Exit Function
    For Each conn In ThisWorkbook.Connections
        If InStr(1, conn.Name, FX_QUERY_NAME, vbTextCompare) > 0 Then
            Set FindFXConnection = conn
            Exit Function
        End If
    Next conn
End Function

Private Function TopDate() As Variant
    On Error Resume Next
    TopDate = ThisWorkbook.Sheets(FX_INPUT_SHEET).Range("A2").Value
    On Error GoTo 0
End Function

Private Function FormatDate(ByVal v As Variant) As String
    If IsDate(v) Then
        FormatDate = Format(CDate(v), "dd mmm yyyy")
    Else
        FormatDate = "(none)"
    End If
End Function
