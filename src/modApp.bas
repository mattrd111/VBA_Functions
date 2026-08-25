Attribute VB_Name = "modApp"
'==============================================================================
' modApp - application state, progress and timing
'------------------------------------------------------------------------------
' The wrapper every long-running macro needs: turn Excel's housekeeping off
' while you work, put a progress bar in the status bar, time the run, and put
' everything back afterwards - including when the code fails halfway through.
'
' Standalone: no dependency on the other modules in this repository.
'==============================================================================
Option Explicit

Private mNestingDepth As Long
Private mStateSaved As Boolean
Private mCalculation As XlCalculation
Private mScreenUpdating As Boolean
Private mEnableEvents As Boolean
Private mDisplayAlerts As Boolean
Private mDisplayStatusBar As Boolean
Private mTimerStart As Double
Private mTimerRunning As Boolean

'------------------------------------------------------------------------------
' FastMode
' Switches off screen updating, events, alerts and automatic calculation, and
' remembers how it found them. Nested calls are counted, so an inner routine
' cannot turn the screen back on halfway through an outer one.
'
'   FastMode True
'   On Error GoTo Cleanup
'   ... work ...
' Cleanup:
'   FastMode False
'------------------------------------------------------------------------------
Public Sub FastMode(Optional ByVal enable As Boolean = True)
    If enable Then
        If mNestingDepth = 0 Then SaveState
        mNestingDepth = mNestingDepth + 1

        On Error Resume Next
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.Cursor = xlWait
        On Error GoTo 0
    Else
        If mNestingDepth > 0 Then mNestingDepth = mNestingDepth - 1
        If mNestingDepth = 0 Then RestoreState
    End If
End Sub

'------------------------------------------------------------------------------
' ResetApp
' Panic button: puts Excel back to normal whatever the nesting count says.
' Worth wiring to a button - a macro that fails inside FastMode leaves the
' screen frozen until something calls this.
'------------------------------------------------------------------------------
Public Sub ResetApp()
    mNestingDepth = 0
    RestoreState
End Sub

'------------------------------------------------------------------------------
' ShowProgress
' Draws a progress bar in the status bar. Call it every few hundred rows rather
' than every row - the DoEvents that keeps Excel responsive is not free.
'
'   If r Mod 500 = 0 Then ShowProgress r, lastRow, "Checking invoices"
'------------------------------------------------------------------------------
Public Sub ShowProgress(ByVal current As Double, ByVal total As Double, _
                        Optional ByVal message As String = "Working", _
                        Optional ByVal barWidth As Long = 20)
    Dim fraction As Double
    Dim filled As Long

    If total <= 0 Then Exit Sub
    If barWidth < 1 Then barWidth = 20

    fraction = current / total
    If fraction < 0 Then fraction = 0
    If fraction > 1 Then fraction = 1
    filled = Int(fraction * barWidth)

    On Error Resume Next
    Application.DisplayStatusBar = True
    Application.StatusBar = message & "   [" & String$(filled, ChrW(9608)) & _
                            String$(barWidth - filled, ChrW(9617)) & "] " & _
                            Format$(fraction, "0%") & "   " & _
                            Format$(current, "#,##0") & " of " & Format$(total, "#,##0")
    On Error GoTo 0
    DoEvents
End Sub

'------------------------------------------------------------------------------
' ClearStatusBar
' Hands the status bar back to Excel.
'------------------------------------------------------------------------------
Public Sub ClearStatusBar()
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' StartTimer / ElapsedSeconds / ElapsedText
' Simple stopwatch for "why is this slow" questions. Survives midnight.
'
'   StartTimer
'   ... work ...
'   Debug.Print "Took " & ElapsedText
'------------------------------------------------------------------------------
Public Sub StartTimer()
    mTimerStart = Timer
    mTimerRunning = True
End Sub

Public Function ElapsedSeconds() As Double
    Dim elapsed As Double

    If Not mTimerRunning Then Exit Function
    elapsed = Timer - mTimerStart
    If elapsed < 0 Then elapsed = elapsed + 86400#      ' the clock passed midnight
    ElapsedSeconds = elapsed
End Function

Public Function ElapsedText() As String
    Dim secs As Double

    secs = ElapsedSeconds
    If secs < 60 Then
        ElapsedText = Format$(secs, "0.00") & "s"
    ElseIf secs < 3600 Then
        ElapsedText = Int(secs / 60) & "m " & Format$(secs - Int(secs / 60) * 60, "0.0") & "s"
    Else
        ElapsedText = Int(secs / 3600) & "h " & Int((secs - Int(secs / 3600) * 3600) / 60) & "m"
    End If
End Function

'------------------------------------------------------------------------------
' DebugLog
' Timestamped line in the Immediate window - Ctrl+G to read it.
'------------------------------------------------------------------------------
Public Sub DebugLog(ParamArray parts() As Variant)
    Dim i As Long
    Dim output As String

    For i = LBound(parts) To UBound(parts)
        If i > LBound(parts) Then output = output & " "
        output = output & AsText(parts(i))
    Next i
    Debug.Print Format$(Now, "hh:nn:ss") & "  " & output
End Sub

'------------------------------------------------------------------------------
' LogToFile
' Appends a timestamped line to a text file, creating it if needed. Useful when
' a macro runs unattended and the Immediate window is not there to read.
'------------------------------------------------------------------------------
Public Sub LogToFile(ByVal filePath As String, ByVal message As String)
    Dim handle As Integer

    If Len(filePath) = 0 Then Exit Sub

    On Error Resume Next
    handle = FreeFile
    Open filePath For Append As #handle
    Print #handle, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & message
    Close #handle
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' ErrorInfo
' One line describing the current error, for logs and message boxes.
'
'   Cleanup:
'     FastMode False
'     If Err.Number <> 0 Then MsgBox ErrorInfo("ImportInvoices")
'------------------------------------------------------------------------------
Public Function ErrorInfo(Optional ByVal procedureName As String = "") As String
    Dim text As String

    If Err.Number = 0 Then Exit Function

    text = "Error " & Err.Number & " in " & _
           IIf(Len(procedureName) > 0, procedureName, "(unknown)") & ": " & Err.Description
    If Len(Err.source) > 0 Then text = text & " [" & Err.source & "]"

    ErrorInfo = text
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Sub SaveState()
    On Error Resume Next
    mScreenUpdating = Application.ScreenUpdating
    mEnableEvents = Application.EnableEvents
    mDisplayAlerts = Application.DisplayAlerts
    mDisplayStatusBar = Application.DisplayStatusBar
    mCalculation = Application.Calculation      ' raises with no workbook open
    If Err.Number <> 0 Then
        mCalculation = xlCalculationAutomatic
        Err.Clear
    End If
    On Error GoTo 0
    mStateSaved = True
End Sub

Private Sub RestoreState()
    On Error Resume Next
    If mStateSaved Then
        Application.ScreenUpdating = mScreenUpdating
        Application.EnableEvents = mEnableEvents
        Application.DisplayAlerts = mDisplayAlerts
        Application.DisplayStatusBar = mDisplayStatusBar
        Application.Calculation = mCalculation
    Else
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.DisplayStatusBar = True
        Application.Calculation = xlCalculationAutomatic
    End If

    Application.StatusBar = False
    Application.Cursor = xlDefault
    On Error GoTo 0

    mStateSaved = False
End Sub

Private Function AsText(ByVal value As Variant) As String
    If IsObject(value) Then AsText = "<" & TypeName(value) & ">": Exit Function
    If IsError(value) Then AsText = "#ERROR": Exit Function
    If IsNull(value) Then AsText = "Null": Exit Function
    If IsEmpty(value) Then AsText = "Empty": Exit Function
    If IsArray(value) Then AsText = "<array>": Exit Function
    AsText = CStr(value)
End Function
