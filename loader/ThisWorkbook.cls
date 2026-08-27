VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================================================
' ThisWorkbook - Workbook Doctor loader
'------------------------------------------------------------------------------
' Opening one workbook from inside another's Workbook_Open is unreliable, so
' the work is handed to OnTime and happens a moment later, once Excel has
' finished starting.
'==============================================================================
Option Explicit

Private Sub Workbook_Open()
    ScheduleStart
End Sub

Private Sub Workbook_AddinInstall()
    ScheduleStart
End Sub

Private Sub Workbook_AddinUninstall()
    RemoveLoaderMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RemoveLoaderMenu
End Sub

Private Sub ScheduleStart()
    On Error Resume Next
    Application.OnTime Now, "'" & ThisWorkbook.Name & "'!StartWorkbookDoctor"
    On Error GoTo 0
End Sub
