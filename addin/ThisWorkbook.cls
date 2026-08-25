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
' ThisWorkbook - Workbook Doctor add-in
'------------------------------------------------------------------------------
' Puts the menu up when the add-in loads and takes it down again when it goes,
' so Excel is never left with a menu pointing at code that is not loaded.
'
' The build script copies everything from Option Explicit down into the
' add-in's own ThisWorkbook module - document modules cannot be imported as
' new components.
'==============================================================================
Option Explicit

Private Sub Workbook_Open()
    BuildDoctorMenu
End Sub

Private Sub Workbook_AddinInstall()
    BuildDoctorMenu
End Sub

Private Sub Workbook_AddinUninstall()
    RemoveDoctorMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RemoveDoctorMenu
End Sub
