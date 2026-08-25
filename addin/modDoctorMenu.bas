Attribute VB_Name = "modDoctorMenu"
'==============================================================================
' modDoctorMenu - the menu the add-in puts on the ribbon
'------------------------------------------------------------------------------
' Built with CommandBars rather than ribbon XML, so the whole add-in is plain
' importable VBA with nothing to unzip or hand-edit. Excel shows it as
' "Workbook Doctor" on the Add-ins tab.
'
' Depends on: every other Doctor module
'==============================================================================
Option Explicit

Private Const MENU_TAG As String = "WorkbookDoctorMenu"
Private Const MENU_CAPTION As String = "Workbook &Doctor"

'------------------------------------------------------------------------------
' BuildDoctorMenu
' Called when the add-in loads. Safe to call twice - it clears itself first.
'------------------------------------------------------------------------------
Public Sub BuildDoctorMenu()
    Dim root As CommandBarPopup

    RemoveDoctorMenu

    On Error Resume Next
    Set root = Application.CommandBars("Worksheet Menu Bar").Controls.Add( _
                   Type:=msoControlPopup, Temporary:=True)
    If root Is Nothing Then Exit Sub

    root.Caption = MENU_CAPTION
    root.Tag = MENU_TAG

    AddButton root, "&Audit workbook (size and bloat)...", "AuditWorkbook", 23, False
    AddButton root, "Audit &model (formulas)...", "AuditModel", 466, False
    AddButton root, "Formula ma&p (this sheet)", "ShowFormulaMap", 1000, False
    AddButton root, "Select fla&gged cells (this sheet)", "SelectFlaggedCells", 350, False

    AddButton root, "Clean &everything (safe)...", "CleanEverything", 358, True

    AddButton root, "Clean &names...", "CleanNames", 472, True
    AddButton root, "Clean &styles...", "CleanStyles", 1018, False
    AddButton root, "&Reset used range...", "ResetUsedRange", 358, False
    AddButton root, "Remove in&visible objects...", "RemoveInvisibleShapes", 493, False
    AddButton root, "Delete e&mpty sheets...", "DeleteEmptySheets", 292, False
    AddButton root, "Clean conditional &formatting...", "CleanConditionalFormats", 435, False

    AddButton root, "&List defined names", "ShowNamesReport", 353, True
    AddButton root, "List cell st&yles", "ShowStylesReport", 353, False
    AddButton root, "External lin&ks...", "ShowExternalLinks", 1663, False
    AddButton root, "&Break external links...", "BreakExternalLinks", 1664, False
    AddButton root, "&Unhide all sheets", "UnhideAllSheets", 2110, False

    AddButton root, "&Trim and clean selection", "TrimAndCleanSelection", 384, True
    AddButton root, "Selection to &values", "SelectionToValues", 370, False
    AddButton root, "Remove &hyperlinks from selection", "RemoveHyperlinksFromSelection", 1691, False

    AddButton root, "Bac&kup this workbook", "BackupActiveWorkbook", 3, True
    AddButton root, "A&bout Workbook Doctor", "AboutDoctor", 487, True

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' RemoveDoctorMenu
' Called when the add-in unloads, so nothing is left behind pointing at code
' that is no longer there.
'------------------------------------------------------------------------------
Public Sub RemoveDoctorMenu()
    Dim control As CommandBarControl
    Dim guard As Long

    On Error Resume Next
    Do
        guard = guard + 1
        Set control = Nothing
        Set control = Application.CommandBars("Worksheet Menu Bar").FindControl( _
                          Tag:=MENU_TAG, Recursive:=True)
        If control Is Nothing Then Exit Do
        control.Delete
    Loop While guard < 50
    On Error GoTo 0
End Sub

'==============================================================================
' Private helpers
'==============================================================================
Private Sub AddButton(ByVal parent As CommandBarPopup, ByVal caption As String, _
                      ByVal macroName As String, ByVal iconId As Long, _
                      ByVal startsGroup As Boolean)
    Dim button As CommandBarButton

    On Error Resume Next
    Set button = parent.Controls.Add(Type:=msoControlButton, Temporary:=True)
    If button Is Nothing Then Exit Sub

    button.Caption = caption
    button.OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
    button.Tag = MENU_TAG
    button.BeginGroup = startsGroup
    button.Style = msoButtonIconAndCaption
    If iconId > 0 Then button.FaceId = iconId
    On Error GoTo 0
End Sub
