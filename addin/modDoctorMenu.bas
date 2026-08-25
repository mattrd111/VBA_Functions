Attribute VB_Name = "modDoctorMenu"
'==============================================================================
' modDoctorMenu - the menu the add-in puts on the ribbon
'------------------------------------------------------------------------------
' Built with CommandBars rather than ribbon XML, so the whole add-in stays plain
' importable VBA with nothing to unzip or hand-edit. Excel shows it as
' "Workbook Doctor" on the Add-ins tab.
'
' Grouped into submenus, because a flat list of thirty tools is a list nobody
' reads: Audit reports and changes nothing, Clean changes the workbook, Data
' reshapes an extract into something you can pivot, Cells acts on a selection.
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
    Dim group As CommandBarPopup

    RemoveDoctorMenu

    On Error Resume Next
    Set root = Application.CommandBars("Worksheet Menu Bar").Controls.Add( _
                   Type:=msoControlPopup, Temporary:=True)
    If root Is Nothing Then Exit Sub

    root.Caption = MENU_CAPTION
    root.Tag = MENU_TAG

    '--- Audit: reads, reports, changes nothing ------------------------------
    Set group = AddPopup(root, "&Audit  (changes nothing)", False)
    AddButton group, "Audit &workbook  (size and bloat)...", "AuditWorkbook", 23, False
    AddButton group, "Audit &model  (formulas)...", "AuditModel", 466, False
    AddButton group, "Formula ma&p  (this sheet)", "ShowFormulaMap", 1000, False
    AddButton group, "Select fla&gged cells  (this sheet)", "SelectFlaggedCells", 350, False
    AddButton group, "List defined &names", "ShowNamesReport", 353, True
    AddButton group, "List cell &styles", "ShowStylesReport", 353, False
    AddButton group, "&External links...", "ShowExternalLinks", 1663, False

    '--- Clean: changes the workbook -----------------------------------------
    Set group = AddPopup(root, "&Clean  (changes the workbook)", False)
    AddButton group, "Clean &everything  (safe)...", "CleanEverything", 358, False
    AddButton group, "Clean &names...", "CleanNames", 472, True
    AddButton group, "Clean &styles...", "CleanStyles", 1018, False
    AddButton group, "&Reset used range...", "ResetUsedRange", 358, False
    AddButton group, "Remove in&visible objects...", "RemoveInvisibleShapes", 493, False
    AddButton group, "Delete e&mpty sheets...", "DeleteEmptySheets", 292, False
    AddButton group, "Clean conditional &formatting...", "CleanConditionalFormats", 435, False
    AddButton group, "&Break external links...", "BreakExternalLinks", 1664, True

    '--- Data: turning extracts into tables ----------------------------------
    Set group = AddPopup(root, "&Data", False)
    AddButton group, "Stack selected &sheets...", "StackSelectedSheets", 292, False
    AddButton group, "Stack &files in a folder...", "StackFilesInFolder", 23, False
    AddButton group, "&Unpivot a cross-tab...", "UnpivotSelection", 431, True
    AddButton group, "Fill &blanks down...", "FillBlanksDown", 370, False
    AddButton group, "Fu&zzy match two lists...", "FuzzyMatchLists", 1728, True
    AddButton group, "F&und maths reference sheet", "ShowFundMathsReference", 1665, True

    '--- Format: the Alpha house style ---------------------------------------
    Set group = AddPopup(root, "&Format", False)
    AddButton group, "Cycle &number format", "CycleNumberFormat", 1731, False
    AddButton group, "Cycle &percent format", "CyclePercentFormat", 402, False
    AddButton group, "Cycle &date format", "CycleDateFormat", 125, False
    AddButton group, "Style as Alpha &table", "StyleAlphaTable", 1663, True
    AddButton group, "Style as &header row", "StyleAlphaHeader", 293, False
    AddButton group, "Style as t&otal row", "StyleAlphaTotal", 459, False
    AddButton group, "House type on this &sheet...", "StyleAlphaSheet", 291, False
    AddButton group, "&Colour inputs and formulas...", "ColourInputsAndFormulas", 401, True
    AddButton group, "Alpha chart c&olours", "ApplyAlphaChartColours", 435, True
    AddButton group, "Alpha p&alette reference", "ShowAlphaPalette", 1000, False

    '--- Cells: acts on the selection ----------------------------------------
    Set group = AddPopup(root, "Cel&ls", False)
    AddButton group, "&Trim and clean selection", "TrimAndCleanSelection", 384, False
    AddButton group, "Selection to &values", "SelectionToValues", 370, False
    AddButton group, "Remove &hyperlinks from selection", "RemoveHyperlinksFromSelection", 1691, False

    '--- Loose ends -----------------------------------------------------------
    AddButton root, "&Unhide all sheets", "UnhideAllSheets", 2110, True
    AddButton root, "Bac&kup this workbook", "BackupActiveWorkbook", 3, False
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
    Loop While guard < 100
    On Error GoTo 0
End Sub

'==============================================================================
' Private helpers
'==============================================================================
Private Function AddPopup(ByVal parent As CommandBarPopup, ByVal caption As String, _
                          ByVal startsGroup As Boolean) As CommandBarPopup
    Dim popup As CommandBarPopup

    On Error Resume Next
    Set popup = parent.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    If popup Is Nothing Then Exit Function

    popup.Caption = caption
    popup.Tag = MENU_TAG
    popup.BeginGroup = startsGroup
    On Error GoTo 0

    Set AddPopup = popup
End Function

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
