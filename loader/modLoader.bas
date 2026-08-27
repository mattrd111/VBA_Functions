Attribute VB_Name = "modLoader"
'==============================================================================
' modLoader - keeps Workbook Doctor up to date from a shared folder
'------------------------------------------------------------------------------
' An add-in cannot replace itself while Excel has it open, which is why the
' obvious "download over the top on startup" does not work. This gets round it
' by never installing the real add-in at all.
'
'   the loader        installed on each machine, tiny, rarely changes
'   the master copy   sits in the synced SharePoint folder, never installed
'   the cache         %LOCALAPPDATA%\WorkbookDoctor\WorkbookDoctor.xlam
'
' On Excel start the loader compares the master against the cache, copies it
' down if it is newer, and opens the cache. The master is never the open file,
' so you can replace it whenever you like - including while people have Excel
' running. Anyone offline keeps working from their cache.
'
' BEFORE YOU DISTRIBUTE THIS: set SOURCE_RELATIVE below to where the master
' sits inside each person's sync folder. The bit before it is their own user
' folder, which is why one constant works for everybody.
'==============================================================================
Option Explicit

' EDIT ME. Everything after %USERPROFILE%\ - a synced SharePoint library looks
' like "<Tenant>\<Site> - <Library>\...".
Private Const SOURCE_RELATIVE As String = "Alpha FMC\Tools - Documents\Workbook Doctor\WorkbookDoctor.xlam"

Private Const ADDIN_FILE As String = "WorkbookDoctor.xlam"
Private Const CACHE_FOLDER As String = "WorkbookDoctor"
Private Const SETTINGS_APP As String = "WorkbookDoctor"
Private Const SETTINGS_SECTION As String = "Loader"
Private Const SETTINGS_KEY As String = "SourcePath"
Private Const MENU_TAG As String = "WorkbookDoctorLoaderMenu"

'==============================================================================
' Startup
'==============================================================================
'------------------------------------------------------------------------------
' StartWorkbookDoctor
' Called a moment after Excel finishes starting - opening a workbook from inside
' Workbook_Open itself is unreliable, so ThisWorkbook defers to here.
'------------------------------------------------------------------------------
Public Sub StartWorkbookDoctor()
    Dim source As String
    Dim cache As String
    Dim updated As Boolean

    BuildLoaderMenu

    cache = CachePath()
    source = ResolveSource()

    If Len(source) > 0 Then
        If SourceIsNewer(source, cache) Then updated = CopyDown(source, cache)
    End If

    If Len(Dir$(cache)) = 0 Then
        ' Nothing cached and nothing to copy - first run, off the network.
        MsgBox "Workbook Doctor could not be loaded." & vbNewLine & vbNewLine & _
               IIf(Len(source) = 0, _
                   "The shared copy was not found. Use Add-ins tab > Workbook Doctor " & _
                   "Updates > Change the shared folder to point at it.", _
                   "The shared copy at" & vbNewLine & source & vbNewLine & _
                   "could not be copied, and there is no local copy yet."), _
               vbExclamation, "Workbook Doctor"
        Exit Sub
    End If

    OpenAddIn cache

    If updated Then
        On Error Resume Next
        Application.StatusBar = "Workbook Doctor updated to the version dated " & _
                                Format$(FileDateTime(cache), "dd mmm yyyy hh:nn")
        On Error GoTo 0
    End If
End Sub

'==============================================================================
' Menu commands
'==============================================================================
'------------------------------------------------------------------------------
' CheckForUpdatesNow
' Swaps the running add-in for a newer one without restarting Excel: the loaded
' copy is closed, the new one copied down, and it is opened again.
'------------------------------------------------------------------------------
Public Sub CheckForUpdatesNow()
    Dim source As String
    Dim cache As String

    cache = CachePath()
    source = ResolveSource()

    If Len(source) = 0 Then
        MsgBox "The shared copy has not been found." & vbNewLine & vbNewLine & _
               "Use 'Change the shared folder' to point at it.", vbExclamation, "Workbook Doctor"
        Exit Sub
    End If

    If Not SourceIsNewer(source, cache) Then
        MsgBox "You already have the current version." & vbNewLine & vbNewLine & _
               "Shared copy: " & FileStamp(source) & vbNewLine & _
               "Yours:       " & FileStamp(cache), vbInformation, "Workbook Doctor"
        Exit Sub
    End If

    If MsgBox("A newer Workbook Doctor is available." & vbNewLine & vbNewLine & _
              "Shared copy: " & FileStamp(source) & vbNewLine & _
              "Yours:       " & FileStamp(cache) & vbNewLine & vbNewLine & _
              "Update now? The add-in will reload - anything you have open is untouched.", _
              vbYesNo + vbQuestion, "Workbook Doctor") <> vbYes Then Exit Sub

    CloseAddIn

    If Not CopyDown(source, cache) Then
        OpenAddIn cache                          ' put back what was there
        MsgBox "The update could not be copied." & vbNewLine & vbNewLine & _
               "The version you had is still loaded. Try again later, or close " & _
               "every Excel window and run this again.", vbExclamation, "Workbook Doctor"
        Exit Sub
    End If

    OpenAddIn cache

    MsgBox "Updated to the version dated " & FileStamp(cache) & ".", _
           vbInformation, "Workbook Doctor"
End Sub

'------------------------------------------------------------------------------
' ChangeSharedFolder
' Points this machine at the master copy and remembers it.
'------------------------------------------------------------------------------
Public Sub ChangeSharedFolder()
    Dim dialog As Object
    Dim chosen As String

    On Error GoTo Done
    Set dialog = Application.FileDialog(3)                  ' msoFileDialogFilePicker
    With dialog
        .title = "Find the shared WorkbookDoctor.xlam"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel add-in", "*.xlam"
        If .Show <> -1 Then Exit Sub
        chosen = .SelectedItems(1)
    End With

    If Len(chosen) = 0 Then Exit Sub

    SaveSetting SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY, chosen

    MsgBox "Workbook Doctor will now update from:" & vbNewLine & vbNewLine & chosen, _
           vbInformation, "Workbook Doctor"

    CheckForUpdatesNow
    Exit Sub
Done:
End Sub

'------------------------------------------------------------------------------
' ShowLoaderStatus
' Where it is loading from and how old each copy is - the first thing to look at
' when someone says their menu has gone.
'------------------------------------------------------------------------------
Public Sub ShowLoaderStatus()
    Dim source As String
    Dim cache As String
    Dim loaded As String

    source = ResolveSource()
    cache = CachePath()
    loaded = IIf(AddInIsOpen(), "yes", "no")

    MsgBox "Workbook Doctor loader" & vbNewLine & String$(46, "-") & vbNewLine & vbNewLine & _
           "Shared copy" & vbNewLine & _
           IIf(Len(source) = 0, "  not found", "  " & source & vbNewLine & "  " & FileStamp(source)) & _
           vbNewLine & vbNewLine & _
           "Your copy" & vbNewLine & "  " & cache & vbNewLine & "  " & FileStamp(cache) & _
           vbNewLine & vbNewLine & _
           "Loaded right now: " & loaded & vbNewLine & vbNewLine & _
           "The shared copy is only ever read. Your copy is what Excel runs, which " & _
           "is why it keeps working when you are off the network.", _
           vbInformation, "Workbook Doctor"
End Sub

'==============================================================================
' The loader's own menu
'==============================================================================
Public Sub BuildLoaderMenu()
    Dim root As CommandBarPopup

    RemoveLoaderMenu

    On Error Resume Next
    Set root = Application.CommandBars("Worksheet Menu Bar").Controls.Add( _
                   Type:=msoControlPopup, Temporary:=True)
    If root Is Nothing Then Exit Sub

    root.Caption = "Workbook Doctor &Updates"
    root.Tag = MENU_TAG

    AddButton root, "&Check for updates now", "CheckForUpdatesNow", 1018
    AddButton root, "Change the &shared folder...", "ChangeSharedFolder", 23
    AddButton root, "&Where is it loading from?", "ShowLoaderStatus", 487
    On Error GoTo 0
End Sub

Public Sub RemoveLoaderMenu()
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
    Loop While guard < 20
    On Error GoTo 0
End Sub

'==============================================================================
' Private helpers
'==============================================================================
' Where the master copy is: whatever this machine was told, otherwise the
' standard place inside the user's own sync folder.
Private Function ResolveSource() As String
    Dim remembered As String
    Dim standard As String

    remembered = GetSetting(SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY, "")
    If Len(remembered) > 0 Then
        If Len(Dir$(remembered)) > 0 Then ResolveSource = remembered: Exit Function
    End If

    standard = Environ$("USERPROFILE")
    If Len(standard) = 0 Then Exit Function
    If Right$(standard, 1) <> "\" Then standard = standard & "\"
    standard = standard & SOURCE_RELATIVE

    On Error Resume Next
    If Len(Dir$(standard)) > 0 Then
        ResolveSource = standard
        SaveSetting SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY, standard
    End If
    On Error GoTo 0
End Function

Private Function CachePath() As String
    Dim folder As String

    folder = Environ$("LOCALAPPDATA")
    If Len(folder) = 0 Then folder = Environ$("APPDATA")
    If Len(folder) = 0 Then Exit Function
    If Right$(folder, 1) <> "\" Then folder = folder & "\"

    CachePath = folder & CACHE_FOLDER & "\" & ADDIN_FILE
End Function

Private Function SourceIsNewer(ByVal source As String, ByVal cache As String) As Boolean
    On Error GoTo Done
    If Len(Dir$(source)) = 0 Then Exit Function
    If Len(Dir$(cache)) = 0 Then SourceIsNewer = True: Exit Function

    ' A minute of slack, because a synced file's stamp can wobble by seconds.
    SourceIsNewer = (FileDateTime(source) > DateAdd("n", 1, FileDateTime(cache)))
Done:
End Function

Private Function CopyDown(ByVal source As String, ByVal cache As String) As Boolean
    Dim folder As String

    On Error GoTo Done
    folder = Left$(cache, InStrRev(cache, "\") - 1)
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder

    FileCopy source, cache
    CopyDown = (Len(Dir$(cache)) > 0)
Done:
End Function

Private Sub OpenAddIn(ByVal path As String)
    On Error Resume Next
    If AddInIsOpen() Then Exit Sub
    Application.Workbooks.Open Filename:=path, ReadOnly:=True
    On Error GoTo 0
End Sub

Private Sub CloseAddIn()
    Dim wb As Workbook

    On Error Resume Next
    Set wb = Application.Workbooks(ADDIN_FILE)
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function AddInIsOpen() As Boolean
    Dim wb As Workbook

    On Error Resume Next
    Set wb = Application.Workbooks(ADDIN_FILE)
    On Error GoTo 0
    AddInIsOpen = Not wb Is Nothing
End Function

Private Function FileStamp(ByVal path As String) As String
    On Error GoTo Missing
    If Len(Dir$(path)) = 0 Then GoTo Missing
    FileStamp = Format$(FileDateTime(path), "dd mmm yyyy hh:nn")
    Exit Function
Missing:
    FileStamp = "not there"
End Function

Private Sub AddButton(ByVal parent As CommandBarPopup, ByVal caption As String, _
                      ByVal macroName As String, ByVal iconId As Long)
    Dim button As CommandBarButton

    On Error Resume Next
    Set button = parent.Controls.Add(Type:=msoControlButton, Temporary:=True)
    If button Is Nothing Then Exit Sub

    button.Caption = caption
    button.OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
    button.Tag = MENU_TAG
    button.Style = msoButtonIconAndCaption
    If iconId > 0 Then button.FaceId = iconId
    On Error GoTo 0
End Sub
