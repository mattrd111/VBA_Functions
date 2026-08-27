Attribute VB_Name = "BuildAddIn"
'==============================================================================
' BuildAddIn - builds WorkbookDoctor.xlam without PowerShell
'------------------------------------------------------------------------------
' For machines where running .ps1 files is blocked, which is most managed ones.
' This does the same job from inside Excel.
'
' HOW TO USE IT
'
'   1  Excel: File > Options > Trust Center > Trust Center Settings >
'      Macro Settings > tick "Trust access to the VBA project object model".
'      The build cannot write code into a workbook without it. Untick it
'      afterwards if you would rather - it is only needed while building.
'
'   2  Open a new blank workbook. Alt+F11 for the editor.
'
'   3  File > Import File... and pick THIS file (build\BuildAddIn.bas).
'
'   4  Put the cursor in one of these and press F5:
'
'      BuildWorkbookDoctor     build it and install it on this machine.
'                              The one to use if you just want it working here.
'
'      BuildMasterForSharing   build it into a folder you pick, without
'                              installing. This is the copy that goes in the
'                              shared SharePoint folder for the loader to
'                              collect - see loader\README.md.
'
'      BuildLoader             build and install the loader, which is what each
'                              person installs once when you are updating from
'                              a shared folder.
'
'   5  It asks for the folder holding src and addin - point it at the folder
'      you unzipped, then let it run.
'
'   6  Close this builder workbook without saving. The add-in is installed and
'      the menu is on the Add-ins tab.
'
' This module is NOT part of the add-in. It never gets imported into it.
'==============================================================================
Option Explicit

Private Const ADDIN_NAME As String = "WorkbookDoctor.xlam"
Private Const LOADER_NAME As String = "WorkbookDoctorLoader.xlam"
Private Const ADDIN_TITLE As String = "Workbook Doctor"

'------------------------------------------------------------------------------
' BuildWorkbookDoctor
' Build the add-in and switch it on here. The everyday one.
'------------------------------------------------------------------------------
Public Sub BuildWorkbookDoctor()
    BuildPackage False, True
End Sub

'------------------------------------------------------------------------------
' BuildMasterForSharing
' Build the add-in into a folder you choose and do not install it. Put the
' result in the shared folder the loader reads from.
'------------------------------------------------------------------------------
Public Sub BuildMasterForSharing()
    BuildPackage False, False
End Sub

'------------------------------------------------------------------------------
' BuildLoader
' Build the loader and switch it on here. Each person installs this once; it
' keeps itself pointed at the shared copy from then on.
'------------------------------------------------------------------------------
Public Sub BuildLoader()
    BuildPackage True, True
End Sub

'==============================================================================
' The worker
'==============================================================================
Private Sub BuildPackage(ByVal loaderOnly As Boolean, ByVal installHere As Boolean)
    Dim sourceFolder As String
    Dim target As Workbook
    Dim project As Object
    Dim imported As Long
    Dim outputPath As String
    Dim report As String
    Dim outputName As String
    Dim classPath As String

    If Not TrustIsOn() Then
        MsgBox "Excel will not let this write code into a workbook yet." & vbNewLine & vbNewLine & _
               "File > Options > Trust Center > Trust Center Settings >" & vbNewLine & _
               "Macro Settings > tick 'Trust access to the VBA project object model'" & vbNewLine & _
               vbNewLine & "then run this again.", vbExclamation, ADDIN_TITLE
        Exit Sub
    End If

    sourceFolder = AskForFolder()
    If Len(sourceFolder) = 0 Then Exit Sub

    If Not FolderHasSource(sourceFolder, loaderOnly) Then
        MsgBox "That folder does not hold the source." & vbNewLine & vbNewLine & _
               IIf(loaderOnly, "Expected a 'loader' subfolder.", _
                               "Expected 'src' and 'addin' subfolders.") & vbNewLine & _
               "Pick the folder you unzipped.", vbExclamation, ADDIN_TITLE
        Exit Sub
    End If

    outputName = IIf(loaderOnly, LOADER_NAME, ADDIN_NAME)

    If installHere Then
        outputPath = Application.UserLibraryPath & outputName
        If Not ReadyToOverwrite(outputName, outputPath) Then Exit Sub
    Else
        outputPath = AskForFolder("Where should " & outputName & " be written?")
        If Len(outputPath) = 0 Then Exit Sub
        outputPath = outputPath & outputName
    End If

    Application.ScreenUpdating = False
    On Error GoTo Failed

    Set target = Application.Workbooks.Add
    Set project = target.VBProject

    If loaderOnly Then
        imported = ImportFolder(project, sourceFolder & "loader\", report)
        classPath = sourceFolder & "loader\ThisWorkbook.cls"
    Else
        imported = ImportFolder(project, sourceFolder & "src\", report)
        imported = imported + ImportFolder(project, sourceFolder & "addin\", report)
        classPath = sourceFolder & "addin\ThisWorkbook.cls"
    End If

    If imported = 0 Then
        target.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "No modules were found to import. Check the folder you picked.", _
               vbExclamation, ADDIN_TITLE
        Exit Sub
    End If

    If Not MergeThisWorkbook(project, target, classPath) Then
        report = report & vbNewLine & "  ThisWorkbook.cls could not be merged - the menu will not appear."
    End If

    SetTitle target, loaderOnly
    target.IsAddin = True

    Application.DisplayAlerts = False
    target.SaveAs Filename:=outputPath, FileFormat:=55        ' xlOpenXMLAddIn
    Application.DisplayAlerts = True

    target.Close SaveChanges:=False
    Application.ScreenUpdating = True

    If installHere Then Install outputPath

    MsgBox outputName & IIf(installHere, " built and installed.", " built.") & vbNewLine & vbNewLine & _
           imported & " modules imported." & vbNewLine & _
           outputPath & vbNewLine & vbNewLine & _
           IIf(installHere, _
               "Look on the Add-ins tab of the ribbon." & vbNewLine & vbNewLine, _
               "Put this in the shared folder the loader reads from." & vbNewLine & vbNewLine) & _
           "Close this builder workbook without saving." & _
           IIf(Len(report) > 0, vbNewLine & vbNewLine & "Notes:" & report, ""), _
           vbInformation, ADDIN_TITLE
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not target Is Nothing Then target.Close SaveChanges:=False
    On Error GoTo 0

    MsgBox "The build stopped." & vbNewLine & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description & _
           IIf(Err.Number = 1004, vbNewLine & vbNewLine & _
               "That is usually the trust setting. See the notes at the top of this module.", ""), _
           vbCritical, ADDIN_TITLE
End Sub

'==============================================================================
' Private helpers
'==============================================================================
' True when Excel will let VBA touch a workbook's code.
Private Function TrustIsOn() As Boolean
    Dim components As Object

    On Error Resume Next
    Set components = ThisWorkbook.VBProject.VBComponents
    TrustIsOn = (Err.Number = 0) And Not (components Is Nothing)
    On Error GoTo 0
End Function

Private Function AskForFolder(Optional ByVal prompt As String = _
                             "Pick the folder holding src, addin and loader") As String
    Dim dialog As Object
    Dim path As String

    On Error GoTo Done
    Set dialog = Application.FileDialog(4)                     ' msoFileDialogFolderPicker
    With dialog
        .title = prompt
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        path = .SelectedItems(1)
    End With

    If Right$(path, 1) <> "\" Then path = path & "\"
    AskForFolder = path
Done:
End Function

Private Function FolderHasSource(ByVal folder As String, ByVal loaderOnly As Boolean) As Boolean
    If loaderOnly Then
        FolderHasSource = (Len(Dir$(folder & "loader\*.bas")) > 0)
    Else
        FolderHasSource = (Len(Dir$(folder & "src\*.bas")) > 0) And _
                          (Len(Dir$(folder & "addin\*.bas")) > 0)
    End If
End Function

' Imports every .bas in a folder. Returns how many went in.
Private Function ImportFolder(ByVal project As Object, ByVal folder As String, _
                              ByRef report As String) As Long
    Dim fileName As String
    Dim count As Long

    fileName = Dir$(folder & "*.bas")
    Do While Len(fileName) > 0
        On Error Resume Next
        project.VBComponents.Import folder & fileName
        If Err.Number <> 0 Then
            report = report & vbNewLine & "  " & fileName & " - " & Err.Description
            Err.Clear
        Else
            count = count + 1
        End If
        On Error GoTo 0
        fileName = Dir$
    Loop

    ImportFolder = count
End Function

' A document module cannot be imported as a new component, so its code is copied
' into the ThisWorkbook that already exists. Everything above Option Explicit is
' the class header and is dropped.
Private Function MergeThisWorkbook(ByVal project As Object, ByVal target As Workbook, _
                                   ByVal path As String) As Boolean
    Dim handle As Integer
    Dim textLine As String
    Dim code As String
    Dim started As Boolean

    If Len(Dir$(path)) = 0 Then Exit Function

    On Error GoTo Done
    handle = FreeFile
    Open path For Input As #handle
    Do While Not EOF(handle)
        Line Input #handle, textLine
        If Not started Then
            If InStr(1, textLine, "Option Explicit", vbTextCompare) > 0 Then started = True
        End If
        If started Then code = code & textLine & vbCrLf
    Loop
    Close #handle

    If Len(code) = 0 Then Exit Function

    ' Workbook.CodeName is what the module is actually called in the project,
    ' which is not "ThisWorkbook" on a localised Excel.
    project.VBComponents(target.CodeName).CodeModule.AddFromString code
    MergeThisWorkbook = True
Done:
End Function

Private Sub SetTitle(ByVal target As Workbook, ByVal loaderOnly As Boolean)
    On Error Resume Next
    target.BuiltinDocumentProperties("Title").value = _
        ADDIN_TITLE & IIf(loaderOnly, " Loader", "")
    target.BuiltinDocumentProperties("Comments").value = _
        IIf(loaderOnly, "Keeps Workbook Doctor up to date from a shared folder.", _
                        "Clean-up, audit and fund tools for Excel.") & _
        " github.com/mattrd111/VBA_Functions"
    On Error GoTo 0
End Sub

Private Sub Install(ByVal path As String)
    Dim item As AddIn

    On Error Resume Next
    Set item = Application.AddIns.Add(path, False)
    If item Is Nothing Then Exit Sub
    item.Installed = True
    On Error GoTo 0
End Sub

' An add-in already loaded holds its file open, so it has to go first.
Private Function ReadyToOverwrite(ByVal fileName As String, ByVal path As String) As Boolean
    Dim item As AddIn
    Dim wasInstalled As Boolean

    ReadyToOverwrite = True
    If Len(Dir$(path)) = 0 Then Exit Function

    On Error Resume Next
    For Each item In Application.AddIns
        If StrComp(item.Name, fileName, vbTextCompare) = 0 Then
            If item.Installed Then
                wasInstalled = True
                item.Installed = False
            End If
        End If
    Next item
    On Error GoTo 0

    If wasInstalled Then
        MsgBox "An older " & fileName & " was loaded, so it has been switched off " & _
               "and will be replaced.", vbInformation, ADDIN_TITLE
    End If
End Function
