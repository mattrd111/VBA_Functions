Attribute VB_Name = "vba_io"
Option Explicit

' Round-trip between a workbook's VBA project and the repo's text files, so git
' (and the review agents) see real source instead of an opaque vbaProject.bin.
'
' One-time setup: Excel > File > Options > Trust Center > Trust Center Settings
' > Macro Settings > tick "Trust access to the VBA project object model".
'
' ExportModules writes every component to <repo>\src\ (tests go to <repo>\tests\).
' ImportModules replaces the project's components from those folders.
' Both take the repo root as an optional argument. With no argument they work it
' out: a workbook in <repo>\workbook\ resolves to <repo>, anything else uses the
' workbook's own folder. So from a workbook in workbook\, just run ExportModules.

Public Sub ExportModules(Optional ByVal repoRoot As String = "")
    Dim comp As Object, targetDir As String, ext As String, saved As Long
    On Error GoTo Fail

    If Len(repoRoot) = 0 Then repoRoot = RepoRootFromWorkbook()
    For Each comp In ThisWorkbook.VBProject.VBComponents
        ext = ComponentExtension(comp.Type)
        If Len(ext) = 0 Then GoTo NextComp
        If comp.Name = "vba_io" Then GoTo NextComp   ' this helper stays out of src\

        If Left$(comp.Name, 5) = "Test_" Or comp.Name Like "*TestRunner*" Or comp.Name = "modAssert" Then
            targetDir = repoRoot & "\tests"
        Else
            targetDir = repoRoot & "\src"
        End If
        EnsureFolder targetDir
        comp.Export targetDir & "\" & comp.Name & ext
        saved = saved + 1
NextComp:
    Next comp

    MsgBox saved & " component(s) exported under " & repoRoot, vbInformation
    Exit Sub
Fail:
    MsgBox "Export failed (" & Err.Number & "): " & Err.Description & vbLf & vbLf & _
           "If this is error 1004, enable 'Trust access to the VBA project object model'.", vbExclamation
End Sub

Public Sub ImportModules(Optional ByVal repoRoot As String = "")
    Dim folders As Variant, folder As Variant, file As String, imported As Long
    Dim proj As Object
    On Error GoTo Fail

    If Len(repoRoot) = 0 Then repoRoot = RepoRootFromWorkbook()
    Set proj = ThisWorkbook.VBProject
    folders = Array(repoRoot & "\src", repoRoot & "\tests")

    For Each folder In folders
        file = Dir(folder & "\*.*")
        Do While Len(file) > 0
            If IsSourceFile(file) Then
                RemoveComponent proj, BaseName(file)
                proj.VBComponents.Import folder & "\" & file
                imported = imported + 1
            End If
            file = Dir
        Loop
    Next folder

    MsgBox imported & " component(s) imported from " & repoRoot, vbInformation
    Exit Sub
Fail:
    MsgBox "Import failed (" & Err.Number & "): " & Err.Description, vbExclamation
End Sub

Private Function RepoRootFromWorkbook() As String
    ' The workbook normally lives in <repo>\workbook\, so the repo root is one
    ' level up. Anywhere else, assume the workbook sits in the repo root itself.
    Dim path As String, leaf As String
    path = ThisWorkbook.path
    leaf = LCase$(Mid$(path, InStrRev(path, "\") + 1))
    If leaf = "workbook" Then
        RepoRootFromWorkbook = Left$(path, InStrRev(path, "\") - 1)
    Else
        RepoRootFromWorkbook = path
    End If
End Function

Private Sub RemoveComponent(ByVal proj As Object, ByVal name As String)
    Dim comp As Object
    On Error Resume Next
    Set comp = proj.VBComponents(name)
    On Error GoTo 0
    If comp Is Nothing Then Exit Sub
    ' Document modules (ThisWorkbook, sheets) cannot be removed - skip them.
    If comp.Type = 100 Then Exit Sub
    proj.VBComponents.Remove comp
End Sub

Private Function ComponentExtension(ByVal compType As Long) As String
    Select Case compType
        Case 1: ComponentExtension = ".bas"   ' standard module
        Case 2: ComponentExtension = ".cls"   ' class module
        Case 3: ComponentExtension = ".frm"   ' user form
        Case Else: ComponentExtension = ""    ' 100 = document module, exported separately if wanted
    End Select
End Function

Private Function IsSourceFile(ByVal file As String) As Boolean
    Dim ext As String
    ext = LCase$(Mid$(file, InStrRev(file, ".")))
    IsSourceFile = (ext = ".bas" Or ext = ".cls" Or ext = ".frm")
End Function

Private Function BaseName(ByVal file As String) As String
    BaseName = Left$(file, InStrRev(file, ".") - 1)
End Function

Private Sub EnsureFolder(ByVal path As String)
    If Len(Dir(path, vbDirectory)) = 0 Then MkDir path
End Sub
