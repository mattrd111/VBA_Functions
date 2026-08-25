Attribute VB_Name = "modFile"
'==============================================================================
' modFile - file and folder helpers
'------------------------------------------------------------------------------
' Existence checks, path building, folder listings, text and CSV files, and the
' file and folder pickers.
'
' Standalone: no dependency on the other modules in this repository.
' Uses late-bound Scripting.FileSystemObject and ADODB.Stream, so no project
' reference is needed.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' FileExists / FolderExists
' True only for a file / only for a folder, so a folder never passes as a file.
'------------------------------------------------------------------------------
Public Function FileExists(ByVal filePath As String) As Boolean
    If Len(filePath) = 0 Then Exit Function
    On Error Resume Next
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(filePath)
    On Error GoTo 0
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    If Len(folderPath) = 0 Then Exit Function
    On Error Resume Next
    FolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(folderPath)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' EnsureFolder
' Creates a folder and every missing parent above it. True when the folder
' exists afterwards.
'
'   EnsureFolder "C:\Reports\2024\Q3"
'------------------------------------------------------------------------------
Public Function EnsureFolder(ByVal folderPath As String) As Boolean
    Dim fso As Object
    Dim parent As String

    If Len(folderPath) = 0 Then Exit Function

    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")

    folderPath = Replace$(folderPath, "/", "\")
    Do While Len(folderPath) > 3 And Right$(folderPath, 1) = "\"
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop
    If fso.FolderExists(folderPath) Then EnsureFolder = True: Exit Function

    parent = ParentFolder(folderPath)
    If Len(parent) = 0 Then Exit Function
    If StrComp(parent, folderPath, vbTextCompare) = 0 Then Exit Function
    If Not fso.FolderExists(parent) Then
        If Not EnsureFolder(parent) Then Exit Function
    End If

    fso.CreateFolder folderPath
    EnsureFolder = fso.FolderExists(folderPath)
    Exit Function
Failed:
    EnsureFolder = False
End Function

'------------------------------------------------------------------------------
' JoinPath
' Joins path parts with single backslashes, however many the caller supplied.
'
'   JoinPath("C:\Reports\", "\2024", "summary.xlsx")
'       ->  "C:\Reports\2024\summary.xlsx"
'------------------------------------------------------------------------------
Public Function JoinPath(ParamArray parts() As Variant) As String
    Dim i As Long
    Dim piece As String
    Dim result As String

    For i = LBound(parts) To UBound(parts)
        piece = Replace$(CStr(parts(i)), "/", "\")
        Do While Left$(piece, 1) = "\"
            piece = Mid$(piece, 2)
        Loop
        Do While Right$(piece, 1) = "\"
            piece = Left$(piece, Len(piece) - 1)
        Loop
        If Len(piece) > 0 Then
            If Len(result) = 0 Then
                result = piece
            Else
                result = result & "\" & piece
            End If
        End If
    Next i

    JoinPath = result
End Function

'------------------------------------------------------------------------------
' FileNameOnly / BaseName / FileExtension / ParentFolder
' Pure string work - the file does not have to exist.
'
'   FileNameOnly("C:\data\report.final.xlsx")  ->  "report.final.xlsx"
'   BaseName("C:\data\report.final.xlsx")      ->  "report.final"
'   FileExtension("C:\data\report.final.xlsx") ->  "xlsx"
'   ParentFolder("C:\data\report.final.xlsx")  ->  "C:\data"
'------------------------------------------------------------------------------
Public Function FileNameOnly(ByVal filePath As String) As String
    Dim p As Long

    filePath = Replace$(filePath, "/", "\")
    p = InStrRev(filePath, "\")
    If p = 0 Then FileNameOnly = filePath Else FileNameOnly = Mid$(filePath, p + 1)
End Function

Public Function BaseName(ByVal filePath As String) As String
    Dim name As String
    Dim p As Long

    name = FileNameOnly(filePath)
    p = InStrRev(name, ".")
    If p <= 1 Then BaseName = name Else BaseName = Left$(name, p - 1)
End Function

Public Function FileExtension(ByVal filePath As String) As String
    Dim name As String
    Dim p As Long

    name = FileNameOnly(filePath)
    p = InStrRev(name, ".")
    If p > 1 And p < Len(name) Then FileExtension = Mid$(name, p + 1)
End Function

Public Function ParentFolder(ByVal filePath As String) As String
    Dim p As Long

    filePath = Replace$(filePath, "/", "\")
    Do While Len(filePath) > 3 And Right$(filePath, 1) = "\"
        filePath = Left$(filePath, Len(filePath) - 1)
    Loop
    p = InStrRev(filePath, "\")
    If p > 1 Then ParentFolder = Left$(filePath, p - 1)
End Function

'------------------------------------------------------------------------------
' UniqueFilePath
' A path that is not in use yet, inserting (1), (2) ... before the extension -
' so an overnight run never overwrites yesterday's output.
'------------------------------------------------------------------------------
Public Function UniqueFilePath(ByVal filePath As String) As String
    Dim folder As String
    Dim stem As String
    Dim ext As String
    Dim candidate As String
    Dim n As Long

    candidate = filePath
    If Not FileExists(candidate) Then UniqueFilePath = candidate: Exit Function

    folder = ParentFolder(filePath)
    stem = BaseName(filePath)
    ext = FileExtension(filePath)

    Do
        n = n + 1
        candidate = JoinPath(folder, stem & " (" & n & ")" & IIf(Len(ext) > 0, "." & ext, ""))
    Loop While FileExists(candidate) And n < 10000

    UniqueFilePath = candidate
End Function

'------------------------------------------------------------------------------
' ListFiles
' Full paths of the files in a folder as a Collection, optionally recursing
' into subfolders. The pattern is matched case insensitively.
'
'   For Each p In ListFiles("C:\Inbox", "*.csv", True)
'------------------------------------------------------------------------------
Public Function ListFiles(ByVal folderPath As String, _
                          Optional ByVal pattern As String = "*.*", _
                          Optional ByVal includeSubfolders As Boolean = False) As Collection
    Dim fso As Object
    Dim result As Collection

    Set result = New Collection
    Set ListFiles = result

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function
    If Not fso.FolderExists(folderPath) Then Exit Function
    On Error GoTo 0

    CollectFiles fso.GetFolder(folderPath), pattern, includeSubfolders, result
End Function

'------------------------------------------------------------------------------
' ReadTextFile
' Whole file as a string. Pass a charset such as "utf-8" for files that are not
' in the system code page.
'------------------------------------------------------------------------------
Public Function ReadTextFile(ByVal filePath As String, Optional ByVal charset As String = "") As String
    Dim stream As Object
    Dim fso As Object
    Dim handle As Object

    If Not FileExists(filePath) Then Exit Function

    On Error GoTo Failed
    If Len(charset) = 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set handle = fso.OpenTextFile(filePath, 1)          ' ForReading
        If Not handle.AtEndOfStream Then ReadTextFile = handle.ReadAll
        handle.Close
    Else
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2                                     ' adTypeText
        stream.charset = charset
        stream.Open
        stream.LoadFromFile filePath
        ReadTextFile = stream.ReadText(-1)                  ' adReadAll
        stream.Close
    End If
    Exit Function
Failed:
    ReadTextFile = ""
End Function

'------------------------------------------------------------------------------
' WriteTextFile
' Writes a string to a file in the system code page, creating parent folders as
' needed. Returns True on success.
'------------------------------------------------------------------------------
Public Function WriteTextFile(ByVal filePath As String, ByVal content As String, _
                              Optional ByVal appendToFile As Boolean = False) As Boolean
    Dim fso As Object
    Dim handle As Object

    If Len(filePath) = 0 Then Exit Function
    EnsureFolder ParentFolder(filePath)

    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If appendToFile And FileExists(filePath) Then
        Set handle = fso.OpenTextFile(filePath, 8)          ' ForAppending
    Else
        Set handle = fso.CreateTextFile(filePath, True)
    End If
    handle.Write content
    handle.Close
    WriteTextFile = True
    Exit Function
Failed:
    WriteTextFile = False
End Function

'------------------------------------------------------------------------------
' WriteUtf8File
' Writes UTF-8, with or without the byte order mark. Excel needs the mark to
' open a CSV with accents correctly; most other tools would rather not see it.
'------------------------------------------------------------------------------
Public Function WriteUtf8File(ByVal filePath As String, ByVal content As String, _
                              Optional ByVal includeByteOrderMark As Boolean = True) As Boolean
    Dim textStream As Object
    Dim binaryStream As Object

    If Len(filePath) = 0 Then Exit Function
    EnsureFolder ParentFolder(filePath)

    On Error GoTo Failed
    Set textStream = CreateObject("ADODB.Stream")
    textStream.Type = 2                                     ' adTypeText
    textStream.charset = "utf-8"
    textStream.Open
    textStream.WriteText content

    If includeByteOrderMark Then
        textStream.SaveToFile filePath, 2                   ' adSaveCreateOverWrite
    Else
        textStream.Position = 0
        textStream.Type = 1                                 ' adTypeBinary
        textStream.Position = 3                             ' step over the BOM
        Set binaryStream = CreateObject("ADODB.Stream")
        binaryStream.Type = 1
        binaryStream.Open
        textStream.CopyTo binaryStream
        binaryStream.SaveToFile filePath, 2
        binaryStream.Close
    End If
    textStream.Close

    WriteUtf8File = True
    Exit Function
Failed:
    WriteUtf8File = False
End Function

'------------------------------------------------------------------------------
' ExportRangeToCsv
' Writes a range to CSV with proper quoting, without going through
' SaveAs - so the workbook keeps its name, format and unsaved changes.
' Dates are written as yyyy-mm-dd, which every system reads the same way.
'
'   ExportRangeToCsv ws.Range("A1:F500"), "C:\Out\extract.csv"
'------------------------------------------------------------------------------
Public Function ExportRangeToCsv(ByVal rng As Range, ByVal filePath As String, _
                                 Optional ByVal delimiter As String = ",", _
                                 Optional ByVal includeByteOrderMark As Boolean = True) As Boolean
    Dim data As Variant
    Dim lines() As String
    Dim fields() As String
    Dim r As Long, c As Long

    If rng Is Nothing Then Exit Function
    If Len(filePath) = 0 Then Exit Function

    If rng.Cells.CountLarge = 1 Then
        ReDim data(1 To 1, 1 To 1)
        data(1, 1) = rng.value
    Else
        data = rng.value
    End If

    ReDim lines(1 To UBound(data, 1) - LBound(data, 1) + 1)
    ReDim fields(1 To UBound(data, 2) - LBound(data, 2) + 1)

    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            fields(c - LBound(data, 2) + 1) = CsvField(data(r, c), delimiter)
        Next c
        lines(r - LBound(data, 1) + 1) = Join(fields, delimiter)
    Next r

    ExportRangeToCsv = WriteUtf8File(filePath, Join(lines, vbCrLf) & vbCrLf, includeByteOrderMark)
End Function

'------------------------------------------------------------------------------
' PickFile / PickFolder
' The standard dialogs, returning "" when the user cancels.
'
'   path = PickFile("Choose the extract", "CSV files", "*.csv")
'------------------------------------------------------------------------------
Public Function PickFile(Optional ByVal title As String = "Select a file", _
                         Optional ByVal filterName As String = "Excel files", _
                         Optional ByVal filterPattern As String = "*.xls*") As String
    Dim dialog As Object

    On Error GoTo Failed
    Set dialog = Application.FileDialog(3)                  ' msoFileDialogFilePicker
    With dialog
        .title = title
        .AllowMultiSelect = False
        .Filters.Clear
        If Len(filterPattern) > 0 Then .Filters.Add filterName, filterPattern
        .Filters.Add "All files", "*.*"
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
Failed:
End Function

Public Function PickFolder(Optional ByVal title As String = "Select a folder") As String
    Dim dialog As Object

    On Error GoTo Failed
    Set dialog = Application.FileDialog(4)                  ' msoFileDialogFolderPicker
    With dialog
        .title = title
        .AllowMultiSelect = False
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
Failed:
End Function

'------------------------------------------------------------------------------
' DeleteFileSafe
' Deletes a file if it is there, clearing the read only flag first. True when
' the file is gone afterwards.
'------------------------------------------------------------------------------
Public Function DeleteFileSafe(ByVal filePath As String) As Boolean
    If Not FileExists(filePath) Then DeleteFileSafe = True: Exit Function

    On Error Resume Next
    SetAttr filePath, vbNormal
    Kill filePath
    On Error GoTo 0
    DeleteFileSafe = Not FileExists(filePath)
End Function

'------------------------------------------------------------------------------
' FileSizeBytes / FileModified
' 0 and 0 when the file is not there.
'------------------------------------------------------------------------------
Public Function FileSizeBytes(ByVal filePath As String) As Double
    If Not FileExists(filePath) Then Exit Function
    On Error Resume Next
    FileSizeBytes = CreateObject("Scripting.FileSystemObject").GetFile(filePath).Size
    On Error GoTo 0
End Function

Public Function FileModified(ByVal filePath As String) As Date
    If Not FileExists(filePath) Then Exit Function
    On Error Resume Next
    FileModified = CreateObject("Scripting.FileSystemObject").GetFile(filePath).DateLastModified
    On Error GoTo 0
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Sub CollectFiles(ByVal folder As Object, ByVal pattern As String, _
                         ByVal includeSubfolders As Boolean, ByRef result As Collection)
    Dim item As Object
    Dim child As Object

    On Error Resume Next
    For Each item In folder.Files
        If LCase$(item.name) Like LCase$(pattern) Then result.Add item.Path
    Next item

    If includeSubfolders Then
        For Each child In folder.SubFolders
            CollectFiles child, pattern, includeSubfolders, result
        Next child
    End If
    On Error GoTo 0
End Sub

' One CSV field: quoted when it has to be, with embedded quotes doubled.
Private Function CsvField(ByVal value As Variant, ByVal delimiter As String) As String
    Dim s As String

    If IsError(value) Then
        s = "#ERROR"
    ElseIf IsNull(value) Or IsEmpty(value) Then
        s = ""
    ElseIf VarType(value) = vbDate Then
        If Int(CDate(value)) = CDate(value) Then
            s = Format$(value, "yyyy-mm-dd")
        Else
            s = Format$(value, "yyyy-mm-dd hh:nn:ss")
        End If
    ElseIf VarType(value) = vbBoolean Then
        s = IIf(CBool(value), "TRUE", "FALSE")
    ElseIf VarType(value) = vbString Then
        s = CStr(value)                                     ' text keeps any leading zeros
    ElseIf IsNumeric(value) Then
        s = Trim$(Str$(value))                              ' no thousands separator, always a dot
    Else
        s = CStr(value)
    End If

    If InStr(s, delimiter) > 0 Or InStr(s, """") > 0 Or InStr(s, vbCr) > 0 _
       Or InStr(s, vbLf) > 0 Or s <> Trim$(s) Then
        s = """" & Replace$(s, """", """""") & """"
    End If

    CsvField = s
End Function
