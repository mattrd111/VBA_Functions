Attribute VB_Name = "modWorkbook"
'==============================================================================
' modWorkbook - workbook and worksheet helpers
'------------------------------------------------------------------------------
' Creating, finding, clearing and deleting sheets safely, opening and locating
' workbooks, dumping data to a new sheet, exporting to PDF.
'
' Standalone: no dependency on the other modules in this repository.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' SheetExists
' True when a sheet of that name exists in the workbook (the active workbook
' when none is given). Charts and dialog sheets count.
'------------------------------------------------------------------------------
Public Function SheetExists(ByVal sheetName As String, Optional ByVal wb As Workbook) As Boolean
    Dim sh As Object

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function

    On Error Resume Next
    Set sh = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sh Is Nothing
End Function

'------------------------------------------------------------------------------
' GetSheet
' The worksheet of that name, or Nothing - no error, so you can test the result
' instead of wrapping the call in an error handler.
'------------------------------------------------------------------------------
Public Function GetSheet(ByVal sheetName As String, Optional ByVal wb As Workbook) As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function

    On Error Resume Next
    Set GetSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' GetOrCreateSheet
' Returns the named worksheet, creating it at the end of the workbook if it is
' not there yet. Set clearIfExists to start from an empty sheet either way.
'------------------------------------------------------------------------------
Public Function GetOrCreateSheet(ByVal sheetName As String, Optional ByVal wb As Workbook, _
                                 Optional ByVal clearIfExists As Boolean = False) As Worksheet
    Dim ws As Worksheet

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function

    sheetName = SafeSheetName(sheetName)
    Set ws = GetSheet(sheetName, wb)

    If ws Is Nothing Then
        If wb.Worksheets.Count = 0 Then
            Set ws = wb.Worksheets.Add
        Else
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        End If

        On Error Resume Next
        ws.Name = sheetName
        If Err.Number <> 0 Then                 ' a chart sheet already has the name
            Err.Clear
            ws.Name = UniqueSheetName(sheetName, wb)
        End If
        On Error GoTo 0
    ElseIf clearIfExists Then
        ClearSheet ws
    End If

    Set GetOrCreateSheet = ws
End Function

'------------------------------------------------------------------------------
' ClearSheet
' Empties a sheet: contents, formats, shapes, conditional formatting, filters
' and merged cells. Leaves the sheet itself in place.
'------------------------------------------------------------------------------
Public Sub ClearSheet(ByVal ws As Worksheet, Optional ByVal keepFormats As Boolean = False)
    Dim shp As Shape
    Dim i As Long

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.UnMerge
    ws.Cells.FormatConditions.Delete
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        shp.Delete
    Next i
    On Error GoTo 0

    If keepFormats Then
        ws.Cells.ClearContents
    Else
        ws.Cells.Clear
        ws.Cells.ColumnWidth = ws.StandardWidth
        ws.Cells.RowHeight = ws.StandardHeight
    End If
End Sub

'------------------------------------------------------------------------------
' DeleteSheet
' Deletes a sheet without the confirmation prompt and without raising when it
' is not there. Refuses to delete the last visible sheet, which Excel forbids.
'------------------------------------------------------------------------------
Public Sub DeleteSheet(ByVal sheetName As String, Optional ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim alertsWereOn As Boolean

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Set ws = GetSheet(sheetName, wb)
    If ws Is Nothing Then Exit Sub
    If VisibleSheetCount(wb) <= 1 And ws.Visible = xlSheetVisible Then Exit Sub

    alertsWereOn = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ws.Delete
    On Error GoTo 0
    Application.DisplayAlerts = alertsWereOn
End Sub

'------------------------------------------------------------------------------
' SafeSheetName
' Turns any string into something Excel will accept as a sheet name: strips
' : \ / ? * [ ], trims to 31 characters and never returns an empty name.
'------------------------------------------------------------------------------
Public Function SafeSheetName(ByVal proposedName As String, _
                              Optional ByVal replacement As String = "-") As String
    Dim bad As Variant
    Dim item As Variant
    Dim s As String

    s = Trim$(proposedName)
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For Each item In bad
        s = Replace$(s, CStr(item), replacement)
    Next item

    s = Replace$(s, vbTab, " ")
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop

    s = Trim$(s)
    If Len(s) > 31 Then s = Trim$(Left$(s, 31))
    Do While Left$(s, 1) = "'"
        s = Mid$(s, 2)
    Loop
    Do While Right$(s, 1) = "'"
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then s = "Sheet"
    If StrComp(s, "History", vbTextCompare) = 0 Then s = "History_"

    SafeSheetName = s
End Function

'------------------------------------------------------------------------------
' UniqueSheetName
' A sheet name that is not in use yet, appending (2), (3) ... when needed.
'------------------------------------------------------------------------------
Public Function UniqueSheetName(ByVal baseName As String, Optional ByVal wb As Workbook) As String
    Dim candidate As String
    Dim suffix As String
    Dim n As Long
    Dim root As String

    If wb Is Nothing Then Set wb = ActiveWorkbook
    root = SafeSheetName(baseName)
    candidate = root

    Do While SheetExists(candidate, wb)
        n = n + 1
        suffix = " (" & CStr(n + 1) & ")"
        If Len(root) + Len(suffix) > 31 Then
            candidate = Left$(root, 31 - Len(suffix)) & suffix
        Else
            candidate = root & suffix
        End If
    Loop

    UniqueSheetName = candidate
End Function

'------------------------------------------------------------------------------
' SheetNames
' Worksheet names as a 1-based array. Set visibleOnly to skip hidden sheets.
'------------------------------------------------------------------------------
Public Function SheetNames(Optional ByVal wb As Workbook, _
                           Optional ByVal visibleOnly As Boolean = False) As Variant
    Dim names() As String
    Dim ws As Worksheet
    Dim n As Long

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then SheetNames = Array(): Exit Function

    ReDim names(1 To wb.Worksheets.Count)
    For Each ws In wb.Worksheets
        If Not visibleOnly Or ws.Visible = xlSheetVisible Then
            n = n + 1
            names(n) = ws.Name
        End If
    Next ws

    If n = 0 Then SheetNames = Array(): Exit Function
    ReDim Preserve names(1 To n)
    SheetNames = names
End Function

'------------------------------------------------------------------------------
' WorkbookIsOpen / GetOpenWorkbook
' Pass either a file name ("Budget.xlsx") or a full path.
'------------------------------------------------------------------------------
Public Function WorkbookIsOpen(ByVal nameOrPath As String) As Boolean
    WorkbookIsOpen = Not GetOpenWorkbook(nameOrPath) Is Nothing
End Function

Public Function GetOpenWorkbook(ByVal nameOrPath As String) As Workbook
    Dim wb As Workbook
    Dim target As String

    target = nameOrPath
    If InStr(target, "\") > 0 Or InStr(target, "/") > 0 Then
        target = Mid$(target, InStrRev(Replace$(target, "/", "\"), "\") + 1)
    End If
    If Len(target) = 0 Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, target, vbTextCompare) = 0 Then
            Set GetOpenWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

'------------------------------------------------------------------------------
' OpenWorkbook
' Opens a workbook, or hands back the already open one. Returns Nothing when
' the file is missing or cannot be opened.
'------------------------------------------------------------------------------
Public Function OpenWorkbook(ByVal filePath As String, _
                             Optional ByVal readOnlyMode As Boolean = True, _
                             Optional ByVal updateLinks As Boolean = False) As Workbook
    Dim wb As Workbook

    Set wb = GetOpenWorkbook(filePath)
    If Not wb Is Nothing Then Set OpenWorkbook = wb: Exit Function
    If Len(Dir$(filePath)) = 0 Then Exit Function

    On Error Resume Next
    Set OpenWorkbook = Application.Workbooks.Open(Filename:=filePath, _
                                                  UpdateLinks:=IIf(updateLinks, 3, 0), _
                                                  ReadOnly:=readOnlyMode)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' DumpToSheet
' Writes a 2D array to a fresh sheet in one operation, with a bold frozen
' header row, and returns the sheet. The quickest way to get results in front
' of someone.
'
'   Set ws = DumpToSheet(results, "Exceptions")
'------------------------------------------------------------------------------
Public Function DumpToSheet(ByVal data As Variant, ByVal sheetName As String, _
                            Optional ByVal wb As Workbook, _
                            Optional ByVal hasHeaderRow As Boolean = True) As Worksheet
    Dim ws As Worksheet
    Dim block As Variant
    Dim rows As Long, cols As Long
    Dim i As Long

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Set wb = Application.Workbooks.Add

    Set ws = GetOrCreateSheet(sheetName, wb, True)
    Set DumpToSheet = ws
    If Not IsArray(data) Then Exit Function

    On Error GoTo Done
    If Dimensions(data) = 1 Then
        ReDim block(1 To UBound(data) - LBound(data) + 1, 1 To 1)
        For i = 1 To UBound(block, 1)
            block(i, 1) = data(LBound(data) + i - 1)
        Next i
    Else
        block = data
    End If

    rows = UBound(block, 1) - LBound(block, 1) + 1
    cols = UBound(block, 2) - LBound(block, 2) + 1
    If rows < 1 Or cols < 1 Then Exit Function

    ws.Range("A1").Resize(rows, cols).value = block

    If hasHeaderRow Then
        With ws.Range("A1").Resize(1, cols)
            .Font.Bold = True
            .Interior.Color = RGB(242, 242, 242)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
    End If

    ws.Columns.AutoFit
Done:
End Function

'------------------------------------------------------------------------------
' ExportToPdf
' Exports a workbook, sheet or range to PDF. Returns True on success.
'
'   ExportToPdf ws, "C:\Reports\Summary.pdf"
'------------------------------------------------------------------------------
Public Function ExportToPdf(ByVal target As Object, ByVal filePath As String, _
                            Optional ByVal openAfter As Boolean = False) As Boolean
    If target Is Nothing Then Exit Function
    If Len(filePath) = 0 Then Exit Function

    On Error GoTo Failed
    target.ExportAsFixedFormat Type:=xlTypePDF, _
                               Filename:=filePath, _
                               Quality:=xlQualityStandard, _
                               IncludeDocProperties:=True, _
                               IgnorePrintAreas:=False, _
                               OpenAfterPublish:=openAfter
    ExportToPdf = True
    Exit Function
Failed:
    ExportToPdf = False
End Function

'------------------------------------------------------------------------------
' RefreshAllAndWait
' Refreshes queries, connections and pivot caches and does not return until
' they have finished - unlike a bare RefreshAll, which carries straight on.
'------------------------------------------------------------------------------
Public Sub RefreshAllAndWait(Optional ByVal wb As Workbook)
    Dim conn As Object
    Dim ws As Worksheet
    Dim qt As Object

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    For Each conn In wb.Connections
        Select Case conn.Type
            Case 1                                  ' xlConnectionTypeOLEDB
                conn.OLEDBConnection.BackgroundQuery = False
            Case 2                                  ' xlConnectionTypeODBC
                conn.ODBCConnection.BackgroundQuery = False
        End Select
    Next conn
    For Each ws In wb.Worksheets
        For Each qt In ws.QueryTables
            qt.BackgroundQuery = False
        Next qt
    Next ws
    On Error GoTo 0

    wb.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    DoEvents
End Sub

'------------------------------------------------------------------------------
' ProtectAllSheets / UnprotectAllSheets
' Blank password means no password.
'------------------------------------------------------------------------------
Public Sub ProtectAllSheets(Optional ByVal wb As Workbook, Optional ByVal password As String = "")
    Dim ws As Worksheet

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    For Each ws In wb.Worksheets
        If Len(password) = 0 Then
            ws.Protect UserInterfaceOnly:=True
        Else
            ws.Protect password:=password, UserInterfaceOnly:=True
        End If
    Next ws
    On Error GoTo 0
End Sub

Public Sub UnprotectAllSheets(Optional ByVal wb As Workbook, Optional ByVal password As String = "")
    Dim ws As Worksheet

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    For Each ws In wb.Worksheets
        If Len(password) = 0 Then
            ws.Unprotect
        Else
            ws.Unprotect password:=password
        End If
    Next ws
    On Error GoTo 0
End Sub

'==============================================================================
' Private helpers
'==============================================================================
Private Function Dimensions(ByVal arr As Variant) As Long
    Dim d As Long
    Dim test As Long

    If Not IsArray(arr) Then Exit Function
    On Error GoTo Done
    Do
        d = d + 1
        test = LBound(arr, d)
    Loop While d < 60
Done:
    Dimensions = d - 1
End Function

Private Function VisibleSheetCount(ByVal wb As Workbook) As Long
    Dim sh As Object
    Dim total As Long

    For Each sh In wb.Sheets
        If sh.Visible = xlSheetVisible Then total = total + 1
    Next sh

    VisibleSheetCount = total
End Function
