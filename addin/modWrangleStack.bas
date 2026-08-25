Attribute VB_Name = "modWrangleStack"
'==============================================================================
' modWrangleStack - putting many extracts into one table
'------------------------------------------------------------------------------
' The first job in every data room: thirty-six monthly extracts, or twelve
' entity tabs, that need to become one table before anything can be analysed.
' They never have their columns in the same order, one of them is missing a
' column entirely, and two of them spell a header differently.
'
' This matches columns by header NAME rather than position, keeps a column of
' its own for wherever each row came from, and reports which source was missing
' what - which is usually the finding, not the footnote.
'
' Depends on: modDoctorCommon, modFile, modWorkbook, modRange, modDictionary,
'             modString, modApp (from \src)
'==============================================================================
Option Explicit

Private Const MAX_TOTAL_CELLS As Double = 8000000#

'------------------------------------------------------------------------------
' StackSelectedSheets   (menu action)
' Ctrl+click the sheet tabs you want, then run this. No dialog to drive.
'------------------------------------------------------------------------------
Public Sub StackSelectedSheets()
    Dim wb As Workbook
    Dim selected As Object
    Dim sources As Collection
    Dim sh As Object
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim source As Variant

    Set wb = TargetWorkbook(False)
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    Set selected = ActiveWindow.SelectedSheets
    On Error GoTo 0
    If selected Is Nothing Then Exit Sub

    If selected.count < 2 Then
        MsgBox "Ctrl+click the sheet tabs you want to stack, then run this again." & vbNewLine & _
               vbNewLine & "Only '" & ActiveSheet.Name & "' is selected.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    headerRow = AskForNumber("Which row holds the headers on each sheet?", 1)
    If headerRow < 1 Then Exit Sub

    StartTimer
    FastMode True

    Set sources = New Collection
    For Each sh In selected
        If TypeOf sh Is Worksheet Then
            Set ws = sh
            source = CollectSource(ws, headerRow, ws.Name, ws.Name)
            If IsArray(source) Then sources.Add source
        End If
    Next sh

    FastMode False
    BuildStack sources, "Stacked sheets", wb.Name
End Sub

'------------------------------------------------------------------------------
' StackFilesInFolder   (menu action)
' Every workbook in a folder, one sheet from each.
'------------------------------------------------------------------------------
Public Sub StackFilesInFolder()
    Dim folderPath As String
    Dim pattern As String
    Dim sheetName As String
    Dim headerRow As Long
    Dim files As Collection
    Dim sources As Collection
    Dim path As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim source As Variant
    Dim weOpenedIt As Boolean
    Dim done As Long
    Dim cancelled As Boolean

    folderPath = PickFolder("Pick the folder holding the extracts")
    If Len(folderPath) = 0 Then Exit Sub

    pattern = AskForText("Which files? (* and ? are allowed)", "*.xls*", cancelled)
    If cancelled Or Len(pattern) = 0 Then Exit Sub

    sheetName = AskForText("Which sheet in each file?" & vbNewLine & vbNewLine & _
                           "Leave this blank to take the first sheet with data on it.", "", cancelled)
    If cancelled Then Exit Sub

    headerRow = AskForNumber("Which row holds the headers?", 1)
    If headerRow < 1 Then Exit Sub

    Set files = ListFiles(folderPath, pattern, False)
    If files.count = 0 Then
        MsgBox "No files matching " & pattern & " in" & vbNewLine & folderPath, _
               vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    If MsgBox(files.count & " file(s) found. Open each one and stack them?", _
              vbYesNo + vbQuestion, DOCTOR_NAME) <> vbYes Then Exit Sub

    StartTimer
    FastMode True

    Set sources = New Collection
    For Each path In files
        done = done + 1
        ShowProgress done, files.count, "Reading " & FileNameOnly(CStr(path))

        weOpenedIt = (GetOpenWorkbook(CStr(path)) Is Nothing)
        Set wb = OpenWorkbook(CStr(path), True)

        If Not wb Is Nothing Then
            Set ws = PickSheet(wb, sheetName)
            If Not ws Is Nothing Then
                source = CollectSource(ws, headerRow, FileNameOnly(CStr(path)), ws.Name)
                If IsArray(source) Then sources.Add source
            End If
            If weOpenedIt Then wb.Close SaveChanges:=False
        End If
        Set wb = Nothing
    Next path

    FastMode False
    ClearStatusBar
    BuildStack sources, "Stacked files", folderPath
End Sub

'==============================================================================
' Private - collecting
'==============================================================================
' Array(sourceName, sheetName, headers(), data()) or Empty when there is
' nothing usable on the sheet.
Private Function CollectSource(ByVal ws As Worksheet, ByVal headerRow As Long, _
                               ByVal sourceName As String, ByVal sheetName As String) As Variant
    Dim lastR As Long, lastC As Long
    Dim headers() As String
    Dim data As Variant
    Dim c As Long

    If ws Is Nothing Then Exit Function

    lastC = LastColumn(ws, headerRow)
    lastR = LastRow(ws)
    If lastC < 1 Or lastR <= headerRow Then Exit Function

    ReDim headers(1 To lastC)
    For c = 1 To lastC
        headers(c) = CleanText(ws.Cells(headerRow, c).value)
    Next c

    data = ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastR, lastC)).value
    If Not IsArray(data) Then
        Dim single2D As Variant
        ReDim single2D(1 To 1, 1 To 1)
        single2D(1, 1) = data
        data = single2D
    End If

    CollectSource = Array(sourceName, sheetName, headers, data)
End Function

Private Function PickSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    If Len(sheetName) > 0 Then
        Set PickSheet = GetSheet(sheetName, wb)
        Exit Function
    End If

    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If Application.CountA(ws.Cells) > 0 Then
                Set PickSheet = ws
                Exit Function
            End If
        End If
    Next ws
End Function

'==============================================================================
' Private - building the output
'==============================================================================
Private Sub BuildStack(ByVal sources As Collection, ByVal title As String, ByVal origin As String)
    Dim masterIndex As Object
    Dim masterOrder As Collection
    Dim source As Variant
    Dim headers As Variant
    Dim data As Variant
    Dim key As String
    Dim c As Long, r As Long
    Dim totalRows As Long
    Dim out As Variant
    Dim outRow As Long, outCol As Long
    Dim book As Workbook
    Dim ws As Worksheet
    Dim rows As Collection

    If sources Is Nothing Then Exit Sub
    If sources.count = 0 Then
        MsgBox "Nothing to stack - none of the sources had a header row with data under it.", _
               vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    Set masterIndex = NewDictionary(True)
    Set masterOrder = New Collection

    For Each source In sources
        headers = source(2)
        data = source(3)
        totalRows = totalRows + (UBound(data, 1) - LBound(data, 1) + 1)

        For c = LBound(headers) To UBound(headers)
            key = CStr(headers(c))
            If Len(key) > 0 Then
                If Not masterIndex.Exists(key) Then
                    masterOrder.Add key
                    masterIndex.Add key, masterOrder.count
                End If
            End If
        Next c
    Next source

    If masterOrder.count = 0 Then
        MsgBox "No column headers were found. Check the header row number.", _
               vbExclamation, DOCTOR_NAME
        Exit Sub
    End If

    If CDbl(totalRows + 1) * CDbl(masterOrder.count + 2) > MAX_TOTAL_CELLS Then
        MsgBox "That would build a table of " & Format$(totalRows, "#,##0") & " rows by " & _
               (masterOrder.count + 2) & " columns, which is more than this tool will hold in " & _
               "memory at once." & vbNewLine & vbNewLine & _
               "Stack it in smaller batches, or use Power Query for something this size.", _
               vbExclamation, DOCTOR_NAME
        Exit Sub
    End If

    FastMode True

    ReDim out(1 To totalRows + 1, 1 To masterOrder.count + 2)
    out(1, 1) = "Source"
    out(1, 2) = "Source sheet"
    For c = 1 To masterOrder.count
        out(1, c + 2) = masterOrder(c)
    Next c

    outRow = 1
    For Each source In sources
        headers = source(2)
        data = source(3)

        For r = LBound(data, 1) To UBound(data, 1)
            outRow = outRow + 1
            out(outRow, 1) = source(0)
            out(outRow, 2) = source(1)

            For c = LBound(headers) To UBound(headers)
                key = CStr(headers(c))
                If Len(key) > 0 Then
                    outCol = masterIndex(key) + 2
                    out(outRow, outCol) = data(r, LBound(data, 2) + c - LBound(headers))
                End If
            Next c
        Next r
    Next source

    Set book = Application.Workbooks.Add
    Set ws = DumpToSheet(out, "Stacked", book, True)
    If Not ws Is Nothing Then
        FreezeHeader ws, 1
        AutoFitColumns ws, 40, 9
    End If

    Set rows = BuildStackReport(sources, masterOrder, masterIndex, totalRows, title, origin)
    ShowReport rows, "Stack report"

    FastMode False

    MsgBox Format$(totalRows, "#,##0") & " row(s) from " & sources.count & " source(s) stacked " & _
           "into " & masterOrder.count & " column(s), in " & ElapsedText & "." & vbNewLine & _
           vbNewLine & "Check the stack report for columns a source did not have.", _
           vbInformation, DOCTOR_NAME
End Sub

Private Function BuildStackReport(ByVal sources As Collection, ByVal masterOrder As Collection, _
                                  ByVal masterIndex As Object, ByVal totalRows As Long, _
                                  ByVal title As String, ByVal origin As String) As Collection
    Dim rows As Collection
    Dim source As Variant
    Dim headers As Variant
    Dim data As Variant
    Dim present As Object
    Dim missing As String
    Dim extra As String
    Dim c As Long
    Dim key As Variant
    Dim rowCount As Long

    Set rows = NewReport()
    ReportRow rows, DOCTOR_NAME & " - " & title
    ReportRow rows, "From", origin
    ReportRow rows, "Built", Format$(Now, "dd mmm yyyy hh:nn")
    ReportRow rows, "Sources", sources.count
    ReportRow rows, "Rows", Format$(totalRows, "#,##0")
    ReportRow rows, "Columns", masterOrder.count

    ReportHeading rows, "By source"
    ReportRow rows, "Source", "Sheet", "Rows", "Columns", "Columns it did not have"

    For Each source In sources
        headers = source(2)
        data = source(3)
        rowCount = UBound(data, 1) - LBound(data, 1) + 1

        Set present = NewDictionary(True)
        For c = LBound(headers) To UBound(headers)
            If Len(CStr(headers(c))) > 0 Then
                If Not present.Exists(CStr(headers(c))) Then present.Add CStr(headers(c)), True
            End If
        Next c

        missing = ""
        For Each key In masterOrder
            If Not present.Exists(CStr(key)) Then
                missing = missing & IIf(Len(missing) = 0, "", ", ") & CStr(key)
            End If
        Next key

        ReportRow rows, source(0), source(1), rowCount, present.count, _
                  IIf(Len(missing) = 0, "-", missing)
    Next source

    ReportHeading rows, "Columns, in the order they first appeared"
    ReportRow rows, "#", "Column"
    For c = 1 To masterOrder.count
        ReportRow rows, c, masterOrder(c)
    Next c

    Set BuildStackReport = rows
End Function

'==============================================================================
' Shared prompts
'==============================================================================
Public Function AskForNumber(ByVal prompt As String, ByVal defaultValue As Long) As Long
    Dim answer As Variant

    answer = Application.InputBox(prompt, DOCTOR_NAME, defaultValue, Type:=1)
    If VarType(answer) = vbBoolean Then Exit Function
    If Not IsNumeric(answer) Then Exit Function

    AskForNumber = CLng(answer)
End Function

' cancelled comes back True when the user pressed Cancel, which a blank answer
' on its own cannot tell you.
Public Function AskForText(ByVal prompt As String, ByVal defaultValue As String, _
                           Optional ByRef cancelled As Boolean) As String
    Dim answer As Variant

    cancelled = False
    answer = Application.InputBox(prompt, DOCTOR_NAME, defaultValue, Type:=2)

    If VarType(answer) = vbBoolean Then
        cancelled = True
        Exit Function
    End If

    AskForText = CStr(answer)
End Function
