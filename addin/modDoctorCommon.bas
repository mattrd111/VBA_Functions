Attribute VB_Name = "modDoctorCommon"
'==============================================================================
' modDoctorCommon - shared plumbing for the Workbook Doctor add-in
'------------------------------------------------------------------------------
' Picking the workbook to work on, confirmations, backups, and the report
' builder every tool writes its findings into.
'
' Depends on: modFile, modWorkbook, modRange, modApp (from \src)
'==============================================================================
Option Explicit

Public Const DOCTOR_NAME As String = "Workbook Doctor"
Public Const DOCTOR_VERSION As String = "1.0"

'------------------------------------------------------------------------------
' TargetWorkbook
' The workbook the tools act on - the active one, never the add-in itself.
' Returns Nothing and explains why when there is nothing safe to work on.
'------------------------------------------------------------------------------
Public Function TargetWorkbook(Optional ByVal needsWriteAccess As Boolean = True) As Workbook
    Dim wb As Workbook

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Open a workbook first.", vbInformation, DOCTOR_NAME
        Exit Function
    End If

    If wb Is ThisWorkbook Then
        MsgBox DOCTOR_NAME & " will not operate on itself." & vbNewLine & vbNewLine & _
               "Switch to the workbook you want to clean and try again.", vbInformation, DOCTOR_NAME
        Exit Function
    End If

    If wb.IsAddin Then
        MsgBox "The active workbook is an add-in. Switch to an ordinary workbook first.", _
               vbInformation, DOCTOR_NAME
        Exit Function
    End If

    If needsWriteAccess And wb.ReadOnly Then
        MsgBox "'" & wb.Name & "' is open read only, so nothing can be changed." & vbNewLine & _
               vbNewLine & "Reopen it with write access first.", vbExclamation, DOCTOR_NAME
        Exit Function
    End If

    Set TargetWorkbook = wb
End Function

'------------------------------------------------------------------------------
' Confirm
' Yes/No question with the standard "no undo" warning attached.
'------------------------------------------------------------------------------
Public Function Confirm(ByVal message As String, Optional ByVal warnNoUndo As Boolean = True) As Boolean
    Dim text As String

    text = message
    If warnNoUndo Then
        text = text & vbNewLine & vbNewLine & _
               "Ctrl+Z will NOT undo this. Take a backup first if you are unsure."
    End If

    Confirm = (MsgBox(text, vbYesNo + vbExclamation + vbDefaultButton2, DOCTOR_NAME) = vbYes)
End Function

'------------------------------------------------------------------------------
' BackupWorkbook
' Saves a timestamped copy beside the original and returns its path, or "" if
' the workbook has never been saved anywhere.
'------------------------------------------------------------------------------
Public Function BackupWorkbook(ByVal wb As Workbook) As String
    Dim target As String
    Dim extension As String

    If wb Is Nothing Then Exit Function
    If Len(wb.Path) = 0 Then Exit Function

    extension = FileExtension(wb.Name)
    target = JoinPath(wb.Path, BaseName(wb.Name) & " (backup " & Format$(Now, "yyyy-mm-dd hhnn") & ")")
    If Len(extension) > 0 Then target = target & "." & extension
    target = UniqueFilePath(target)

    On Error GoTo Failed
    wb.SaveCopyAs target
    BackupWorkbook = target
    Exit Function
Failed:
    BackupWorkbook = ""
End Function

'------------------------------------------------------------------------------
' OfferBackup
' Asks whether to back up before a destructive run. False means the user chose
' Cancel and the caller should stop.
'------------------------------------------------------------------------------
Public Function OfferBackup(ByVal wb As Workbook) As Boolean
    Dim answer As VbMsgBoxResult
    Dim path As String

    If wb Is Nothing Then Exit Function

    If Len(wb.Path) = 0 Then
        OfferBackup = Confirm("'" & wb.Name & "' has never been saved, so no backup copy " & _
                              "can be made." & vbNewLine & vbNewLine & "Carry on anyway?", False)
        Exit Function
    End If

    answer = MsgBox("Save a backup copy of '" & wb.Name & "' first?" & vbNewLine & vbNewLine & _
                    "Yes  - copy it, then carry on" & vbNewLine & _
                    "No   - carry on without a backup" & vbNewLine & _
                    "Cancel - stop", vbYesNoCancel + vbQuestion, DOCTOR_NAME)

    Select Case answer
        Case vbYes
            path = BackupWorkbook(wb)
            If Len(path) = 0 Then
                OfferBackup = Confirm("The backup could not be written." & vbNewLine & vbNewLine & _
                                      "Carry on anyway?", False)
            Else
                MsgBox "Backup saved as:" & vbNewLine & vbNewLine & path, vbInformation, DOCTOR_NAME
                OfferBackup = True
            End If
        Case vbNo
            OfferBackup = True
        Case Else
            OfferBackup = False
    End Select
End Function

'==============================================================================
' Report builder
'------------------------------------------------------------------------------
' Tools collect rows, then hand them to ShowReport, which writes them to a new
' workbook. Rows with only a first cell filled are treated as headings.
'==============================================================================
Public Function NewReport() As Collection
    Set NewReport = New Collection
End Function

Public Sub ReportRow(ByVal rows As Collection, ParamArray cells() As Variant)
    Dim copyOfCells As Variant

    If rows Is Nothing Then Exit Sub
    copyOfCells = cells
    rows.Add copyOfCells
End Sub

Public Sub ReportHeading(ByVal rows As Collection, ByVal title As String)
    If rows Is Nothing Then Exit Sub
    If rows.count > 0 Then rows.Add Array("")
    rows.Add Array(title)
End Sub

'------------------------------------------------------------------------------
' ShowReport
' Writes the collected rows to a new workbook and returns the sheet.
'------------------------------------------------------------------------------
Public Function ShowReport(ByVal rows As Collection, ByVal reportTitle As String) As Worksheet
    Dim data As Variant
    Dim item As Variant
    Dim ws As Worksheet
    Dim book As Workbook
    Dim r As Long, c As Long
    Dim columnCount As Long

    If rows Is Nothing Then Exit Function
    If rows.count = 0 Then Exit Function

    For Each item In rows
        If IsArray(item) Then
            If UBound(item) - LBound(item) + 1 > columnCount Then
                columnCount = UBound(item) - LBound(item) + 1
            End If
        End If
    Next item
    If columnCount = 0 Then columnCount = 1

    ReDim data(1 To rows.count, 1 To columnCount)
    For r = 1 To rows.count
        item = rows(r)
        If IsArray(item) Then
            For c = 1 To UBound(item) - LBound(item) + 1
                data(r, c) = item(LBound(item) + c - 1)
            Next c
        End If
    Next r

    Set book = Application.Workbooks.Add
    Set ws = DumpToSheet(data, Left$(reportTitle, 31), book, False)
    If ws Is Nothing Then Exit Function

    RemoveSpareSheets book, ws
    FormatReport ws, columnCount
    Set ShowReport = ws
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Sub FormatReport(ByVal ws As Worksheet, ByVal columnCount As Long)
    Dim r As Long
    Dim lastR As Long

    lastR = LastRow(ws)
    If lastR = 0 Then Exit Sub

    With ws.Range("A1").Resize(lastR, columnCount)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .VerticalAlignment = xlTop
    End With

    ' A row with only the first cell filled is a heading.
    For r = 1 To lastR
        If IsHeadingRow(ws, r, columnCount) Then
            With ws.Cells(r, 1)
                .Font.Bold = True
                .Font.Size = 11
            End With
        End If
    Next r

    ws.Range("A1").Font.Size = 14
    AutoFitColumns ws, 55, 10
    FreezeHeader ws, 1

    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Private Function IsHeadingRow(ByVal ws As Worksheet, ByVal r As Long, ByVal columnCount As Long) As Boolean
    If Len(CStr(ws.Cells(r, 1).value)) = 0 Then Exit Function
    If columnCount <= 1 Then IsHeadingRow = True: Exit Function

    IsHeadingRow = (Application.CountA(ws.Cells(r, 2).Resize(1, columnCount - 1)) = 0)
End Function

' A new workbook arrives with a blank sheet the report does not need.
Private Sub RemoveSpareSheets(ByVal book As Workbook, ByVal keep As Worksheet)
    Dim ws As Worksheet
    Dim alertsWereOn As Boolean

    alertsWereOn = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    For Each ws In book.Worksheets
        If Not ws Is keep Then
            If Application.CountA(ws.Cells) = 0 And ws.Shapes.count = 0 Then
                If book.Worksheets.count > 1 Then ws.Delete
            End If
        End If
    Next ws
    On Error GoTo 0

    Application.DisplayAlerts = alertsWereOn
End Sub
