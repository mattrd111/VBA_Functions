Attribute VB_Name = "modDoctorStyles"
'==============================================================================
' modDoctorStyles - cell style clean-up
'------------------------------------------------------------------------------
' Workbooks that have had sheets copied into them for years accumulate tens of
' thousands of cell styles: "Normal 2", "Normal 2 2", "Comma 3 4 2" and so on.
' They bloat the file, slow every open and eventually produce "Too many
' different cell formats".
'
' Two modes:
'   Duplicates  Only styles whose name is a built-in name with numbers stuck on
'               the end, plus "Style 1" style leftovers. Very low risk.
'   All custom  Every style Excel did not create. Anything actually using one of
'               them falls back to Normal formatting.
'
' Direct formatting applied to cells is never affected either way.
'
' Depends on: modDoctorCommon, modDictionary, modApp (from \src)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' ShowStylesReport   (menu action)
'------------------------------------------------------------------------------
Public Sub ShowStylesReport()
    Dim wb As Workbook
    Dim duplicates As Collection, others As Collection
    Dim builtInCount As Long
    Dim rows As Collection
    Dim item As Variant

    Set wb = TargetWorkbook(False)
    If wb Is Nothing Then Exit Sub

    StartTimer
    FastMode True
    builtInCount = ClassifyStyles(wb, duplicates, others)
    FastMode False

    Set rows = NewReport()
    ReportRow rows, DOCTOR_NAME & " - Cell styles"
    ReportRow rows, "Workbook", wb.Name
    ReportRow rows, "Scanned in", ElapsedText

    ReportHeading rows, "Summary"
    ReportRow rows, "Total styles", wb.Styles.count
    ReportRow rows, "Built in", builtInCount, "Excel's own - never deleted"
    ReportRow rows, "Duplicates of built-ins", duplicates.count, "Safe to delete"
    ReportRow rows, "Other custom styles", others.count, "Deleted only in 'all custom' mode"

    If duplicates.count > 0 Then
        ReportHeading rows, "Duplicates of built-in styles (" & duplicates.count & ")"
        For Each item In duplicates
            ReportRow rows, CStr(item)
        Next item
    End If

    If others.count > 0 Then
        ReportHeading rows, "Other custom styles (" & others.count & ")"
        For Each item In others
            ReportRow rows, CStr(item)
        Next item
    End If

    ShowReport rows, "Cell styles"
End Sub

'------------------------------------------------------------------------------
' CleanStyles   (menu action)
'------------------------------------------------------------------------------
Public Sub CleanStyles()
    Dim wb As Workbook
    Dim duplicates As Collection, others As Collection
    Dim builtInCount As Long
    Dim answer As VbMsgBoxResult
    Dim deleted As Long, failed As Long
    Dim cancelled As Boolean

    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub

    StartTimer
    FastMode True
    builtInCount = ClassifyStyles(wb, duplicates, others)
    FastMode False

    If duplicates.count + others.count = 0 Then
        MsgBox "'" & wb.Name & "' has " & wb.Styles.count & " styles and all of them are " & _
               "Excel's own. Nothing to clean.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    answer = MsgBox("'" & wb.Name & "' has " & Format$(wb.Styles.count, "#,##0") & " cell styles:" & _
                    vbNewLine & vbNewLine & _
                    "  " & Format$(builtInCount, "#,##0") & " built in (kept)" & vbNewLine & _
                    "  " & Format$(duplicates.count, "#,##0") & " duplicates of built-ins, such as 'Normal 2'" & vbNewLine & _
                    "  " & Format$(others.count, "#,##0") & " other custom styles" & vbNewLine & vbNewLine & _
                    "Yes     delete the " & Format$(duplicates.count, "#,##0") & " duplicates only (low risk)" & vbNewLine & _
                    "No      delete all " & Format$(duplicates.count + others.count, "#,##0") & " custom styles" & vbNewLine & _
                    "Cancel  stop" & vbNewLine & vbNewLine & _
                    "Cells using a deleted style fall back to Normal formatting. Formatting " & _
                    "applied directly to cells is not affected. This cannot be undone with Ctrl+Z.", _
                    vbYesNoCancel + vbExclamation, DOCTOR_NAME)

    If answer = vbCancel Then Exit Sub
    If Not OfferBackup(wb) Then Exit Sub

    StartTimer
    FastMode True
    deleted = DeleteStyles(wb, duplicates, failed, cancelled)
    If answer = vbNo And Not cancelled Then
        deleted = deleted + DeleteStyles(wb, others, failed, cancelled)
    End If
    FastMode False
    ClearStatusBar

    MsgBox Format$(deleted, "#,##0") & " style(s) deleted in " & ElapsedText & "." & vbNewLine & _
           Format$(wb.Styles.count, "#,##0") & " remain." & _
           IIf(failed > 0, vbNewLine & Format$(failed, "#,##0") & " could not be deleted - Excel is using them.", "") & _
           IIf(cancelled, vbNewLine & vbNewLine & "Stopped early because you pressed Esc.", "") & _
           vbNewLine & vbNewLine & "Save the workbook to keep the change.", vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' ClassifyStyles
' Sorts every style into built-in, duplicate-of-built-in, or other custom.
' Returns the built-in count and fills the two collections with style names.
'------------------------------------------------------------------------------
Public Function ClassifyStyles(ByVal wb As Workbook, ByRef duplicates As Collection, _
                               ByRef others As Collection) As Long
    Dim builtInNames As Object
    Dim styleNames As Collection
    Dim styleName As Variant
    Dim st As Style
    Dim i As Long
    Dim total As Long
    Dim builtInCount As Long

    Set duplicates = New Collection
    Set others = New Collection
    Set builtInNames = NewDictionary(True)
    Set styleNames = New Collection
    If wb Is Nothing Then Exit Function

    total = wb.Styles.count

    ' One pass: read every style once, because each read is a COM call and there
    ' can be tens of thousands of them.
    For i = 1 To total
        If i Mod 250 = 0 Then ShowProgress i, total, "Reading styles"

        On Error Resume Next
        Set st = wb.Styles(i)
        If Err.Number <> 0 Then
            Err.Clear
        Else
            If st.builtIn Then
                builtInCount = builtInCount + 1
                If Not builtInNames.Exists(UCase$(st.Name)) Then builtInNames.Add UCase$(st.Name), True
            Else
                styleNames.Add st.Name
            End If
        End If
        On Error GoTo 0
    Next i

    For Each styleName In styleNames
        If LooksLikeDuplicate(CStr(styleName), builtInNames) Then
            duplicates.Add CStr(styleName)
        Else
            others.Add CStr(styleName)
        End If
    Next styleName

    ClearStatusBar
    ClassifyStyles = builtInCount
End Function

'------------------------------------------------------------------------------
' DeleteStyles
' Deletes styles by name. Esc stops the run - useful when a workbook turns out
' to have 60,000 of them.
'------------------------------------------------------------------------------
Public Function DeleteStyles(ByVal wb As Workbook, ByVal styleNames As Collection, _
                             ByRef failed As Long, ByRef cancelled As Boolean) As Long
    Dim i As Long
    Dim deleted As Long
    Dim total As Long

    If wb Is Nothing Or styleNames Is Nothing Then Exit Function
    total = styleNames.count
    If total = 0 Then Exit Function

    Application.EnableCancelKey = xlErrorHandler

    For i = 1 To total
        If i Mod 100 = 0 Then ShowProgress i, total, "Deleting styles (Esc to stop)"

        On Error Resume Next
        wb.Styles(CStr(styleNames(i))).Delete

        If Err.Number = 18 Then                     ' Esc
            Err.Clear
            cancelled = True
            Exit For
        ElseIf Err.Number <> 0 Then
            failed = failed + 1
            Err.Clear
        Else
            deleted = deleted + 1
        End If
        On Error GoTo 0
    Next i

    Application.EnableCancelKey = xlInterrupt
    DeleteStyles = deleted
End Function

'==============================================================================
' Private helpers
'==============================================================================
' "Normal 2", "Comma 2 3", "Style 17" - a built-in name with numbers stuck on.
Private Function LooksLikeDuplicate(ByVal styleName As String, ByVal builtInNames As Object) As Boolean
    Dim stem As String

    stem = StripTrailingNumbers(styleName)
    If Len(stem) = 0 Then Exit Function
    If StrComp(stem, styleName, vbTextCompare) = 0 Then Exit Function   ' no numbers were stripped

    If builtInNames.Exists(UCase$(stem)) Then LooksLikeDuplicate = True
    If StrComp(stem, "Style", vbTextCompare) = 0 Then LooksLikeDuplicate = True
End Function

Private Function StripTrailingNumbers(ByVal text As String) As String
    Dim s As String
    Dim p As Long

    s = Trim$(text)
    Do
        p = InStrRev(s, " ")
        If p = 0 Then Exit Do
        If Not IsAllDigits(Mid$(s, p + 1)) Then Exit Do
        s = Trim$(Left$(s, p - 1))
    Loop

    StripTrailingNumbers = s
End Function

Private Function IsAllDigits(ByVal text As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function
